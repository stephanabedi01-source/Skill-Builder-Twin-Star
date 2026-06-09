#!/usr/bin/env python3
"""
audit_workbook.py - scan a target workbook and emit a reviewable JSON map.

The map records, per tab, a tab-type guess (+confidence) and structural geometry
(label columns, data columns, header band); per row, a row-class guess
(+confidence); and per cell, provenance (hardcode / formula / cross-sheet link).
These are GUESSES. Claude reviews and corrects the map -- sampling real cells,
applying references/classification.md, fixing low-confidence calls -- BEFORE
apply_format.py consumes it. Nothing here changes the workbook.

Usage:
    python audit_workbook.py TARGET.xlsx [-o audit_map.json]

Design split: this script makes only the structural/keyword guesses a machine can
make. The semantic judgment (which tier a total is, banner vs underlined header,
working vs presentation) is intentionally left soft -- low confidence -- so the
reviewing model knows where to look.
"""

import argparse
import json
import re
import sys

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

import style_constants as sc

# ---------------------------------------------------------------------------
# Cell provenance
# ---------------------------------------------------------------------------
_ERROR_TOKENS = re.compile(r"#(REF!|NAME\?|DIV/0!|N/A|NULL!|NUM!|VALUE!|SPILL!|CALC!|GETTING_DATA)")
_SHEET_REF = re.compile(r"(?:'[^']+'|[A-Za-z_\\][\w.]*)!")  # SheetName! or 'Sheet Name'!


def provenance(cell):
    """hardcode | formula | cross_sheet | text | empty -- mechanical from content."""
    v = cell.value
    if v is None:
        return "empty"
    is_formula = cell.data_type == "f" or (isinstance(v, str) and v.startswith("="))
    if is_formula:
        s = v if isinstance(v, str) else getattr(v, "text", "=")
        body = _ERROR_TOKENS.sub("", s)
        return "cross_sheet" if _SHEET_REF.search(body) else "formula"
    if isinstance(v, str):
        return "text" if v.strip() else "empty"
    return "hardcode"  # number / date / bool constant -> blue candidate


def _is_number(cell):
    return cell.data_type in ("n",) and cell.value is not None and not isinstance(cell.value, bool)


# ---------------------------------------------------------------------------
# Sheet-level structure detection
# ---------------------------------------------------------------------------
def _used_bounds(ws, max_rows=800, max_cols=220):
    """(min_row, max_row, min_col, max_col), capped so huge tabs stay fast.
    Caps are generous enough to cover a 10-year monthly model plus its annuals
    block; tabs bigger than this are virtually always working/source tabs whose
    bodies we leave untouched anyway."""
    return (
        max(ws.min_row, 1),
        min(ws.max_row, max_rows),
        max(ws.min_column, 1),
        min(ws.max_column, max_cols),
    )


def _dominant_font(ws, bounds):
    counts = {}
    r0, r1, c0, c1 = bounds
    for row in ws.iter_rows(min_row=r0, max_row=min(r1, r0 + 120), min_col=c0, max_col=c1):
        for c in row:
            if c.value is not None and c.font and c.font.name:
                counts[c.font.name] = counts.get(c.font.name, 0) + 1
    return max(counts, key=counts.get) if counts else None


def _column_profiles(ws, bounds):
    """Per column: counts of text vs numeric/formula cells and avg text length."""
    r0, r1, c0, c1 = bounds
    prof = {}
    for col in range(c0, c1 + 1):
        text = num = textlen = nonempty = short = 0
        for row in range(r0, r1 + 1):
            cell = ws.cell(row=row, column=col)
            v = cell.value
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            nonempty += 1
            if isinstance(v, str) and not v.startswith("="):
                text += 1
                textlen += len(v.strip())
                if len(v.strip()) <= 2:
                    short += 1
            else:
                num += 1  # number, date, bool, or formula
        prof[col] = {
            "text": text, "num": num, "nonempty": nonempty,
            "avglen": (textlen / text) if text else 0, "short": short,
        }
    return prof


def detect_geometry(ws, bounds):
    """Guess marker col, label cols, data column span (the universal chassis)."""
    r0, r1, c0, c1 = bounds
    prof = _column_profiles(ws, bounds)

    # data_col_start: first column that begins a run where numeric/formula dominates.
    data_start = None
    cols = list(range(c0, c1 + 1))
    for i, col in enumerate(cols):
        run = cols[i:i + 3]
        dataish = sum(
            1 for cc in run
            if prof[cc]["nonempty"] and prof[cc]["num"] >= max(1, prof[cc]["text"])
        )
        if prof[col]["nonempty"] and prof[col]["num"] >= max(1, prof[col]["text"]) and dataish >= min(2, len(run)):
            data_start = col
            break
    if data_start is None:
        # Fallback: first non-empty column after any leading text columns.
        data_start = next((c for c in cols if prof[c]["num"] > 0), (c0 + 1 if c1 > c0 else c0))

    # data_col_end: last column at/after data_start that carries any content.
    data_end = data_start
    for col in range(data_start, c1 + 1):
        if prof.get(col, {}).get("nonempty", 0):
            data_end = col

    # label columns: text-bearing columns left of data_start (skip an empty col-A margin).
    label_cols, marker_col = [], None
    for col in range(c0, data_start):
        p = prof[col]
        if p["nonempty"] == 0:
            continue  # empty margin
        if p["text"] and p["short"] >= max(1, int(0.6 * p["text"])) and p["avglen"] <= 2.5:
            marker_col = col  # narrow column of single-char flags (e.g. "X")
        else:
            label_cols.append(col)
    if not label_cols:  # labels may sit in col A itself
        label_cols = [c for c in range(c0, data_start) if prof[c]["nonempty"]] or [c0]

    return {
        "marker_col": get_column_letter(marker_col) if marker_col else None,
        "label_cols": [get_column_letter(c) for c in label_cols],
        "data_col_start": get_column_letter(data_start),
        "data_col_end": get_column_letter(data_end),
    }


# ---------------------------------------------------------------------------
# Tab-type guessing
# ---------------------------------------------------------------------------
_NAME_RULES = [
    ("balance_sheet",    ["balance sheet", "statement of financial position",
                          "financial position"],                              0.92),
    ("cash_flow",        ["cash flow", "statement of cash flows"],             0.90),
    ("income_statement", ["income statement", "p&l", "p & l", "profit and loss",
                          "statement of operations", "p and l"],              0.88),
    ("liquidity",        ["liquidity", "borrowing base", "availability",
                          "revolver", "abl"],                                  0.82),
    ("output",           ["output", " out", "cim", "ufcf", "tear sheet",
                          "presentation"],                                     0.70),
    ("assumptions",      ["assumption", "drivers", "key inputs"],              0.82),
    ("cover",            ["cover", "title page", "disclaimer", "front page"],  0.85),
    ("forecast",         ["forecast", "cogs", "sales", "build", "projection",
                          "schedule", "roll", "model"],                        0.55),
]
# Generic name hints for unstyled working/scratch/source tabs. Kept deliberately
# example-agnostic -- no tab names lifted from any particular model.
_WORKING_HINTS = ["bridge", "scratch", "raw", "dump", "temp", "tmp", "wip", "work",
                  "paste", "import", "export", "tieout", "tie out", "tie-out",
                  "reconcil", "staging", "helper", "data pull", "data dump", "backup",
                  "do not print", "hidden", "junk", "sandbox"]


def guess_tab_type(name, ws, bounds, dom_font, existing_tab_rgb):
    n = name.strip().lower()
    if n.endswith(">"):
        return "divider", 0.99, "name ends with '>'"
    if ws.sheet_state != "visible":
        return "hidden", 0.9, f"sheet_state={ws.sheet_state}"

    # Strong structural signals from how the original modeler already tagged tabs.
    if existing_tab_rgb == "FF00B050":
        return "working", 0.8, "tab already colored working-green"
    if existing_tab_rgb == "FF00B0F0":
        return "source", 0.8, "tab already colored source-blue"

    # Name keywords (first match wins; order = specificity).
    name_guess = None
    for ttype, kws, conf in _NAME_RULES:
        if any(k in n for k in kws):
            name_guess = (ttype, conf, f"name matches {kws!r}")
            break
    if any(h in n for h in _WORKING_HINTS):
        # Working-ish name unless it clearly named a real statement/output.
        if not name_guess or name_guess[1] < 0.85:
            return "working", 0.6, "name looks like a working/scratch/source tab"

    if dom_font == sc.FONT_OUTPUT:  # Segoe UI -> presentation output page
        if not name_guess or name_guess[0] not in ("output",):
            return "output", 0.8, "dominant font is Segoe UI"

    if name_guess:
        return name_guess

    # Unnamed structure: gridlines left ON usually means an unstyled working tab.
    if ws.sheet_view.showGridLines is not False:
        return "working", 0.45, "gridlines on, no presentation name -> likely working"
    return "schedule", 0.4, "generic schedule (unclassified)"


# How each tab type behaves. format_body=False => leave the body untouched.
TAB_DEFAULTS = {
    #                     format_body  tab_color           header_family  numfmt_family  font
    "cover":            (True,  None,                       None,         "core",    sc.FONT_BODY),
    "assumptions":      (True,  None,                       "assumptions","core",    sc.FONT_BODY),
    "income_statement": (True,  "core_statement",           "core",       "core",    sc.FONT_BODY),
    "balance_sheet":    (True,  "core_statement",           "core",       "balance", sc.FONT_BODY),
    "cash_flow":        (True,  "core_statement",           "core",       "balance", sc.FONT_BODY),
    "liquidity":        (True,  "core_statement",           "liquidity",  "core",    sc.FONT_BODY),
    "forecast":         (True,  "ancillary",                "ancillary",  "core",    sc.FONT_BODY),
    "carveout":         (True,  "carveout",                 "core",       "core",    sc.FONT_BODY),
    "output":           (True,  "output",                   "core",       "core",    sc.FONT_OUTPUT),
    "schedule":         (True,  None,                       "core",       "core",    sc.FONT_BODY),
    "working":          (False, "working",                  "core",       "core",    sc.FONT_BODY),
    "source":           (False, "source",                   "core",       "core",    sc.FONT_BODY),
    "divider":          (False, None,                       "core",       "core",    sc.FONT_BODY),
    "hidden":           (False, None,                       "core",       "core",    sc.FONT_BODY),
}

# ---------------------------------------------------------------------------
# Row-class guessing
# ---------------------------------------------------------------------------
_UNIT_NOTE = re.compile(r"(in thousands|in millions|in 000|#'?s in|\$ in|000s|\(\$000|\$mm|\$000)", re.I)
_TOTAL_KW = ["total", "subtotal", "gross profit", "gross margin", "ebitda", "ebit",
             "net income", "net loss", "net sales", "net revenue", "net change",
             "net increase", "net decrease", "net cash", "cash at end", "cash at beginning",
             "ending cash", "beginning cash", "availability", "net working capital"]
_TIER3_KW = ["gross profit", "total assets", "total liabilities & equity",
             "total liabilities and equity", "total liabilities and members",
             "total equity and liabilities", "cash at end", "ending cash",
             "cash at end of period", "total availability", "net income", "net loss"]
_TIER2_KW = ["total cost of sales", "ebitda", "total liabilities", "net change in cash",
             "net increase in cash", "net decrease in cash", "total revenue", "total debt"]
_PCT_KW = ["%", "margin", "growth", "% of", "yoy", "y/y", "as a %", "rate"]


def _label_of(ws, row, label_cols):
    parts = []
    for letter in label_cols:
        v = ws.cell(row=row, column=column_index_from_string(letter)).value
        if isinstance(v, str) and v.strip():
            parts.append(v.strip())
    return " ".join(parts)


def _is_year(cell):
    v = cell.value
    return (_is_number(cell) and float(v).is_integer() and 1990 <= v <= 2100)


def classify_row(ws, row, geo, header_last):
    """Return dict(class, tier, confidence, label, reason) for one row.
    `header_last` is the ABSOLUTE row number of the last header-band row."""
    label_cols = geo["label_cols"]
    c_start = column_index_from_string(geo["data_col_start"])
    c_end = column_index_from_string(geo["data_col_end"])
    marker = column_index_from_string(geo["marker_col"]) if geo["marker_col"] else None

    label = _label_of(ws, row, label_cols)
    low = label.lower()

    data_cells = [ws.cell(row=row, column=c) for c in range(c_start, c_end + 1)]
    provs = [provenance(c) for c in data_cells]
    n_num = sum(1 for c in data_cells if _is_number(c))
    n_form = sum(1 for p in provs if p in ("formula", "cross_sheet"))
    n_data = sum(1 for p in provs if p != "empty")
    has_pct_fmt = any("%" in (c.number_format or "") for c in data_cells)
    marker_val = ws.cell(row=row, column=marker).value if marker else None
    af_text = any(isinstance(c.value, str) and c.value.strip().lower() in ("actual", "forecast", "fcst", "budget")
                  for c in data_cells)

    def out(cls, conf, reason, tier=None):
        return {"row": row, "class": cls, "tier": tier, "confidence": round(conf, 2),
                "label": label[:60], "reason": reason}

    # --- A row of bare 4-digit years is a year/period axis, whether or not it fell
    #     inside the detected header band (covers annual models with integer headers).
    if n_num and n_form == 0 and all(_is_year(c) for c in data_cells if _is_number(c)) \
            and not any(k in low for k in _TOTAL_KW):
        return out("header_year", 0.7 if row <= header_last else 0.55, "row of 4-digit years -> period axis")

    # --- Header band (top of sheet): dates, year helper, FY band, toggles, actual/forecast
    if row <= header_last:
        if any(c.is_date or "mmm" in (c.number_format or "").lower() or "yy" in (c.number_format or "").lower()
               for c in data_cells if c.value is not None):
            return out("header_date", 0.7, "date-like header row")
        if any(isinstance(c.value, str) and c.value.strip().startswith("=YEAR(") for c in data_cells):
            return out("header_year", 0.7, "=YEAR() helper row")
        if "toggle" in low:
            return out("toggle", 0.7, "toggle label")
        if "actual" in low or "forecast" in low or af_text:
            return out("actual_forecast", 0.7, "actual/forecast marker")
        if re.search(r"\bfy\s?\d{2,4}\b", low) or re.fullmatch(r"\s*\d{4}\s*", label):
            return out("header_fy", 0.6, "FY band label")
    if af_text and n_num == 0:
        return out("actual_forecast", 0.6, "actual/forecast marker row")

    if n_data == 0 and not label:
        return out("spacer", 0.8, "empty row")

    # --- Check rows (do these before totals; "check" can co-occur with numbers)
    if re.search(r"\bcheck\b", low) or low.strip() in ("tie", "ties", "tie-out", "tieout"):
        return out("check", 0.8, "label contains 'check'")

    # --- Unit note
    if _UNIT_NOTE.search(label) and n_num == 0:
        return out("unit_note", 0.65, "units note")

    # --- Banner vs sub-section label vs underlined header (label row, no numeric data)
    if label and n_num == 0 and n_form == 0:
        if marker_val and str(marker_val).strip().upper() == "X":
            return out("banner", 0.7, "label row with 'X' marker -> section banner")
        if low.endswith(":"):
            return out("subsection_label", 0.6, "label ends with ':' -> sub-section label")
        if label.isupper() and len(label) > 2:
            return out("banner", 0.5, "ALL-CAPS label row -> banner candidate")
        return out("label", 0.45, "text label row (banner/subsection/header? review)")

    # --- Percent / margin rows
    if (any(k in low for k in _PCT_KW) and "%" in low) or has_pct_fmt or (
        n_num and all(-1.0 <= (c.value or 0) <= 1.0 for c in data_cells if _is_number(c)) and "margin" in low
    ):
        return out("percent", 0.65, "percent / margin row")

    # --- Total rows + tier guess
    if any(k in low for k in _TOTAL_KW) and (n_form or n_num):
        tier = 1
        conf = 0.55
        if any(k in low for k in _TIER3_KW):
            tier, conf = 3, 0.6
        elif any(k in low for k in _TIER2_KW):
            tier, conf = 2, 0.6
        cls = {1: "subtotal", 2: "major_total", 3: "headline_total"}[tier]
        return out(cls, conf, f"total keyword -> tier {tier} guess (REVIEW tier)", tier=tier)

    # --- Data rows
    if n_data:
        # First data row since the last banner/subsection becomes the "$ first row".
        if n_num >= n_form:
            return out("data", 0.55, "loaded/numeric data row")
        return out("data", 0.5, "calculated data row")

    if label:
        return out("label", 0.4, "label-only row")
    return out("spacer", 0.6, "blank")


_BLOCK_RESET = {"banner", "subbanner", "subsection_label", "header_date", "header_fy",
                "header_year", "actual_forecast", "toggle", "subtotal", "major_total",
                "headline_total"}


def _mark_block_firsts(rows):
    """The first $-denominated data row of each block becomes 'data_first' (gets the
    '$' number format); later rows in the block stay plain. A block resets after any
    banner / sub-section header / total."""
    expect_first = True
    for rc in rows:
        c = rc["class"]
        if c in _BLOCK_RESET:
            expect_first = True
        elif c == "data" and expect_first:
            rc["class"] = "data_first"
            rc["reason"] += " | first $ row of block"
            expect_first = False


def detect_header_rows(ws, geo, bounds):
    """How many rows form the top header band (toggles/years/FY/dates/actual-forecast)."""
    r0 = bounds[0]
    c_start = column_index_from_string(geo["data_col_start"])
    last = r0 - 1
    for row in range(r0, min(bounds[1], r0 + 14) + 1):
        rowcells = [ws.cell(row=row, column=c) for c in range(c_start, min(c_start + 6, bounds[3] + 1))]
        label = _label_of(ws, row, geo["label_cols"]).lower()
        nonempty = [c for c in rowcells if c.value is not None]
        looks_header = (
            any(c.is_date for c in rowcells if c.value is not None)
            or any("yy" in (c.number_format or "").lower() for c in rowcells)
            or any(isinstance(c.value, str) and c.value.startswith("=YEAR(") for c in rowcells)
            or (nonempty and all(_is_year(c) for c in nonempty if _is_number(c)) and any(_is_number(c) for c in nonempty))
            or any(isinstance(c.value, str) and c.value.strip().lower() in ("actual", "forecast", "fcst", "budget") for c in rowcells)
            or "toggle" in label or "actual" in label or "forecast" in label
            or re.search(r"\bfy\s?\d{2,4}\b", label)
        )
        if looks_header:
            last = row
    return max(0, last - r0 + 1) if last >= r0 else 0


def detect_input_blocks(ws, geo, header_rows, bounds, tab_type):
    """Propose triple-mark rectangles of numeric hardcodes -- only where hardcodes
    are genuinely *inputs* (assumptions). On statements, hardcodes are loaded data
    (blue font, NOT yellow), so we never auto-yellow them; the reviewer adds blocks
    for live-column inputs (e.g. liquidity advance rates) by hand."""
    if tab_type not in ("assumptions",):
        return []
    c0 = column_index_from_string(geo["data_col_start"])
    c1 = column_index_from_string(geo["data_col_end"])
    r0 = bounds[0] + header_rows
    grid = {}
    for r in range(r0, bounds[1] + 1):
        for c in range(c0, c1 + 1):
            grid[(r, c)] = _is_number(ws.cell(row=r, column=c)) and provenance(ws.cell(row=r, column=c)) == "hardcode"
    # Connected components (4-neighbour) -> bounding boxes.
    seen, blocks = set(), []
    for (r, c), v in grid.items():
        if not v or (r, c) in seen:
            continue
        stack, cells = [(r, c)], []
        while stack:
            cr, cc = stack.pop()
            if (cr, cc) in seen or not grid.get((cr, cc)):
                continue
            seen.add((cr, cc))
            cells.append((cr, cc))
            stack += [(cr + 1, cc), (cr - 1, cc), (cr, cc + 1), (cr, cc - 1)]
        rs = [x[0] for x in cells]; cs = [x[1] for x in cells]
        blocks.append({
            "min_row": min(rs), "max_row": max(rs),
            "min_col": get_column_letter(min(cs)), "max_col": get_column_letter(max(cs)),
            "numfmt": "percent_input",
        })
    return blocks


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def audit(path):
    wb = load_workbook(path, data_only=False)
    sheets = []
    for ws in wb.worksheets:
        tc = ws.sheet_properties.tabColor
        existing_rgb = None
        if tc is not None:
            try:
                existing_rgb = tc.rgb if isinstance(tc.rgb, str) else None
            except Exception:
                existing_rgb = None

        bounds = _used_bounds(ws)
        dom_font = _dominant_font(ws, bounds)
        ttype, tconf, treason = guess_tab_type(ws.title, ws, bounds, dom_font, existing_rgb)
        fmt_body, tab_color_key, header_family, numfmt_family, font_family = TAB_DEFAULTS[ttype]

        sheet = {
            "name": ws.title,
            "tab_type": ttype,
            "tab_type_confidence": tconf,
            "tab_type_reason": treason,
            "format_body": fmt_body,
            "gridlines_off": bool(fmt_body),
            "tab_color": sc.TAB_COLORS.get(tab_color_key) if tab_color_key else None,
            "font_family": font_family,
            "header_family": header_family,
            "numfmt_family": numfmt_family,
        }

        if fmt_body:
            geo = detect_geometry(ws, bounds)
            header_rows = detect_header_rows(ws, geo, bounds)
            sheet.update(geo)
            sheet["header_rows"] = header_rows
            sheet["freeze"] = f"{geo['data_col_start']}{header_rows + 1}" if header_rows else None
            sheet["input_blocks"] = detect_input_blocks(ws, geo, header_rows, bounds, ttype)
            header_last = bounds[0] + header_rows - 1  # absolute last header-band row
            rows = []
            for r in range(bounds[0], bounds[1] + 1):
                rc = classify_row(ws, r, geo, header_last)
                if rc["class"] != "spacer":  # keep the map lean; spacers need no formatting
                    rows.append(rc)
            _mark_block_firsts(rows)
            sheet["rows"] = rows
        sheets.append(sheet)
    return {"workbook": path, "sheets": sheets}


def main():
    ap = argparse.ArgumentParser(description="Audit a workbook into a reviewable JSON map.")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--output", default="audit_map.json")
    args = ap.parse_args()

    m = audit(args.workbook)
    with open(args.output, "w") as f:
        json.dump(m, f, indent=2)

    # Human-readable summary so the reviewer immediately sees where to look.
    print(f"Audited {args.workbook} -> {args.output}")
    print(f"{len(m['sheets'])} tabs:\n")
    for s in m["sheets"]:
        flag = "" if s["tab_type_confidence"] >= 0.75 else "   <-- LOW CONFIDENCE, review"
        body = ""
        if s.get("format_body"):
            nrows = len(s.get("rows", []))
            low = sum(1 for r in s.get("rows", []) if r["confidence"] < 0.6)
            body = f"  [labels {s.get('label_cols')} data {s.get('data_col_start')}:{s.get('data_col_end')} hdr {s.get('header_rows')} | {nrows} rows, {low} low-conf]"
        else:
            body = "  [body left UNFORMATTED]"
        print(f"  {s['name']:<32} {s['tab_type']:<17} conf={s['tab_type_confidence']:<4}{body}{flag}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
