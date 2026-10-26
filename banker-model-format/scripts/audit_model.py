"""Audit an Excel model against the house rules and print what to fix, tab by tab.

Usage: python3 audit_model.py model.xlsx [--font Garamond] [--max 8]

Checks (each finding names the cell):
  font        any non-empty cell not in the house font
  hardcode    a typed number that is not blue (header rows, year labels and toggles excepted)
  link        a formula that reads another tab but is not green
  formula     a same-tab formula that is green or blue
  literal     a formula with a typed constant in it (=A1*1.05): move the 1.05 to a blue input cell
  numfmt      a number format outside the house list (dates and toggles allowed)
  longtext    a long sentence sitting in a cell outside the label column: it belongs in a cell comment
  layout      gridlines on, no title in row 2 or 3, no 'End' marker in column A
  delivery    workbook not opening on the first tab, tabs not at A1
It reads formulas, not values, so it needs no recalculation. Findings are advice: a reviewer decides.
"""
import argparse
import re
from collections import defaultdict

from openpyxl import load_workbook

HOUSE_NF = {
    '_(#,##0_)_%;\\(#,##0\\)_%;_("–"_)_%;_(@_)_%',
    '_(#,##0.0_)_%;\\(#,##0.0\\)_%;_("–"_)_%;_(@_)_%',
    '_(#,##0.00_)_%;\\(#,##0.00\\)_%;_("–"_)_%;_(@_)_%',
    '_([$$]#,##0_)_%;\\([$$]#,##0\\)_%;_("–"_)_%;_(@_)_%',
    '_([$$]#,##0.0_)_%;\\([$$]#,##0.0\\)_%;_("–"_)_%;_(@_)_%',
    '_([$$]#,##0.00_)_%;\\([$$]#,##0.00\\)_%;_("–"_)_%;_(@_)_%',
    '_(#,##0.0%_);\\(#,##0.0%\\);_("–"_)_%;_(@_)_%',
    '_(#,##0.00%_);\\(#,##0.00%\\);_("–"_)_%;_(@_)_%',
    '_(#,##0.000%_);\\(#,##0.000%\\);_("–"_)_%;_(@_)_%',
    '_(#,##0.0"x"_)_%;\\(#,##0.0"x"\\)_%;_("–"_)_%;_(@_)_%',
    '_(#,##0.00"x"_)_%;\\(#,##0.00"x"\\)_%;_("–"_)_%;_(@_)_%',
    '_(#,##0.000_)_%;\\(#,##0.000\\)_%;_("–"_)_%;_(@_)_%',
    '_([$$]#,##0.000_)_%;\\([$$]#,##0.000\\)_%;_("–"_)_%;_(@_)_%',
    '_(0.0\\x_)_)_\';_(\\(0.0\\x\\)_\'_\';_("–"_)_%;_(@_)_%',
    '"S + "0.0', '"Year "0', 'General', 'General"A"', '0"E"', '#"E"', '0', '0%', '0.0', ';;;', '0_)_%', '"S + "#',
}
def canon(nf):
    """Excel and LibreOffice escape literals differently ("x" vs \\x); compare without quotes and backslashes."""
    return nf.replace("\\", "").replace('"', "")


HOUSE_CANON = {canon(x) for x in HOUSE_NF}
ALLOWED_NF = re.compile(r"(m+|d+|y+)[^%]*$|^\[=|^\"Case")         # dates, conditional toggles, case labels
BLUE, GREEN = ("FF0000FF", "0000FF"), ("FF008000", "008000")
HEADER_FILLS = {"FFE8E8E8", "FFD1D1D1", "FF002060", "FF215F9A", "FF4E95D9", "FFA6CAEC"}
LITERAL = re.compile(r"(?<![A-Za-z$:!\d.])(\d+\.\d+|\d{3,})(?![\d:]|[A-Za-z(])")   # 1.05, 250 (not A12, not 1:1)
OK_LITERALS = {"100", "365", "360", "1000", "0.5"}
CELL_REF = re.compile(r"\$?[A-Z]{1,3}\$?\d+")


def rgb(color):
    if color is None:
        return None
    return color.rgb if color.type == "rgb" else f"theme{color.theme}"


def fill_rgb(cell):
    if not cell.fill.fill_type:
        return None
    c = cell.fill.fgColor
    return c.rgb if c.type == "rgb" else f"theme{c.theme}{c.tint:+.2f}"


def is_header(cell):
    f = fill_rgb(cell)
    return f in HEADER_FILLS or (f or "").startswith(("theme2", "theme3"))


def audit(path, house_font="Garamond", max_show=8):
    wb = load_workbook(path)
    out = defaultdict(lambda: defaultdict(list))
    for ws in wb.worksheets:
        if ws.title.endswith(">"):
            continue
        t = ws.title
        if ws.sheet_view.showGridLines:
            out[t]["layout"].append("gridlines are on")
        if not any(ws[f"C{r}"].value for r in (2, 3)):
            out[t]["layout"].append("no tab title in C2 / C3")
        if t != "Cover Page" and not any(str(ws[f"A{r}"].value).strip() == "End" for r in range(1, ws.max_row + 1)):
            out[t]["layout"].append("no 'End' marker in column A")
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                if c.font.name != house_font:
                    out[t]["font"].append(f"{c.coordinate} {c.font.name}")
                color = rgb(c.font.color)
                is_f = isinstance(v, str) and v.startswith("=")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    nf = c.number_format
                    year_like = 1990 <= v <= 2100 and float(v).is_integer()
                    if v != 0 and not is_header(c) and not year_like and not nf.startswith("[=") and color not in BLUE:
                        out[t]["hardcode"].append(f"{c.coordinate}={v}")
                elif is_f and not CELL_REF.search(v) and "!" not in v and re.fullmatch(r"=[\d.+\-*/() ]+", v):
                    if color not in BLUE:                                  # =1053+55 is typed data
                        out[t]["hardcode"].append(f"{c.coordinate}{v}")
                    continue
                elif is_f:
                    cross = "!" in v
                    output_tab = "Output" in t or t == "Cover Page"        # presentation tabs stay black
                    if cross and color not in GREEN and not is_header(c) and c.column > 2 and not output_tab:
                        out[t]["link"].append(f"{c.coordinate}")
                    if not cross and (color in GREEN or color in BLUE):
                        out[t]["formula"].append(f"{c.coordinate} {'green' if color in GREEN else 'blue'}")
                    body = re.sub(r'"[^"]*"', "", v)                     # ignore text inside quotes
                    body = re.sub(r"'[^']*'!", "", body)                 # and sheet names
                    lits = [x for x in LITERAL.findall(body)
                            if x not in OK_LITERALS and not re.fullmatch(r"(19|20)\d\d", x)]
                    if lits and c.column > 2:
                        out[t]["literal"].append(f"{c.coordinate} {v[:40]}")
                elif isinstance(v, str) and len(v) > 70 and c.column > 3 and t != "Cover Page":
                    out[t]["longtext"].append(f"{c.coordinate} '{v[:40]}...'")
                nf = c.number_format
                if not is_f and not isinstance(v, (int, float)):
                    continue
                if canon(nf) not in HOUSE_CANON and not ALLOWED_NF.search(nf):
                    out[t]["numfmt"].append(f"{c.coordinate} {nf}")
        sel = ws.sheet_view.selection[0].activeCell if ws.sheet_view.selection else "A1"
        if (ws.sheet_view.topLeftCell or "A1") != "A1" or (sel or "A1") != "A1":
            out[t]["delivery"].append(f"view not at A1 (top-left {ws.sheet_view.topLeftCell}, cursor {sel})")
    if wb.active is not wb.worksheets[0]:
        out["(workbook)"]["delivery"].append(f"opens on '{wb.active.title}', not the first tab")

    total = 0
    for t, kinds in out.items():
        print(f"\n== {t}")
        for k, items in kinds.items():
            total += len(items)
            show = ", ".join(items[:max_show]) + (f" ... (+{len(items) - max_show})" if len(items) > max_show else "")
            print(f"  {k:9s} {len(items):4d}  {show}")
    print(f"\n{total} finding(s) across {len(out)} tab(s)")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--font", default="Garamond")
    ap.add_argument("--max", type=int, default=8)
    a = ap.parse_args()
    audit(a.path, a.font, a.max)
