#!/usr/bin/env python3
"""
clone_format.py - lossless formatting transplant at the package (ZIP/XML) level.

THE RULE THIS TOOL EXISTS FOR: when a workbook already has an established format,
the job is to replicate that workbook's exact formatting -- clone its style
records byte-for-byte -- not to re-derive formatting from house conventions.
Library round-trips (openpyxl included) are lossy: they strip cached formula
values, drop embedded images, re-serialize charts (losing styleN.xml/colorsN.xml,
effect lists, manual layouts, deleted-title flags), normalize row/column
geometry, collapse multiple sheetViews, and dedupe defined names. The only
faithful carrier of a workbook's formatting is its own package parts.

So this tool starts from the SOURCE package verbatim -- styles.xml, theme,
numFmts, sheetViews, row/col geometry, tab colors, drawings, charts, media,
printerSettings, defined names, cached values, everything -- and transplants
only CONTENT differences from TARGET into it:

    python clone_format.py SOURCE.xlsx TARGET.xlsx -o OUT.xlsx

  * SOURCE = the formatting donor (the workbook whose look is ground truth).
  * TARGET = the content keeper (same sheet structure; possibly updated values).
  * OUT    = SOURCE's package with TARGET's differing cell contents patched in.

If TARGET's content is identical to SOURCE's (the common "restore / losslessly
reformat" case), zero patches are needed and OUT is SOURCE byte-for-byte --
which is exactly what lossless means. Patched formula cells lose their cached
value, so when any patch touches a formula the tool sets fullCalcOnLoad.

Requires the same sheet names in the same order. It will not invent structure.
"""

import argparse
import hashlib
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date, time, timedelta

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# Content comparison (openpyxl, read-only: fast, and both sides are read with
# the same representation -- formulas as text, dates as datetimes)
# ---------------------------------------------------------------------------
def _canon(v):
    """Canonical form for comparison."""
    if v is None:
        return None
    if isinstance(v, str):
        # An empty string renders exactly like an empty cell, and library
        # round-trips routinely drop them -- treat as empty, never as a diff.
        return v if v != "" else None
    if isinstance(v, (datetime, date, time)):
        # Serial<->datetime conversion carries microsecond float noise; compare
        # at whole-second resolution.
        if isinstance(v, datetime):
            v = (v + timedelta(microseconds=500000)).replace(microsecond=0)
        return ("dt", str(v))
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    # ArrayFormula and friends -> their text
    t = getattr(v, "text", None)
    return ("af", t) if t is not None else ("repr", repr(v))


def _num_close(a, b):
    if isinstance(a, tuple) and isinstance(b, tuple) and a[0] == "n" and b[0] == "n":
        x, y = a[1], b[1]
        return abs(x - y) <= 1e-9 * max(1.0, abs(x), abs(y))
    return False


def diff_content(src_path, tgt_path):
    """Return {sheet_name: {coord: target_value}} for cells whose content differs."""
    from openpyxl import load_workbook
    ws_diffs = {}
    src = load_workbook(src_path, read_only=True, data_only=False)
    tgt = load_workbook(tgt_path, read_only=True, data_only=False)
    if src.sheetnames != tgt.sheetnames:
        raise SystemExit(
            "Sheet structure differs between source and target:\n"
            f"  source: {src.sheetnames}\n  target: {tgt.sheetnames}\n"
            "clone_format requires the same sheets in the same order."
        )
    for name in src.sheetnames:
        s_cells, t_cells = {}, {}
        for ws, store in ((src[name], s_cells), (tgt[name], t_cells)):
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        store[c.coordinate] = _canon(c.value)
        diffs = {}
        for coord in set(s_cells) | set(t_cells):
            a, b = s_cells.get(coord), t_cells.get(coord)
            if a != b and not _num_close(a, b):
                diffs[coord] = b
        if diffs:
            ws_diffs[name] = diffs
        # free read-only resources promptly on big tabs
    src.close()
    tgt.close()
    return ws_diffs


# ---------------------------------------------------------------------------
# Package patching (only runs when content actually differs)
# ---------------------------------------------------------------------------
def _sheet_parts(zf):
    """Map sheet name -> xl/worksheets/sheetN.xml part name, via workbook rels."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_part = {}
    for rel in rels:
        t = rel.get("Target").lstrip("/")
        rid_to_part[rel.get("Id")] = t if t.startswith("xl/") else "xl/" + t
    out = {}
    for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
        out[sh.get("name")] = rid_to_part[sh.get(f"{{{NS_REL}}}id")]
    return out


def _patch_sheet_xml(xml_bytes, diffs):
    """Set/replace cell content in one worksheet part. Keeps every s= style index.
    Returns (new_bytes, touched_formula)."""
    ET.register_namespace("", NS_MAIN)
    root = ET.fromstring(xml_bytes)
    data = root.find(f"{{{NS_MAIN}}}sheetData")
    touched_formula = False
    by_row = {}
    for coord, val in diffs.items():
        m = re.match(r"([A-Z]+)(\d+)", coord)
        by_row.setdefault(int(m.group(2)), {})[coord] = val
    for row in data.findall(f"{{{NS_MAIN}}}row"):
        r = int(row.get("r"))
        if r not in by_row:
            continue
        cells = {c.get("r"): c for c in row.findall(f"{{{NS_MAIN}}}c")}
        for coord, val in by_row[r].items():
            c = cells.get(coord)
            if c is None:
                c = ET.SubElement(row, f"{{{NS_MAIN}}}c")
                c.set("r", coord)  # NOTE: appended; order within row not re-sorted
            # wipe old content, keep style
            for child in list(c):
                c.remove(child)
            c.attrib.pop("t", None)
            if val is None:
                continue
            kind = val[0] if isinstance(val, tuple) else "s"
            if isinstance(val, str) and val.startswith("="):
                f = ET.SubElement(c, f"{{{NS_MAIN}}}f")
                f.text = val[1:]
                touched_formula = True  # no cached <v>: needs recalc
            elif isinstance(val, str):
                c.set("t", "inlineStr")
                is_ = ET.SubElement(c, f"{{{NS_MAIN}}}is")
                t = ET.SubElement(is_, f"{{{NS_MAIN}}}t")
                t.text = val
            elif kind == "b":
                c.set("t", "b")
                ET.SubElement(c, f"{{{NS_MAIN}}}v").text = "1" if val[1] else "0"
            elif kind == "af":
                f = ET.SubElement(c, f"{{{NS_MAIN}}}f")
                f.set("t", "array")
                f.set("ref", coord)
                f.text = (val[1] or "").lstrip("=")
                touched_formula = True
            elif kind == "dt":
                # write the Excel serial; the cell's (source) number format renders it
                from openpyxl.utils.datetime import to_excel
                from datetime import datetime as _dt
                serial = to_excel(_dt.fromisoformat(val[1]))
                ET.SubElement(c, f"{{{NS_MAIN}}}v").text = repr(serial)
            elif kind == "n":
                ET.SubElement(c, f"{{{NS_MAIN}}}v").text = repr(val[1])
            else:
                raise ValueError(f"unsupported patch value at {coord}: {val!r}")
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), touched_formula


def _set_full_calc(xml_bytes):
    ET.register_namespace("", NS_MAIN)
    root = ET.fromstring(xml_bytes)
    calc = root.find(f"{{{NS_MAIN}}}calcPr")
    if calc is None:
        calc = ET.SubElement(root, f"{{{NS_MAIN}}}calcPr")
    calc.set("fullCalcOnLoad", "1")
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def clone(src_path, tgt_path, out_path):
    print("Comparing content (source vs target)...")
    diffs = diff_content(src_path, tgt_path)
    n = sum(len(d) for d in diffs.values())
    if n == 0:
        # Lossless case: the donor already carries the content. Emit it verbatim.
        shutil.copyfile(src_path, out_path)
        print("0 content differences across all sheets.")
        print("-> source package emitted VERBATIM (formatting, drawings, charts,")
        print("   media, cached values, defined names, views: all byte-identical).")
        return 0
    print(f"{n} content difference(s) on {len(diffs)} sheet(s); patching source package...")
    with zipfile.ZipFile(src_path) as zin:
        parts = _sheet_parts(zin)
        names = zin.namelist()
        payload = {nm: zin.read(nm) for nm in names}
    any_formula = False
    for sheet, d in diffs.items():
        part = parts[sheet]
        payload[part], tf = _patch_sheet_xml(payload[part], d)
        any_formula |= tf
        print(f"   {sheet}: {len(d)} cell(s) patched")
    if any_formula:
        payload["xl/workbook.xml"] = _set_full_calc(payload["xl/workbook.xml"])
        print("   formulas were patched -> fullCalcOnLoad set (recalc on open,")
        print("   or run a headless recalc before delivery to restore cached values)")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm in names:
            zout.writestr(nm, payload[nm])
    return 0


def main():
    ap = argparse.ArgumentParser(description="Clone SOURCE's formatting verbatim; keep TARGET's content.")
    ap.add_argument("source", help="formatting donor (ground-truth look)")
    ap.add_argument("target", help="content keeper (same sheet structure)")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    rc = clone(args.source, args.target, args.output)
    h = hashlib.sha256(open(args.output, "rb").read()).hexdigest()[:16]
    print(f"Saved -> {args.output}  (sha256:{h})")
    return rc


if __name__ == "__main__":
    sys.exit(main())
