#!/usr/bin/env python3
"""
merge_format.py - apply a styling layer onto a workbook WITHOUT an openpyxl
round-trip of its content. The Mode-B package-safe applier.

apply_format.py styles a workbook beautifully, but it writes through openpyxl,
which on a feature-rich file strips cached formula values, drops chart style/rels
parts and drawings, expands shared formulas, and loses defined names. For an
already-rich source (charts, data tables, cached values) that we are formatting
from scratch, that damage is unacceptable.

So: run apply_format on a COPY to get a STYLE DONOR (correct styles.xml + per-cell
style indices + column widths + sheet views), then transplant only that styling
onto the ORIGINAL source package -- whose charts, drawings, media, defined names,
cached values, and shared formulas stay byte-for-byte intact.

    python merge_format.py SOURCE.xlsx STYLE_DONOR.xlsx -o OUT.xlsx

Transplanted: xl/styles.xml (whole); each cell's s= index; styled empty cells
(so banner/tier fills span correctly); <cols> widths; <sheetViews> (gridlines,
freeze); <sheetPr> tabColor. Kept from source: all cell values/formulas/cached
<v>, sharedStrings, workbook.xml (defined names), charts, drawings, media, rels.
"""

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, "/home/user/Skill-Builder-Twin-Star/model-format/scripts")
from clone_format import _sheet_parts

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", MAIN)
Q = lambda t: f"{{{MAIN}}}{t}"
_COL = re.compile(r"([A-Z]+)")


def _col_idx(ref):
    s = _COL.match(ref).group(1)
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def _nsdecls(worksheet_xml):
    """All namespace declarations on the <worksheet> root (so a sub-fragment with
    prefixed attrs like xr:uid parses and round-trips)."""
    m = re.search(r"<worksheet\b[^>]*>", worksheet_xml)
    decls = dict(re.findall(r'xmlns:([\w-]+)="([^"]+)"', m.group(0))) if m else {}
    dflt = re.search(r'xmlns="([^"]+)"', m.group(0)) if m else None
    decls[""] = dflt.group(1) if dflt else MAIN
    for p, u in decls.items():
        if p:
            ET.register_namespace(p, u)
    return decls


def _frag(block, decls):
    """Parse a worksheet sub-block, injecting all root ns declarations once."""
    decl = " ".join(f'xmlns{":" + p if p else ""}="{u}"' for p, u in decls.items())
    block = re.sub(r"<sheetData(\s|>)",
                   lambda m: f"<sheetData {decl}{m.group(1)}", block, count=1)
    return ET.fromstring(block)


def _merge_sheetdata(src_xml, don_xml):
    ms = re.search(r"<sheetData\b[^>]*>.*?</sheetData>|<sheetData\s*/>", src_xml, re.S)
    md = re.search(r"<sheetData\b[^>]*>.*?</sheetData>|<sheetData\s*/>", don_xml, re.S)
    if not ms or not md or ms.group(0).endswith("/>"):
        return src_xml  # nothing to merge
    src_sd = _frag(ms.group(0), _nsdecls(src_xml))
    don_sd = _frag(md.group(0), _nsdecls(don_xml))

    # donor: coord -> s index, and the donor cell element (for empty styled cells)
    don_s, don_cell, don_rows = {}, {}, {}
    for row in don_sd.findall(Q("row")):
        don_rows[row.get("r")] = row
        for c in row.findall(Q("c")):
            ref = c.get("r")
            don_s[ref] = c.get("s")
            don_cell[ref] = c

    src_rows = {row.get("r"): row for row in src_sd.findall(Q("row"))}
    out = ET.Element(Q("sheetData"))
    for rnum in sorted(set(src_rows) | set(don_rows), key=int):
        srow, drow = src_rows.get(rnum), don_rows.get(rnum)
        base_row = srow if srow is not None else drow
        nrow = ET.SubElement(out, Q("row"))
        for k, v in base_row.attrib.items():
            nrow.set(k, v)
        scells = {c.get("r"): c for c in srow.findall(Q("c"))} if srow is not None else {}
        coords = set(scells) | {r for r in don_cell if r[len(_COL.match(r).group(1)):] == rnum}
        for ref in sorted(coords, key=_col_idx):
            if ref in scells:
                cell = scells[ref]            # keep source content (value/formula/cached v)
                if ref in don_s and don_s[ref] is not None:
                    cell.set("s", don_s[ref]) # overlay donor style index
                elif ref in don_s:
                    cell.attrib.pop("s", None)
                nrow.append(cell)
            else:
                nrow.append(don_cell[ref])     # donor-only: empty styled cell (banner/tier span)
    merged = ET.tostring(out, encoding="unicode")
    return src_xml[:ms.start()] + merged + src_xml[ms.end():]


def _splice(out_xml, don_xml, tag, anchor_before=None):
    """Replace <tag>..</tag>/<tag/> in out_xml with donor's; if absent, insert it
    just before `anchor_before` (a literal start tag)."""
    pat = rf"<{tag}\b[^>]*>.*?</{tag}>|<{tag}\b[^>]*/>"
    md = re.search(pat, don_xml, re.S)
    if not md:
        return out_xml
    block = md.group(0)
    ms = re.search(pat, out_xml, re.S)
    if ms:
        return out_xml[:ms.start()] + block + out_xml[ms.end():]
    if anchor_before:
        i = out_xml.find(anchor_before)
        if i != -1:
            return out_xml[:i] + block + out_xml[i:]
    return out_xml


def _splice_tabcolor(out_xml, don_xml):
    md = re.search(r"<sheetPr\b[^>]*>.*?</sheetPr>|<sheetPr\b[^>]*/>", don_xml, re.S)
    if not md or "tabColor" not in md.group(0):
        return out_xml
    tc = re.search(r"<tabColor\b[^>]*/>", md.group(0)).group(0)
    ms = re.search(r"<sheetPr\b[^>]*>.*?</sheetPr>|<sheetPr\b[^>]*/>", out_xml, re.S)
    if ms:
        blk = ms.group(0)
        if "tabColor" in blk:
            blk2 = re.sub(r"<tabColor\b[^>]*/>", tc, blk)
        elif blk.endswith("/>"):
            blk2 = blk[:-2] + f">{tc}</sheetPr>"
        else:
            blk2 = blk.replace("</sheetPr>", f"{tc}</sheetPr>", 1)
            if "<tabColor" not in blk2:  # sheetPr had children but our insert missed
                blk2 = re.sub(r"(<sheetPr\b[^>]*>)", r"\1" + tc, blk, count=1)
        return out_xml[:ms.start()] + blk2 + out_xml[ms.end():]
    m = re.search(r"<worksheet\b[^>]*>", out_xml)
    return out_xml[:m.end()] + f"<sheetPr>{tc}</sheetPr>" + out_xml[m.end():]


def _merge_sheet(src_xml, don_xml):
    out = _merge_sheetdata(src_xml, don_xml)
    out = _splice(out, don_xml, "cols", anchor_before="<sheetData")
    out = _splice(out, don_xml, "sheetViews")
    out = _splice_tabcolor(out, don_xml)
    return out


def merge(source, donor, out_path):
    with zipfile.ZipFile(source) as zs:
        names = zs.namelist()
        payload = {n: zs.read(n) for n in names}
        sparts = _sheet_parts(zs)
    with zipfile.ZipFile(donor) as zd:
        dpayload = {n: zd.read(n) for n in zd.namelist()}
        dparts = _sheet_parts(zd)

    payload["xl/styles.xml"] = dpayload["xl/styles.xml"]  # whole styling vocabulary
    merged = 0
    for name, spart in sparts.items():
        dpart = dparts.get(name)
        if not dpart or spart not in payload or dpart not in dpayload:
            continue
        payload[spart] = _merge_sheet(payload[spart].decode("utf-8"),
                                      dpayload[dpart].decode("utf-8")).encode("utf-8")
        merged += 1

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:  # source namelist => charts/media/drawings/defined names all kept
            zo.writestr(n, payload[n])
    print(f"Styling transplanted onto source package: styles.xml + {merged} sheet(s).")
    print(f"Saved -> {out_path}")
    return 0


def main():
    ap = argparse.ArgumentParser(description="Transplant a style donor's formatting onto a source package losslessly.")
    ap.add_argument("source")
    ap.add_argument("style_donor")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    return merge(args.source, args.style_donor, args.output)


if __name__ == "__main__":
    sys.exit(main())
