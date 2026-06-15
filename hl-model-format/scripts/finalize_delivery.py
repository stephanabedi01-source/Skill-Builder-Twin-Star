#!/usr/bin/env python3
"""
finalize_delivery.py - the universal final-save step for EVERY workbook handed back.

Applies the delivery standard, regardless of mode (replicate or format-from-
scratch) and regardless of what the source's views said:

  * The workbook opens on its first VISIBLE tab (workbookView activeTab points
    there; only that sheet's view carries tabSelected="1").
  * Every tab's cursor is at A1 (selection activeCell/sqref = A1) and scrolled to
    the top-left (saved topLeftCell cleared; a frozen pane is KEPT exactly, but
    its panes are scrolled back to their origin and A1 is selected).
  * Every tab is at 85% zoom (zoomScale="85" on each sheet view).

This is a DELIBERATE exception to "preserve the source's views exactly": active
tab, selected cell, scroll position, and zoom always follow this standard at
delivery. Everything else about views (gridlines, view mode, the freeze itself)
is left as the source/earlier steps set it.

It edits only the <bookViews> tag in xl/workbook.xml and each worksheet's
<sheetViews> block, splicing the rest of every part back byte-for-byte. So on a
Mode-A clone, styles, theme, drawings, charts, media, cached values, and defined
names stay intact -- only the view records change, which is exactly the intent.

    python finalize_delivery.py WORKBOOK.xlsx            # in place
    python finalize_delivery.py WORKBOOK.xlsx -o OUT.xlsx
"""

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from openpyxl.utils import get_column_letter

ZOOM = "85"


def _first_visible(workbook_xml):
    """(index, name) of the first sheet whose state is visible."""
    sheets = re.findall(r"<sheet\b[^>]*/>", workbook_xml)
    for i, s in enumerate(sheets):
        state = re.search(r'state="([^"]+)"', s)
        if not state or state.group(1) not in ("hidden", "veryHidden"):
            name = re.search(r'name="([^"]*)"', s)
            return i, (name.group(1) if name else None)
    return 0, None


def _sheet_name_to_part(zf):
    from clone_format import _sheet_parts  # reuse the rels resolver
    return _sheet_parts(zf)


def _set_active_tab(workbook_xml, idx):
    m = re.search(r"<workbookView\b[^>]*?/?>", workbook_xml)
    if not m:
        return workbook_xml
    tag = m.group(0)
    if "activeTab=" in tag:
        new = re.sub(r'activeTab="[^"]*"', f'activeTab="{idx}"', tag)
    else:
        new = tag[:-2] + f' activeTab="{idx}"' + tag[-2:] if tag.endswith("/>") \
            else tag[:-1] + f' activeTab="{idx}">'
    return workbook_xml[:m.start()] + new + workbook_xml[m.end():]


def _fix_sheetview(el, tab_selected):
    el.set("zoomScale", ZOOM)
    if tab_selected:
        el.set("tabSelected", "1")
    else:
        el.attrib.pop("tabSelected", None)
    el.attrib.pop("topLeftCell", None)  # clear saved scroll

    panes = el.findall("pane")
    others = [c for c in el if c.tag not in ("pane", "selection")]
    for c in list(el):
        el.remove(c)
    for o in others:  # preserve anything exotic (rare) ahead of pane/selection
        el.append(o)
    if panes:
        p = panes[0]
        if p.get("state") in ("frozen", "frozenSplit"):
            xs = int(float(p.get("xSplit", "0") or 0))
            ys = int(float(p.get("ySplit", "0") or 0))
            p.set("topLeftCell", f"{get_column_letter(xs + 1)}{ys + 1}")  # scroll to origin
            p.set("activePane", "topLeft")  # so A1 (top-left pane) is the selection
        el.append(p)
    sel = ET.Element("selection")
    sel.set("activeCell", "A1")
    sel.set("sqref", "A1")
    el.append(sel)


def _patch_sheet(part_xml, tab_selected):
    m = re.search(r"<sheetViews\b.*?</sheetViews>", part_xml, re.S)
    if not m:
        return part_xml  # no views block; nothing to normalize
    frag = m.group(0)
    # The fragment inherits namespace prefixes (xr:uid etc.) declared on the
    # worksheet root -- harvest and re-declare them so the fragment parses, and
    # register them so attributes round-trip with their original prefixes.
    root_tag = re.search(r"<worksheet\b[^>]*>", part_xml)
    nsdecls = dict(re.findall(r'xmlns:([\w-]+)="([^"]+)"', root_tag.group(0))) if root_tag else {}
    for p, u in nsdecls.items():
        ET.register_namespace(p, u)
    wrapper = ("<__w " + " ".join(f'xmlns:{p}="{u}"' for p, u in nsdecls.items()) + ">"
               + frag + "</__w>")
    block = ET.fromstring(wrapper)[0]  # the sheetViews element
    for sv in block.findall("sheetView"):
        _fix_sheetview(sv, tab_selected)
    new_block = ET.tostring(block, encoding="unicode")
    return part_xml[:m.start()] + new_block + part_xml[m.end():]


def finalize(path, out_path):
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        payload = {n: zin.read(n) for n in names}
        parts = _sheet_name_to_part(zin)

    wbx = payload["xl/workbook.xml"].decode("utf-8")
    idx, first_name = _first_visible(wbx)
    payload["xl/workbook.xml"] = _set_active_tab(wbx, idx).encode("utf-8")
    first_part = parts.get(first_name)

    touched = 0
    for name, part in parts.items():
        if part not in payload:
            continue
        x = payload[part].decode("utf-8")
        new = _patch_sheet(x, tab_selected=(part == first_part))
        if new != x:
            payload[part] = new.encode("utf-8")
            touched += 1

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, payload[n])
    print(f"Delivery standard applied: opens on '{first_name}' (tab {idx}); "
          f"{touched} sheet view(s) set to A1 + {ZOOM}% zoom.")
    print(f"Saved -> {out_path}")
    return 0


def main():
    ap = argparse.ArgumentParser(description="Apply the universal delivery standard (first tab / A1 / 85% zoom).")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()
    return finalize(args.workbook, args.output or args.workbook)


if __name__ == "__main__":
    sys.exit(main())
