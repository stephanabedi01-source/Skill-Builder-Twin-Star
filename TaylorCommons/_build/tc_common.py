"""Shared helpers for the Taylor Commons slide builders (on top of deck_kit)."""
import copy
import os
import sys

sys.path.insert(0, "/root/.claude/skills/synced/234aa966-38be-4785-ad69-b2aef45c9c62_70a9bf30-54a6-48d3-8be2-e700e02c988d/"
                   "banker-deck-builder/scripts")
from deck_kit import *  # noqa: F401,F403
from deck_kit import THEME, load_theme, textbox, PP_ALIGN, MSO_ANCHOR, qn, Inches, Pt  # noqa: F401

HERE = os.path.dirname(os.path.abspath(__file__))
load_theme(os.path.join(HERE, "theme.json"))


def add_slide_number(slide):
    """Clone the layout's slide-number placeholder onto the slide so the page number shows in the footer."""
    for sp in slide.slide_layout.shapes._spTree.iterchildren(qn("p:sp")):
        ph = sp.find(".//" + qn("p:ph"))
        if ph is not None and ph.get("type") == "sldNum":
            slide.shapes._spTree.append(copy.deepcopy(sp))
            return


def footer_notes(slide, notes=None, sources=None, size=9):
    """Short notes and sources in the footer band (below the rule at 6.82, right of the logo, left of the fund name)."""
    paras = []
    if notes:
        paras.append([(notes, {"c": "grey_txt"})])
    if sources:
        paras.append([(sources, {"c": "grey_txt", "i": True})])
    textbox(slide, 1.45, 6.87, 8.50, 0.55, paras, size=size, name="Footer notes and sources", anchor=MSO_ANCHOR.TOP)


NOTES_BODY = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="3" '
              'name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" '
              'idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr '
              'lang="en-US"/></a:p></p:txBody></p:sp>')


def speaker_notes(slide, text):
    """Write speaker notes. This deck's notes master has no body placeholder, so add one when missing."""
    ns = slide.notes_slide
    if ns.notes_text_frame is None:
        from pptx.oxml import parse_xml
        ns.shapes._spTree.append(parse_xml(NOTES_BODY))
    ns.notes_text_frame.text = text


def closing(slide, runs):
    """Closing line at the deck's body size (12pt)."""
    closing_line(slide, runs, size=12)


def header(slide, x, y, w, text, name, sup=None):
    section_header(slide, x, y, w, text, name, sup=sup, size=15)


def caption(slide, x, y, w, text, name, sup=None, size=11):
    """Italic grey caption under a panel header (11pt on this deck)."""
    runs = [(text, {"i": True, "c": "grey_txt"})] + ([(sup, {"sup": True, "i": True, "c": "grey_txt"})] if sup else [])
    return textbox(slide, x, y, w, 0.22, [runs], size=size, name=name)


def new_slide(prs, title, takeaway):
    """new_content_slide on the plain CONTENT layout, found by name (robust to slides being inserted)."""
    idx = next(i for i, sl in enumerate(prs.slides) if sl.slide_layout.name == "CONTENT")
    return new_content_slide(prs, title, takeaway, template_index=idx)


def photo(slide, key, x, y, w, h, name, focus=0.5):
    """Place img/<key>.jpg center-cropped to exactly w x h inches (focus = vertical crop position 0..1)."""
    from PIL import Image
    src = os.path.join(HERE, "img", f"{key}.jpg")
    out = os.path.join(HERE, "img", "_crops", f"{key}_{w:.2f}x{h:.2f}.jpg")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    im = Image.open(src).convert("RGB")
    iw, ih = im.size
    r = w / h
    if iw / ih > r:
        nw = int(ih * r)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / r)
        top = int((ih - nh) * focus)
        im = im.crop((0, top, iw, top + nh))
    im.save(out, quality=88)
    return add_image(slide, out, x, y, w, h, name=name)


def drop_slide_titled(prs, title, layout=None):
    """Remove the first slide whose title placeholder text equals `title` (optionally on a given layout)."""
    for i, sl in enumerate(prs.slides):
        if layout and sl.slide_layout.name != layout:
            continue
        t = sl.shapes.title
        if t is not None and t.text_frame.text.strip() == title:
            drop_slide(prs, i)
            return i
    return None
