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
