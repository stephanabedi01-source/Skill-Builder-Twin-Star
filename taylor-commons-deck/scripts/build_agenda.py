"""Slide: Agenda. Also puts the deck in agenda order.

Usage: python3 build_agenda.py <in.pptx> <out.pptx>

Sections and order come from the team's agenda (10 sections). Slides are matched by title prefix; page ranges on
the agenda are computed from where each section's slides end up, so they stay right if slides move.
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]


def title_of(sl):
    for ph in sl.placeholders:
        if ph.placeholder_format.type == 1 and ph.has_text_frame:
            return ph.text_frame.text.strip()
    return ""


# (section, descriptor, title prefixes in order)
SECTIONS = [
    ("Asset Overview", "224 units, built 1979, bought in Jun-23 for $25.9M", ["Asset Overview"]),
    ("Site Visit Findings", "Still hard to find; a model unit now, but upkeep is slipping", ["Site Visit: Location",
                                                                                   "Site Visit: Leasing"]),
    ("Location & Market", "DFW is recovering; east Fort Worth Class C is not",
     ["Location Overview", "Migration Trends", "Market Overview", "Submarket"]),
    ("Performance vs. Underwriting", "75% occupied; T12 NOI of $1.43M vs. $2.21M planned",
     ["Then vs. Now", "Operations", "Historical Financials"]),
    ("Comparables", "Comp rents down 8% to 16% since IC; sales at $91K to $137K a unit",
     ["Current Rent Comparables", "Current Sales Comparables"]),
    ("Capital Structure & 2027 Maturity", "9% bridge, HB 21 tax cliff and a $7.1M refinancing gap",
     ["The March 2025 Refinancing", "HB 21", "The 2027 Cliff", "Valuation and Recovery"]),
    ("Renovation Plan", "A $1.8M upgrade budget vs. 12 renovated units", ["Renovation Plan"]),
    ("What Went Wrong", "Operations missed first; the HFC refinancing made it worse", ["What Went Wrong", "HFC Update"]),
    ("Lessons for the Fund", "What the fund does now, and screens to catch this earlier",
     ["How Can the Fund", "Final Recommendation"]),
    ("Appendix", "Photos and supporting detail", ["Appendix", "Asset Photos"]),
]

# ---------------------------------------------------------------- reorder the deck to agenda order
slides = list(prs.slides)
titles = [title_of(sl) for sl in slides]
order = [0]                                   # cover
agenda_idx = next((i for i, t in enumerate(titles) if t == "Agenda"), None)
if agenda_idx is not None:
    order.append(agenda_idx)
for _, _, prefixes in SECTIONS:
    for p in prefixes:
        for i, t in enumerate(titles):
            if i not in order and t.startswith(p):
                order.append(i)
order += [i for i in range(len(slides)) if i not in order]   # anything unmatched stays, at the end
lst = prs.slides._sldIdLst
ids = list(lst)
for el in ids:
    lst.remove(el)
for i in order:
    lst.append(ids[i])

# ---------------------------------------------------------------- page ranges after reordering
titles = [title_of(sl) for sl in prs.slides]
pages = []
for _, _, prefixes in SECTIONS:
    nums = [i + 1 for i, t in enumerate(titles) if any(t.startswith(p) for p in prefixes)]
    pages.append((min(nums), max(nums)) if nums else None)

# ---------------------------------------------------------------- agenda slide
s = new_slide(prs, "Agenda",
              "Three years into a five-year hold, the common equity is out of the money; this update shows why and "
              "what the fund does next")
add_slide_number(s)
L0, R1 = g["L0"], g["R1"]
y0, rh = 1.58, 0.47
for k, ((name, desc, _), pg) in enumerate(zip(SECTIONS, pages)):
    y = y0 + k * rh
    d = 0.3
    circ = box(s, L0 + 0.05, y + (rh - d) / 2 - 0.02, d, d, fill="ink", line=None, shape=MSO_SHAPE.OVAL,
               name=f"Agenda num {k+1}")
    tf = circ.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(tf, [[(str(k + 1), {"b": True, "c": "white"})]], 12, align=PP_ALIGN.CENTER)
    textbox(s, L0 + 0.55, y, 4.2, rh - 0.04, [[(name, {"b": True})]], size=16, anchor=MSO_ANCHOR.MIDDLE,
            name=f"Agenda section {k+1}")
    textbox(s, 5.3, y, 6.3, rh - 0.04, [[(desc, {"i": True, "c": "grey_txt"})]], size=12.5,
            anchor=MSO_ANCHOR.MIDDLE, name=f"Agenda desc {k+1}")
    if pg:
        txt = f"{pg[0]}" if pg[0] == pg[1] else f"{pg[0]} to {pg[1]}".replace(" ", " ")
        textbox(s, R1 - 1.1, y, 1.1, rh - 0.04, [[(txt, {"c": "grey_txt"})]], size=12.5, align=PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"Agenda pages {k+1}")
    if k < len(SECTIONS) - 1:
        line(s, L0, y + rh - 0.02, R1, y + rh - 0.02, "hair", 0.75, name=f"Agenda rule {k+1}")
textbox(s, R1 - 1.1, y0 - 0.26, 1.1, 0.22, [[("Pages", {"b": True, "c": "grey_txt"})]], size=11,
        align=PP_ALIGN.RIGHT, name="Agenda pages header")
speaker_notes(s, """Presenter: TBD

Here is how we will walk through Taylor Commons. We start with the asset and what we saw on our site visit, then the location and the market around it, then how the property has performed against what we underwrote and how it compares to its competitors. The core of the update is section six: the capital structure, the 2025 refinancing into the HFC, the HB 21 tax cliff and the March 2027 loan maturity. We finish with the renovation plan, what went wrong, and what the fund should do now and learn for future deals.""")
place(prs, s, 1)
prs.save(DST)
print("saved", DST)
for (name, _, _), pg in zip(SECTIONS, pages):
    print(f"  {name}: {pg}")
