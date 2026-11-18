"""Slide: Renovation Plan (rebuilt from a teammate's draft; replaces the prior group's 'Renovation Plan').

Usage: python3 build_renovation.py <in.pptx> <out.pptx>

Numbers (teammate's draft unless noted):
  IC plan (Mar-23 IC model): 116 units renovated by Mar-25; $6.5K per 1 bed, $8.0K per 2 bed; rent lift ~$348/mo
    (renovation plus mark-to-market); average market rent $1,310 (IC presentation: sponsor U/W average rent $1,310);
    $1.8M capex reserve at closing (Control Sheet $1,803,172).
  Today: 12 units in renovated layouts, 4 occupied (box score 11/15/2026, verified); cost ~$13,500 per unit (average of
    the first completed units); rent lift ~$70/mo on signed renovated leases; average market rent of renovated
    layouts $1,202 (box score, verified); ~$340K left in capex escrow, ~$661K capitalized as building improvements
    (sponsor balance sheet).
  Flag: the IC unit mix (Mar-23) already coded 154 units as renovated layouts (A1R 10, A2R 55, A3R 21, A4R 15,
    B1R 31, B2R 22); the 11/15/2026 box score codes 12. Unit coding changed, so 12 may understate finished units.
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW, R1 = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"], g["R1"]
old = drop_slide_titled(prs, "Renovation Plan", layout="5_CONTENT")

PLAN, DONE = 116, 12
s = new_slide(prs, "Renovation Plan",
              "The plan called for 116 renovated units by Mar-25; 12 are done, at about twice the budgeted cost, "
              "for a fraction of the rent lift")
add_slide_number(s)

# ------------------------------------------------------------------ left: progress grid + economics
header(s, L0, g["ROW1"], LW, "Units Renovated vs. Plan", "A header", sup="(1)")
cols, sq, gap = 20, 0.3, 0.068
x0, y0 = L0 + 0.02, 1.86
for k in range(PLAN):
    r, c_ = divmod(k, cols)
    box(s, x0 + c_ * (sq + gap), y0 + r * (sq + gap), sq, sq, fill="ink" if k < DONE else "grey_lt", line=None,
        shape=MSO_SHAPE.RECTANGLE, name=f"A unit {k+1}")
yb = y0 + 6 * (sq + gap)
textbox(s, L0, yb + 0.02, LW, 0.26,
        [[("12 renovated", {"b": True}), ("   |   ", {"c": "grey_mid"}),
          ("104 of the 116 planned units never started", {"b": True, "c": "accent"})]], size=12, name="A grid label")

header(s, L0, 4.55, LW, "Renovation Economics, Plan vs. Today", "B header")


def pair(y, label, plan_v, act_v, plan_txt, act_txt, vmax, name):
    textbox(s, L0, y, 1.75, 0.5, [[(label, {"b": True})]], size=12, name=f"{name} label", anchor=MSO_ANCHOR.MIDDLE)
    bx, bw_max = L0 + 1.85, 4.3
    for j, (v, txt, fill, tc) in enumerate([(plan_v, plan_txt, "grey_lt", "ink"), (act_v, act_txt, "accent_fill",
                                                                                   "accent")]):
        w = bw_max * v / vmax
        yy = y + 0.03 + j * 0.24
        box(s, bx, yy, w, 0.2, fill=fill, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"{name} bar {j+1}")
        textbox(s, bx + w + 0.06, yy - 0.02, 1.4, 0.22, [[(txt, {"b": True, "c": tc})]], size=11.5,
                name=f"{name} val {j+1}")


pair(4.98, "Cost per unit", 7250, 13500, "$6.5K to $8.0K plan", "~$13.5K today", 14000, "B cost")
pair(5.56, "Rent lift per unit", 348, 70, "~$348/mo plan", "~$70/mo today", 360, "B lift")

# ------------------------------------------------------------------ right: plan vs. today table + why
header(s, R0, g["ROW1"], RW, "IC Plan vs. Today", "C header", sup="(2)")
rows = [
    ["Units renovated", "116 by Mar-25", "12"],
    ["Cost per unit", "$6.5K to $8.0K", "~$13.5K"],
    ["Rent lift per unit", "~$348/mo", "~$70/mo"],
    ["Avg. market rent", "$1,310", "$1,202"],
    ["Capex", "$1.8M reserve", "~$340K left"],
]
yt = grid_table(s, R0, 1.72, [1.75, 1.55, 1.33], ["", "IC plan", "Today"], rows, "C table", size=12, row_h=0.32,
                hair_rules=True, accent_cols={(k, 2) for k in range(5)})
header(s, R0, yt + 0.16, RW, "Why It Fell Short", "D header")
numbered_items(s, R0, yt + 0.62, RW, [
    [[("Sequencing: ", {"b": True}), ("exteriors first; interiors paused in the 2024 evictions", {})]],
    [[("Cost: ", {"b": True}), ("units came in at about double the budget", {})]],
    [[("Cash: ", {"b": True}), ("the 9.5% bridge loan absorbed the cash to keep going", {})]],
], [0.3, 0.3, 0.3], name="D item", size=12, num_d=0.24, gap=0.14)

closing(s, [("With cash going to debt service, ", {}),
            ("the rent premium that justified the business plan never arrived", {"b": True})])
footer_notes(s, "Notes: (1) Units in renovated layouts, box score 11/15/26 (4 of 12 occupied); the Mar-23 unit mix "
                "coded 154 units as renovated, so coding changed. (2) Today's cost and rent lift from early "
                "completed units and signed leases; ~$661K of capex capitalized.",
             "Sources: IC model and presentation (Mar/Apr-23); TexSun box score, leases and balance sheet.")
speaker_notes(s, """Presenter: TBD

The renovation plan was the engine of the business plan, and it never got going.

At IC, the plan was to renovate 116 units by March 2025, at about $6,500 for a one-bedroom and $8,000 for a two-bedroom, and earn about $348 a month more per unit, which supported an average rent of about $1,310. There was a $1.8 million reserve at closing to pay for it.

Today, 12 units sit in renovated floor plans, shown on the left as the dark squares out of the 116 that were planned. Those units cost about $13,500 each, roughly double the budget, and the signed leases on them show only about $70 a month of lift. Renovated units market at about $1,202, and only 4 of the 12 are occupied. About $340 thousand is left in the capex escrow.

Why: the sponsor did exterior and amenity work first, then paused interiors during the 2024 eviction cleanup. When units were done, they cost twice as much. And once the bridge loan closed at 9.5%, the cash that would have funded renovations went to debt service.

Without renovated units, the rent premium the whole plan depended on never showed up.

Reference notes:
- Teammate's numbers: 116-unit plan, $6.5K/$8.0K, $348 lift, $13.5K actual, $70 lift on signed leases, ~$340K left in escrow, ~$661K capitalized.
- Verified: $1,310 underwritten rent (IC deck); 12 renovated-layout units, 4 occupied, $1,202 average market rent (box score 11/15/2026).
- Caution: the March 2023 unit mix in the IC deck already listed 154 units with 'R' (renovated) layout codes, so the 12 reflects current coding and may understate finished units; confirm with TexSun.
- Asking-rent premium of renovated layouts over the same unrenovated floor plan in the Nov-26 box score: about $130 to $300.""")
place(prs, s, old if old is not None else 18)
prs.save(DST)
print("saved", DST)
