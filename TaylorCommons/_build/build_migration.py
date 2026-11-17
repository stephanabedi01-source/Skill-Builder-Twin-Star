"""Slide: Migration Trends (replaces 'Micro-migration Trends').

Usage: python3 build_migration.py <in.pptx> <out.pptx>

Figures (web sources; Census hosts are blocked in this environment, so each is from a published secondary page):
  Tarrant County, July 2024 -> July 2025 (Census Vintage 2025 components, as published by USAFacts and a
    co-est2025-alldata extract): total +18,428; natural change +12,445; net international ~+12,800;
    net domestic -6,978.  Population 2,248,466 (7/1/2025), +6.5% vs. April 2020 base (Census QuickFacts).
  City of Fort Worth: +23,442 (7/2023 -> 7/2024) to 1,008,106, 9.7% growth since 2020, 11th largest (Census,
    May 2025 release, per Community Impact / Fort Worth Report); +19,512 (2024 -> 2025), 10th largest (Dallas
    Innovates / Fort Worth Inc.).
  ZIP 76112 (ACS 2019-2023 5-year, via censusdepth.com and zip-codes.com): median household income $56,179;
    renter-occupied ~51% of occupied units; median gross rent $1,215.
  Taylor Commons market rent $970 (box score 11/15/2026); fee package $175 (site visit).
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"]
idx = drop_slide_titled(prs, "Micro-migration Trends")

NAT, INTL, DOM = 12.445, 12.8, -6.978
TOT = 18.428
assert abs(NAT + INTL + DOM - TOT) < 0.25   # components sum to the total within rounding

s = new_slide(prs, "Migration Trends",
              "Fort Worth still adds about 20,000 people a year, but Tarrant County now loses domestic movers; growth "
              "comes from births and immigration")
add_slide_number(s)

# ------------------------------------------------------------------ left: illustrative map
from PIL import Image  # noqa: E402

IMG = os.path.join(HERE, "img")
DFW = Image.open(os.path.join(IMG, "dfw_map.png")).convert("RGB")
CROP = (0.01386, 0.01474, 0.14142, 0.13362)        # prior slide's crop of the DFW map
SX0, SY0, SW, SH = 0.67, 1.3, 11.98, 4.9            # where the prior slide showed it


def to_px(x, y):
    cl, cr, ct, cb = CROP
    return ((cl + (x - SX0) / SW * (1 - cl - cr)) * DFW.size[0], (ct + (y - SY0) / SH * (1 - ct - cb)) * DFW.size[1])


MX, MY, MW, MH = L0, 1.72, LW, 4.25
ax0, ay0 = to_px(1.25, 1.55)
ax1, ay1 = to_px(9.30, 6.20)
DFW.crop((int(ax0), int(ay0), int(ax1), int(ay1))).save(os.path.join(IMG, "dfw_mig.jpg"), quality=90)


def m(x, y):
    px, py = to_px(x, y)
    return MX + (px - ax0) / (ax1 - ax0) * MW, MY + (py - ay0) / (ay1 - ay0) * MH


header(s, L0, g["ROW1"], LW, "Who Is Moving In and Out (Illustrative)", "A header", sup="(1)")
add_image(s, os.path.join(IMG, "dfw_mig.jpg"), MX, MY, MW, MH, name="A DFW map")
# Tarrant County, drawn as an approximate dashed outline
ex0, ey0 = m(2.80, 2.62)
ex1, ey1 = m(6.40, 5.45)
box(s, ex0, ey0, ex1 - ex0, ey1 - ey0, fill=None, line="ink", dash=MSO_LINE_DASH_STYLE.DASH, line_w=1.5,
    shape=MSO_SHAPE.OVAL, name="A Tarrant outline")
lx, ly = m(5.05, 2.62)
textbox(s, lx, ly - 0.12, 1.28, 0.24, [[("Tarrant County", {"b": True})]], size=12, name="A Tarrant label",
        fill="white")
# outflow arrows (brick) and inflow arrow (ink)
cx, cy = m(4.40, 4.03)
for nm, (x0, y0), (x1, y1) in [("NW", (3.55, 3.30), (2.20, 2.05)), ("SW", (3.55, 4.85), (2.20, 6.00)),
                               ("N", (4.40, 2.85), (4.40, 1.70))]:
    p0, p1 = m(x0, y0), m(x1, y1)
    arrow(s, p0[0], p0[1], p1[0], p1[1], "accent", 2.25, name=f"A out {nm}")
ox, oy = m(1.30, 1.62)
textbox(s, ox, oy, 1.98, 0.48, [[("(7.0K) net moves", {"b": True, "c": "accent"})],
                                [("to other U.S. counties", {"c": "accent"})]], size=12, name="A out label",
        fill="white")
i0, i1 = m(9.05, 2.55), m(6.05, 3.62)
arrow(s, i0[0], i0[1], i1[0], i1[1], "ink", 2.25, name="A in arrow")
ix, iy = m(7.55, 1.62)
textbox(s, ix, iy, 1.85, 0.48, [[("+12.8K net arrivals", {"b": True})], [("from abroad", {})]], size=12,
        name="A in label", fill="white")
# Fort Worth badge and subject
tx, ty = m(4.815, 4.015)
box(s, tx - 0.11, ty - 0.11, 0.22, 0.22, fill="accent_fill", line="white", line_w=1.25, shape=MSO_SHAPE.OVAL,
    name="A subject dot")
textbox(s, tx + 0.16, ty - 0.13, 1.36, 0.26, [[("Taylor Commons", {"b": True, "c": "accent"})]], size=12,
        name="A subject label", fill="white")
bx, by = m(3.75, 5.62)
b = box(s, bx, by, 2.6, 0.5, fill="white", line="ink", line_w=0.75, name="A FW badge")
b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(b.text_frame, [[("City of Fort Worth: ", {}), ("+19.5K", {"b": True})], [("residents, 2024 to 2025", {})]],
        12, align=PP_ALIGN.CENTER)

# ------------------------------------------------------------------ right: components + local facts
header(s, R0, g["ROW1"], RW, "Tarrant County, Jul-24 to Jul-25 (000s)", "B header")
cats = ["Births", "Abroad", "Domestic", "Net"]
base = [0, NAT, NAT + INTL + DOM, 0]
pos = [NAT, INTL, 0, TOT]
neg = [0, 0, -DOM, 0]
P = Plot(R0 + 0.1, 2.02, RW - 0.2, 1.55, n=4, vmin=0, vmax=28)
chart(s, P, cats, [("Base", base), ("Up", pos), ("Down", neg)],
      [nofill_ser(0, "Base", cats, base),
       bar_ser(1, "Up", cats, pos, "grey_lt", dpt={3: "ink"}),
       bar_ser(2, "Down", cats, neg, "accent_fill")],
      "B components chart", grouping="stacked", gap=60)
bw = P.bar_width(60)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="B baseline")
for k, lvl in [(0, NAT), (1, NAT + INTL), (2, NAT + INTL + DOM)]:
    line(s, P.center(k) + bw / 2, P.v(lvl), P.center(k + 1) - bw / 2, P.v(lvl), "grey_mid", 0.75,
         dash=MSO_LINE_DASH_STYLE.DASH, name=f"B connector {k+1}")
for k, top, txt, colr in [(0, NAT, "+12.4", "ink"), (1, NAT + INTL, "+12.8", "ink"),
                          (2, NAT + INTL, "(7.0)", "accent"), (3, TOT, "+18.4", "ink")]:
    textbox(s, P.center(k) - 0.5, P.v(top) - 0.25, 1.0, 0.22, [[(txt, {"b": True, "c": colr})]], size=12,
            align=PP_ALIGN.CENTER, name=f"B label {k+1}")
    textbox(s, P.center(k) - 0.55, P.y + P.h + 0.04, 1.1, 0.22, [[(cats[k], {"b": True})]], size=11,
            align=PP_ALIGN.CENTER, name=f"B category {k+1}")

header(s, R0, 4.02, RW, "Fort Worth and ZIP 76112", "C header", sup="(2)")
rows = [
    ("Fort Worth growth since 2020", "+9.7%", False),
    ("ZIP median household income", "$56,179", False),
    ("ZIP homes that are rented", "~51%", False),
    ("ZIP median gross rent", "$1,215", False),
    ("Taylor Commons market rent", "$970", True),
]
y = 4.42
for k, (a_, b_, hl) in enumerate(rows):
    colr = "accent" if hl else "ink"
    textbox(s, R0, y, 3.2, 0.26, [[(a_, {"b": hl, "c": colr})]], size=12, name=f"C k{k+1}",
            anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, R0 + 3.2, y, RW - 3.2, 0.26, [[(b_, {"b": True, "c": colr})]], size=12, align=PP_ALIGN.RIGHT,
            name=f"C v{k+1}", anchor=MSO_ANCHOR.MIDDLE)
    y += 0.3
    if k < len(rows) - 1:
        line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"C rule {k+1}")

closing(s, [("Renter demand is there, but not from new domestic arrivals: ", {}),
            ("Taylor Commons must win local renters on price and condition", {"b": True})])
footer_notes(s, "Notes: (1) Illustrative: county outline approximate; arrows show direction, not destinations. Census Vintage 2025 components via USAFacts, rounded. (2) Census city estimates; ACS 2019-2023, gross rent includes utilities.",
             "Sources: U.S. Census Bureau (PEP, ACS) via USAFacts, Community Impact, Dallas Innovates, "
             "censusdepth.com; box score 11/15/26.")
speaker_notes(s, """Presenter: TBD

Are people moving into this area or away? The answer is both, depending on the level.

The city of Fort Worth is one of the fastest-growing big cities in the country. It added about 23,000 people in 2024 and another 19,500 in 2025, passed one million residents, and is now the tenth-largest city in the U.S. It has grown almost 10% since 2020.

But look at the county on the left. Tarrant County grew by about 18,400 people last year, and all of that came from births and from people arriving from abroad. On net, about 7,000 more people moved out of Tarrant County to other U.S. counties than moved in. Domestic movers are going to the outer counties, not into the core.

Around the property, ZIP 76112 is about half renters, with a median household income of about $56 thousand. The median gross rent there is about $1,215 including utilities, and Taylor Commons markets at $970. So rents are not out of reach for local incomes.

The takeaway: there is renter demand here, but it is local, existing renters, not a wave of new people moving in. Taylor Commons has to win those renters on price and condition, which is exactly where the site visit and the operations data say it is falling short.

Reference notes:
- Tarrant County 7/1/2025: 2,248,466, up 6.5% from the April 2020 base (Census QuickFacts, Vintage 2025).
- Components Jul-24 to Jul-25: natural +12,445; net international about +12,800; net domestic -6,978; total +18,428.
- A Census demographer noted Tarrant also had some net domestic out-migration in 2023 to 2024, offset by international migration.
- ZIP figures come from secondary sites quoting ACS 2019-2023; confirm on data.census.gov (table S1901, B25003, B25064) if asked.""")
place(prs, s, idx if idx is not None else 8)
prs.save(DST)
print("saved", DST)
