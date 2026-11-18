"""Slides: Market Overview (DFW) and Submarket Overview (East Fort Worth). Replace the prior group's 'Market
Overview' (6_CONTENT) and 'Submarket Overview' (5_CONTENT) slides, or earlier builds of these slides, in place.

Usage: python3 build_market.py <in.pptx> <out.pptx>

Sources (web; most report hosts are blocked here, so figures come via search results; each labeled on the slide):
  Economy: DFW nonfarm jobs ~4.36M in Aug-26, +1.0% YoY; unemployment 4.4% Aug-26 (BLS via Dallas Fed / Dallas
    Express); Texas 2026 job forecast +0.9% (Dallas Fed); City of Fort Worth +19,512 residents 2024-25 (Census).
  Supply and demand: 2024 completions 41,650 and net absorption 21,024 (MMG); 2026 YTD through Q3 absorption 26,700+
    and ~5,600 above completions (-> ~21,100 completions) (Cushman & Wakefield Q3-26); 24,133 units scheduled for
    the next 12 months and 43,320 under construction (Colliers Q2-26).
  Fundamentals: vacancy 9.4% Q3-26, 10.0% Q1-26 peak; effective rent $1,461, -2.0% YoY (Cushman & Wakefield);
    RealPage +2.4% rent growth next 12 months (via Newmark).
  East Fort Worth (Transwestern factsheets): Q3-25 38 bldgs / 8,068 units, 90.3% occupied, $1,088, -1.4% YoY;
    Q1-26 39 bldgs / 8,359 units, 91.3% occupied, 252 units absorbed, $1,050, -4.6% YoY, 200 under construction.
"""
import sys

from PIL import Image

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW, R1 = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"], g["R1"]
pos_m = drop_slide_titled(prs, "Market Overview", layout="6_CONTENT")
pos_s = drop_slide_titled(prs, "Submarket Overview", layout="5_CONTENT")
old_sub = drop_slide_titled(prs, "Submarket: East Fort Worth and the Comp Set", layout="CONTENT")
IMG = os.path.join(HERE, "img")


def kv_rows(s, y, rows, name, hl=(), split=3.1):
    for k, (a, b) in enumerate(rows):
        colr = "accent" if k in hl else "ink"
        textbox(s, R0, y, split, 0.26, [[(a, {"c": colr})]], size=12, name=f"{name} k{k+1}", anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, R0 + split, y, RW - split, 0.26, [[(b, {"b": True, "c": colr})]], size=12, align=PP_ALIGN.RIGHT,
                name=f"{name} v{k+1}", anchor=MSO_ANCHOR.MIDDLE)
        y += 0.3
        if k < len(rows) - 1:
            line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"{name} rule {k+1}")
    return y


# ============================================================================ Market Overview (DFW)
s = new_slide(prs, "Market Overview: Dallas-Fort Worth",
              "DFW demand now outruns new supply and the pipeline keeps shrinking, but rents are still 2.0% lower "
              "than a year ago")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "The Supply Wave Is Passing (000s of units)", "A header", sup="(1)")
for k, (lab, colr) in enumerate([("Completions", "grey_lt"), ("Net absorption (demand)", "ink")]):
    lx = L0 + 2.6 + k * 1.75
    box(s, lx, 1.80, 0.14, 0.14, fill=colr, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"A legend key {k+1}")
    textbox(s, lx + 0.2, 1.76, 1.9, 0.22, [[(lab, {})]], size=11, name=f"A legend {k+1}")
cats = ["2024", "2026 through Q3", "Next 12 months"]
comp = [41.7, 21.1, 24.1]
absn = [21.0, 26.7, None]
P = Plot(0.8, 2.25, 6.8, 3.05, n=3, vmin=0, vmax=46)
chart(s, P, cats, [("Completions", comp), ("Net absorption", absn)],
      [bar_ser(0, "Completions", cats, comp, "grey_lt", dpt={2: (None, "grey_mid", "dash")}),
       bar_ser(1, "Net absorption", cats, absn, "ink")],
      "A supply chart", grouping="clustered", gap=80, overlap=-8)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="A baseline")
bw = P.slot / (2 + 0.08 + 0.8)
for k in range(3):
    for j, v in enumerate([comp[k], absn[k]]):
        if v is None:
            continue
        bx = P.center(k) + (j - 0.5) * bw * 1.08
        txt = f"{v:.1f}" if not (k == 1 and j == 1) else "26.7+"
        textbox(s, bx - 0.45, P.v(v) - 0.25, 0.9, 0.22, [[(txt, {"b": True})]], size=12, align=PP_ALIGN.CENTER,
                name=f"A label {k+1}-{j+1}")
sub = ["Record supply; demand about half", "Demand above completions", "Scheduled; 43.3K under way"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, P.center(k) - 1.1, P.y + P.h + 0.06, 2.2, 0.44, [[(cat, {"b": True})], [(sb, {"c": "grey_txt", "i": True})]],
            size=11, align=PP_ALIGN.CENTER, name=f"A category {k+1}")
pill(s, P.center(0) + 0.15, P.v(21.0) - 0.75, 1.75, 0.32, [("2x supply vs. demand", {"b": True})], "A pill 2024",
     line_color="ink", size=11)

header(s, R0, g["ROW1"], RW, "Economy", "B header", sup="(2)")
y = kv_rows(s, 1.72, [("Nonfarm jobs (Aug-26)", "4.36M"), ("Job growth, year over year", "+1.0%"),
                      ("Unemployment (Aug-26)", "4.4%"), ("City of Fort Worth, 2024-25", "+19,512 people")], "B",
            split=2.6)
header(s, R0, y + 0.14, RW, "Apartment Fundamentals, Q3 2026", "C header", sup="(3)")
y = kv_rows(s, y + 0.56, [("Vacancy (Q1 peak 10.0%)", "9.4%"), ("Average effective rent", "$1,461"),
                          ("Effective rent, year over year", "(2.0%)"), ("Rent forecast, next 12 months", "+2.4%")],
            "C", hl=(2,), split=3.1)
callout(s, R0, y + 0.18, RW, 0.62, [("Rent gains so far are in ", {}), ("Class A and core Dallas", {"b": True}),
                                    (", not older Class C", {})], "C callout", size=12)
closing(s, [("Supply is clearing across DFW, which helps from 2027, ", {}),
            ("but rent growth has not returned yet, least of all in older product", {"b": True})])
footer_notes(s, "Notes: (1) 2024 per MMG; 2026 per Cushman & Wakefield (completions implied from absorption less the "
                "~5,600-unit excess); next 12 months and under construction per Colliers Q2-26. (2) BLS via Dallas "
                "Fed; Census. (3) Cushman & Wakefield; forecast per RealPage.",
             "Sources: MMG; Cushman & Wakefield; Colliers; Newmark; Dallas Fed; U.S. Census Bureau.")
speaker_notes(s, """Presenter: TBD

Dallas-Fort Worth is still one of the strongest job and population markets in the country: about 4.4 million jobs, growing about 1% a year, unemployment around 4.4%, and the city of Fort Worth adding almost 20,000 people last year.

The apartment story is a supply story. In 2024 the metro delivered more than 41,000 new units, the most on record, and demand absorbed only about half of that. That is what pushed vacancy up and rents down. This year it flipped: through the third quarter renters absorbed more than 26,700 units against roughly 21,000 completed, and only about 24,000 are scheduled over the next twelve months.

Vacancy peaked at 10% early this year and is down to 9.4%. Rents have not recovered yet: average effective rent is about $1,461, still 2% below a year ago, though RealPage forecasts about 2.4% growth over the next year.

For Taylor Commons the timing matters: the market tailwind arrives in 2027, the same year the tax exemption ends and the loan matures, and it shows up first in newer and core Dallas product.

Reference notes:
- MMG: 41,650 completions and 21,024 net absorption in 2024. Cushman & Wakefield Q3-26: 26,700+ absorbed YTD, ~5,600 above completions, vacancy 9.4%, effective rent $1,461 (-2.0% YoY). Colliers Q2-26: 43,320 under construction, 24,133 due in 12 months.
- Sources define vacancy and rent differently (e.g., Northmarq and Matthews quote 6% to 7% vacancy on a stabilized basis); this slide uses Cushman & Wakefield throughout for fundamentals.
- DFW nonfarm jobs ~4.36M (Aug-26), +1.0% YoY; unemployment 4.4%; Dallas Fed 2026 Texas job forecast +0.9%.""")
place(prs, s, pos_m if pos_m is not None else 9)

# ============================================================================ Submarket Overview (East Fort Worth)
s = new_slide(prs, "Submarket Overview: East Fort Worth",
              "East Fort Worth is 91% occupied with only 200 units under construction, but its rents are falling "
              "faster than the metro's")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "East Fort Worth Submarket (Illustrative)", "A header")
SAT = Image.open(os.path.join(IMG, "sat_map.jpg" if os.path.exists(os.path.join(IMG, "sat_map.jpg")) else "sat_map.png")).convert("RGB")
CROP, SX0, SY0, SW, SH = (0.00803, 0.07403, 0.07014, 0.07773), 0.72, 1.3, 11.09, 4.9


def to_px(x, y):
    cl, cr, ct, cb = CROP
    return ((cl + (x - SX0) / SW * (1 - cl - cr)) * SAT.size[0], (ct + (y - SY0) / SH * (1 - ct - cb)) * SAT.size[1])


MX, MY, MW, MH = L0, 1.72, LW, 2.72
ax0, ay0 = to_px(2.6, 1.95)
ax1, ay1 = to_px(11.2, 6.2)
SAT.crop((int(ax0), int(ay0), int(ax1), int(ay1))).resize((1600, int(1600 * MH / MW))).save(
    os.path.join(IMG, "sat_sub.jpg"), quality=88)
add_image(s, os.path.join(IMG, "sat_sub.jpg"), MX, MY, MW, MH, name="A submarket map")


def m(x, y):
    px, py = to_px(x, y)
    return MX + (px - ax0) / (ax1 - ax0) * MW, MY + (py - ay0) / (ay1 - ay0) * MH


e0, e1 = m(4.4, 2.6), m(10.6, 5.75)
box(s, e0[0], e0[1], e1[0] - e0[0], e1[1] - e0[1], fill=None, line="white", dash=MSO_LINE_DASH_STYLE.DASH,
    line_w=2.0, shape=MSO_SHAPE.OVAL, name="A submarket outline")
lab = box(s, e0[0] + 0.35, e0[1] - 0.05, 2.25, 0.3, fill="ink", line=None, name="A submarket label")
lab.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(lab.text_frame, [[("East Fort Worth (approx.)", {"b": True, "c": "white"})]], 11, align=PP_ALIGN.CENTER)
tx, ty = m(6.45, 3.87)
box(s, tx - 0.12, ty - 0.12, 0.24, 0.24, fill="accent_fill", line="white", line_w=1.5, shape=MSO_SHAPE.OVAL,
    name="A subject dot")
tl = box(s, tx + 0.18, ty - 0.16, 1.5, 0.3, fill="accent_fill", line=None, name="A subject label")
tl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(tl.text_frame, [[("Taylor Commons", {"b": True, "c": "white"})]], 11, align=PP_ALIGN.CENTER)
dx, dy = m(2.9, 4.45)
dl = box(s, max(MX + 0.08, dx - 0.2), dy - 0.15, 1.75, 0.3, fill="white", line=None, name="A downtown label")
dl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(dl.text_frame, [[("To downtown, 11 min", {"b": True})]], 11, align=PP_ALIGN.CENTER)

# rent and occupancy vs. DFW
header(s, L0, 4.6, LW, "East Fort Worth vs. DFW", "B header", sup="(1)")


def pair(y, label, items, vmax, name, fmt):
    textbox(s, L0, y, 1.75, 0.48, [[(label, {"b": True})]], size=12, name=f"{name} label", anchor=MSO_ANCHOR.MIDDLE)
    bx, bw_max = L0 + 1.85, 4.2
    for j, (nm, v, fill, tc) in enumerate(items):
        w = bw_max * abs(v) / vmax
        yy = y + 0.02 + j * 0.24
        box(s, bx, yy, w, 0.2, fill=fill, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"{name} bar {j+1}")
        textbox(s, bx + w + 0.06, yy - 0.02, 2.0, 0.22, [[(fmt(v), {"b": True, "c": tc}), (f"  {nm}", {"c": "grey_txt"})]],
                size=11.5, name=f"{name} val {j+1}")


pair(5.5, "Occupancy", [("East Fort Worth", 91.3, "grey_lt", "ink"), ("DFW", 90.6, "grey_lt", "ink")],
     100.0, "B occ", lambda v: f"{v:.1f}%")
pair(5.0, "Rent, year over year", [("East Fort Worth", -4.6, "accent_fill", "accent"), ("DFW", -2.0, "grey_lt", "ink")],
     5.0, "B rent", lambda v: f"({abs(v):.1f}%)")

# right: submarket stats
header(s, R0, g["ROW1"], RW, "Submarket Snapshot", "C header", sup="(2)")
rows = [
    ["Buildings / units", "38 / 8,068", "39 / 8,359"],
    ["Occupancy", "90.3%", "91.3%"],
    ["Average rent", "$1,088", "$1,050"],
    ["Rent, year over year", "(1.4%)", "(4.6%)"],
    ["Units absorbed (quarter)", "n/a", "252"],
    ["Under construction", "n/a", "200"],
]
yt = grid_table(s, R0, 1.72, [2.05, 1.24, 1.34], ["", "Q3-25", "Q1-26"], rows, "C table", size=12, row_h=0.32,
                hair_rules=True, accent_cols={(2, 2), (3, 2)})
callout(s, R0, yt + 0.2, RW, 0.66, [("Rents are ", {}), ("down 4.6%", {"b": True}),
                                    (", over twice the metro's decline", {})], "C callout", size=12)
header(s, R0, yt + 1.1, RW, "Taylor Commons in the Submarket", "D header")
kv_rows(s, yt + 1.52, [("Occupancy (Nov-26)", "74.9%"), ("Market rent (Nov-26)", "$970")], "D", hl=(0,), split=3.0)
closing(s, [("Supply is not the problem here; ", {}),
            ("falling submarket rents cap what an older asset can charge (comps next)", {"b": True})])
footer_notes(s, "Notes: (1) Submarket outline approximate; drive time per prior team map; East Fort Worth per "
                "Transwestern Q1-26, DFW per Cushman & Wakefield Q3-26 (occupancy = 1 - vacancy). (2) Transwestern factsheets Q3-25 and Q1-26.",
             "Sources: Transwestern; Cushman & Wakefield; TexSun box score 11/15/26.")
speaker_notes(s, """Presenter: TBD

East Fort Worth is the submarket Taylor Commons competes in: the older, workforce side of the city east of downtown along I-30, with about 8,400 apartment units in 39 properties.

Supply is not the issue here. Occupancy is about 91% and has improved over the last few quarters, there were 252 units absorbed in the first quarter, and only about 200 units are under construction. Nobody is building new product on this side of town at scale.

The issue is rent. Average rent is about $1,050, and it is down 4.6% from a year ago, more than twice the metro's 2% decline. So this submarket is filling units by cutting price.

Taylor Commons sits below the submarket on both counts: 74.9% occupied against about 91% for the submarket, and $970 market rent against about $1,050. That is the setup for the comps: the competition is filling up by pricing down, and Taylor Commons has not kept up.

Reference notes:
- Transwestern Q3-25: 38 buildings, 8,068 units, 90.3% occupied, $1,088, -1.4% YoY. Q1-26: 39 buildings, 8,359 units, 91.3%, 252 units absorbed, $1,050, -4.6% YoY, 200 under construction. Confirm Q1-26 figures against the PDF; the extract was partly garbled.
- Northmarq: East Fort Worth vacancy fell 190 bps in 2025.
- The older comp set around Taylor Commons runs much weaker than the submarket average: ~77% occupied in the team's CoStar pull (comps section).""")
place(prs, s, (pos_m + 1) if pos_m is not None else (old_sub if old_sub is not None else 10))

prs.save(DST)
print("saved", DST)
