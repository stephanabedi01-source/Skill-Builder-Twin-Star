"""Slides: Market Overview (DFW) and Submarket (East Fort Worth and the comp set). Replace the prior group's
'Market Overview' (6_CONTENT) and 'Submarket Overview' (5_CONTENT) slides in place.

Usage: python3 build_market.py <in.pptx> <out.pptx>

DFW, Q3 2026 (web sources):
  Cushman & Wakefield MarketBeat Q3 2026: vacancy 9.4% (10.0% Q1 peak; 9.6% Q2); YTD net absorption 26,700+ units,
    ~5,600 above completions (-> ~21,100 completions); average effective rent $1,461, -2.0% YoY; U.S. vacancy 8.6%.
  Newmark Q3 2026: new supply ~42,000 units in 2024 vs. ~20,000 through Q3 2026; RealPage forecasts +2.4% rent
    growth over the next 12 months.
East Fort Worth submarket:
  Transwestern Q3 2025: 38 buildings, 8,068 units, 90.3% occupied, $1,088 average rent, -1.4% YoY.
  Northmarq 2025 Q4: East Fort Worth vacancy fell 190 bps in 2025.
Comp set (Taylor_Commons_-_Rent_Comps.xlsx, 'Rent Comps Initial vs. Now'):
  CoStar occupancy (2025 pull) for 17 comps, average 77.1%, median 76.8%.
  Comp median asking rent per SF by floor plan, 2022 vs. 2025, and Taylor Commons' 2025 target rent per SF.
Taylor Commons occupancy 74.9% (box score 11/15/2026).
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW, R1 = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"], g["R1"]
pos_m = drop_slide_titled(prs, "Market Overview", layout="6_CONTENT")
pos_s = drop_slide_titled(prs, "Submarket Overview", layout="5_CONTENT")


def kv_rows(s, y, rows, name, hl=()):
    for k, (a, b) in enumerate(rows):
        colr = "accent" if k in hl else "ink"
        textbox(s, R0, y, 3.1, 0.26, [[(a, {"c": colr})]], size=12, name=f"{name} k{k+1}", anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, R0 + 3.1, y, RW - 3.1, 0.26, [[(b, {"b": True, "c": colr})]], size=12, align=PP_ALIGN.RIGHT,
                name=f"{name} v{k+1}", anchor=MSO_ANCHOR.MIDDLE)
        y += 0.3
        if k < len(rows) - 1:
            line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"{name} rule {k+1}")
    return y


# ============================================================================ Market Overview (DFW)
s = new_slide(prs, "Market Overview: Dallas-Fort Worth",
              "DFW is absorbing more apartments than it delivers and vacancy is down to 9.4%, but effective rents "
              "are still 2.0% lower than a year ago")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "Vacancy Is Falling as Supply Drops", "A header", sup="(1)")
# chart A: vacancy by quarter
cats = ["Q1-26", "Q2-26", "Q3-26"]
vac = [10.0, 9.6, 9.4]
PA = Plot(0.75, 2.25, 3.0, 2.95, n=3, vmin=0, vmax=12)
caption(s, L0, 1.7, 3.4, "DFW apartment vacancy (%)", "A cap 1")
chart(s, PA, cats, [("Vacancy", vac)], [bar_ser(0, "Vacancy", cats, vac, "grey_lt", dpt={2: "ink"})],
      "A vacancy chart", gap=70)
line(s, PA.x, PA.v(0), PA.x + PA.w, PA.v(0), "grey_mid", 0.75, name="A base 1")
line(s, PA.x, PA.v(8.6), PA.x + PA.w, PA.v(8.6), "ink", 1.0, dash=MSO_LINE_DASH_STYLE.DASH, name="A US line")
textbox(s, PA.x + PA.w + 0.04, PA.v(8.6) - 0.11, 0.62, 0.22, [[("U.S. 8.6%", {"b": True})]], size=11,
        name="A US label")
for k, v in enumerate(vac):
    textbox(s, PA.center(k) - 0.45, PA.v(v) - 0.26, 0.9, 0.22, [[(f"{v:.1f}%", {"b": True})]], size=12,
            align=PP_ALIGN.CENTER, name=f"A vac label {k+1}")
    textbox(s, PA.center(k) - 0.45, PA.y + PA.h + 0.05, 0.9, 0.22, [[(cats[k], {"b": True})]], size=11,
            align=PP_ALIGN.CENTER, name=f"A vac cat {k+1}")
# chart B: supply vs. demand
cats2 = ["2024 built", "2026 YTD built", "2026 YTD leased"]
units = [42.0, 21.1, 26.7]
PB = Plot(4.6, 2.25, 3.15, 2.95, n=3, vmin=0, vmax=48)
caption(s, 4.25, 1.7, 3.6, "Apartment units (000s)", "A cap 2")
chart(s, PB, cats2, [("Units", units)], [bar_ser(0, "Units", cats2, units, "grey_lt", dpt={2: "ink"})],
      "A supply chart", gap=70)
line(s, PB.x, PB.v(0), PB.x + PB.w, PB.v(0), "grey_mid", 0.75, name="A base 2")
for k, v in enumerate(units):
    textbox(s, PB.center(k) - 0.5, PB.v(v) - 0.26, 1.0, 0.22, [[(f"~{v:.0f}K" if k < 2 else "26.7K+", {"b": True})]],
            size=12, align=PP_ALIGN.CENTER, name=f"A sup label {k+1}")
    textbox(s, PB.center(k) - 0.6, PB.y + PB.h + 0.05, 1.2, 0.42, [[(cats2[k], {"b": True})]], size=11,
            align=PP_ALIGN.CENTER, name=f"A sup cat {k+1}")

header(s, R0, g["ROW1"], RW, "DFW at a Glance, Q3 2026", "B header", sup="(2)")
y = kv_rows(s, 1.72, [("Vacancy (Q1 peak 10.0%)", "9.4%"), ("Average effective rent", "$1,461"),
                      ("Effective rent, year over year", "(2.0%)"), ("Net absorption, 2026 YTD", "26,700+"),
                      ("Absorption above completions", "~5,600"), ("Rent growth forecast, next 12 mo.", "+2.4%")],
            "B", hl=(2,))
callout(s, R0, y + 0.22, RW, 0.62, [("Supply has halved since 2024; ", {}),
                                    ("rent growth is forecast to turn positive", {"b": True})], "B callout", size=12)
header(s, R0, y + 1.06, RW, "Where the Recovery Shows Up", "C header")
numbered_items(s, R0, y + 1.5, RW, [
    [[("Rent gains: ", {"b": True}), ("five submarkets, led by core Dallas", {})]],
    [[("East Fort Worth: ", {"b": True}), ("older Class C still cutting rents", {})]]],
    [0.3, 0.3], name="C item", size=12, num_d=0.24, gap=0.1)
closing(s, [("The metro is turning, but annual rent gains so far sit in core Dallas submarkets, ", {}),
            ("not east Fort Worth", {"b": True})])
footer_notes(s, "Notes: (1) Vacancy and absorption per Cushman & Wakefield; 2026 YTD deliveries implied from "
                "absorption less the ~5,600-unit excess; 2024 deliveries per Newmark. (2) Forecast per RealPage via "
                "Newmark.",
             "Sources: Cushman & Wakefield MarketBeat DFW Multifamily Q3 2026; Newmark North Texas Multifamily Q3 "
             "2026.")
speaker_notes(s, """Presenter: TBD

Zooming out to the metro. Dallas-Fort Worth went through the biggest apartment supply wave in the country, and it is now working through it.

Vacancy peaked at 10% in the first quarter of this year and is down to 9.4% in the third quarter, still above the national 8.6%. The reason is on the right of the chart: about 42,000 units delivered in 2024, roughly half that so far this year, while renters absorbed more than 26,700 units, about 5,600 more than were completed. DFW is leading the country in absorption.

Rents have not caught up yet. Average effective rent is about $1,461, still 2% lower than a year ago, though that is the smallest decline since early 2024, and RealPage forecasts about 2.4% growth over the next year.

The catch for us: the submarkets with positive rent growth are core Dallas, like Uptown and Downtown. East Fort Worth, and especially the older product around Taylor Commons, is a different story, which is the next slide.

Reference notes:
- Cushman & Wakefield Q3 2026: vacancy 9.4% (10.0% Q1 peak); YTD absorption 26,700+ units, above full-year 2024 and 2025 totals and ~5,600 above completions; effective rent $1,461, -2.0% YoY; five submarkets with positive annual rent growth incl. Uptown Park Cities and Downtown Dallas.
- Newmark Q3 2026: average rent $1,512, occupancy 94.0% (different definition), annual rent change -1.3%; supply ~42,000 units in 2024 vs. ~20,000 through Q3 2026; RealPage +2.4% next 12 months.
- Rent figures differ by source because of definitions (effective vs. average vs. asking); this slide uses Cushman's effective rent.""")
place(prs, s, pos_m if pos_m is not None else 9)

# ============================================================================ Submarket and comp set
comps = [("Bella Via", 92.5), ("Luna Vista", 91.1), ("Trinity Heights", 89.6), ("The Ryan", 84.5),
         ("The Crosby", 83.1), ("Tides on Oakland Hills", 82.1), ("The Borough", 81.2), ("Cobble Hill", 80.9),
         ("The Brittany", 76.8), ("Casa Villa", 76.2), ("Caden", 74.5), ("Chapman (Tides at 5400)", 74.5),
         ("Dunmore", 72.1), ("500 Flats", 64.5), ("The Mason", 64.5), ("Tides at Meadowbrook", 61.7),
         ("Spur Canyon", 61.1)]
avg = sum(v for _, v in comps) / len(comps)
assert abs(avg - 77.1) < 0.05, avg
rows = [("Taylor Commons (Nov-26)", 74.9)] + comps
rows.sort(key=lambda r: -r[1])
s = new_slide(prs, "Submarket: East Fort Worth and the Comp Set",
              "East Fort Worth runs about 90% occupied, but Taylor Commons' older comp set averages 77% and comp "
              "rents fell 8% to 16% since IC")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "Comp Set Occupancy", "A header", sup="(1)")
cats = [r[0] for r in rows]
vals = [r[1] for r in rows]
tc_k = cats.index("Taylor Commons (Nov-26)")
P = Plot(2.6, 1.72, 4.6, 3.95, n=len(rows), vmin=50, vmax=100, horizontal=True)
chart(s, P, cats, [("Occupancy", vals)], [bar_ser(0, "Occupancy", cats, vals, "grey_lt", dpt={tc_k: "accent_fill"})],
      "A occupancy chart", gap=35)
for v, lab, nm in [(90.3, "Submarket 90.3%", "sub"), (avg, f"Comp avg. {avg:.1f}%", "avg")]:
    x = P.v(v)
    line(s, x, P.y - 0.05, x, P.y + P.h, "ink", 1.0, dash=MSO_LINE_DASH_STYLE.DASH, name=f"A {nm} line")
    textbox(s, x - 0.75, P.y + P.h + 0.03, 1.5, 0.22, [[(lab, {"b": True})]], size=11, align=PP_ALIGN.CENTER,
            name=f"A {nm} label")
for k, (nm, v) in enumerate(rows):
    tc = k == tc_k
    textbox(s, L0, P.center(k) - 0.11, 2.02, 0.22, [[(nm, {"b": tc, "c": "accent" if tc else "ink"})]], size=11,
            align=PP_ALIGN.RIGHT, name=f"A name {k+1}", anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, P.v(v) + 0.04, P.center(k) - 0.1, 0.5, 0.2, [[(f"{v:.1f}%", {"b": tc, "c": "accent" if tc else "ink"})]],
            size=11, name=f"A val {k+1}", anchor=MSO_ANCHOR.MIDDLE, fill="white")

header(s, R0, g["ROW1"], RW, "Comp Asking Rent per SF", "B header", sup="(2)")
rent = [("1 bed, 440-570 SF", 1.8717, 1.5673, 1.7035), ("1 bed, 560-700 SF", 1.6518, 1.3941, 1.3652),
        ("1 bed, 700+ SF", 1.4362, 1.3029, 1.4071), ("2 bed, 1 bath", 1.3997, 1.2841, 1.3826),
        ("2 bed, 2 bath", 1.366, 1.2333, 1.2143)]
trows = []
acc = set()
for k, (nm, a, b, t) in enumerate(rent):
    prem = t / b - 1
    trows.append([nm, f"${a:.2f}", f"${b:.2f}", f"{b / a - 1:.0%}".replace("-", "(") + (")" if b < a else ""),
                  (f"+{prem:.0%}" if prem > 0 else f"({abs(prem):.0%})")])
    acc.add((k, 3))
    if prem > 0:
        acc.add((k, 4))
yb = grid_table(s, R0, 1.72, [1.6, 0.68, 0.68, 0.72, 0.95], ["Floor plan", "2022", "2025", "Chg.", "TC vs. comps"],
                trows, "B table", size=11.5, row_h=0.31, hair_rules=True, accent_cols=acc, header_size=11.5)
header(s, R0, yb + 0.16, RW, "East Fort Worth Submarket", "C header", sup="(3)")
y = kv_rows(s, yb + 0.58, [("Occupancy, 8,068 units", "90.3%"), ("Average rent, year over year", "$1,088 (1.4%)"),
                           ("Vacancy change in 2025", "(190 bps)")], "C")
callout(s, R0, y + 0.2, RW, 0.62, [("The annual report's ", {}), ("~20% vacancy", {"b": True}),
                                   (" describes this older slice", {})], "C callout", size=12)
closing(s, [("The submarket is healing, but Taylor Commons competes in its weakest slice and ", {}),
            ("is priced above the comps on three of five floor plans", {"b": True})])
footer_notes(s, "Notes: (1) CoStar occupancy for comps (2025 pull in the team's comp file); Taylor Commons from the "
                "11/15/26 box score. (2) Comp median asking rent per SF; TC = 2025 target rent. (3) Transwestern "
                "Q3-25; Northmarq Q4-25.",
             "Sources: Team rent comp file (CoStar); TexSun box score; Transwestern; Northmarq; Fund annual report.")
speaker_notes(s, """Presenter: TBD

The east Fort Worth submarket as a whole is in decent shape: about 8,000 units, roughly 90% occupied in the latest Transwestern data, with vacancy falling almost two points last year.

But Taylor Commons does not compete with the whole submarket. It competes with the older garden properties around it, and that set looks very different. On the left are the 17 comps in our comp file with their CoStar occupancy. They average 77%, and several of the Tides-branded and older properties are in the low 60s. Taylor Commons, at 74.9% today, sits right in that group. That matches the annual report's comment that submarket vacancy is hovering around 20%: that is this older slice, not the whole submarket.

On the right, comp rents. Comparing the IC-era survey to the 2025 survey, median asking rent per square foot fell 8% to 16% depending on floor plan. The IC plan assumed rents would rise about 11%.

And Taylor Commons' own pricing is above the comps on three of the five floor plans: the small one-bedrooms by about 9%, the large one-bedrooms by 8%, and the two-bed one-baths by 8%. That is the pricing problem TexSun itself flagged, and it shows up in the 159 days new leases sat vacant.

Reference notes:
- Comp occupancy is CoStar data from the team's 2025 comp pull; Taylor Commons is the 11/15/2026 box score (74.9%). Comp average 77.1%, median 76.8%.
- Rent per SF 2022 vs. 2025 (comp median): A1 $1.87 to $1.57; A2 $1.65 to $1.39; A3/A4 $1.44 to $1.30; B1 $1.40 to $1.28; B2 $1.37 to $1.23. TC 2025 target: $1.70, $1.37, $1.41, $1.38, $1.21.
- Transwestern Q3 2025: East Fort Worth 38 buildings, 8,068 units, 90.3% occupied, $1,088, -1.4% YoY. Northmarq Q4 2025: vacancy down 190 bps in 2025.
- Apartments.com (Jan-26): East Fort Worth 1BR $983, 2BR $1,277, -3.6% YoY.""")
place(prs, s, (pos_m + 1) if pos_m is not None else 10)

prs.save(DST)
print("saved", DST)
