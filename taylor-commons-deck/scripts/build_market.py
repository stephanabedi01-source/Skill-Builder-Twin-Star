"""Slides: Market Overview (DFW) and Submarket Overview (East Fort Worth).

Built on the prior team's slides: same modules (Market: population and economy, employment and income, apartment
fundamentals, investor confidence; Submarket: supply pipeline, rent dynamics, key employers with their logos),
with their placeholder, stale and other-deal content (Arlington / Mid-Cities pipeline, 'asdasd' text) replaced by
current, sourced East Fort Worth and DFW figures.

Usage: python3 build_market.py <in.pptx> <out.pptx>

Sources (web; report hosts blocked here, figures via search results):
  Population: DFW metro +123,557 (Jul-24 to Jul-25), No. 2 U.S. metro by numeric gain, ~8.5M; ~270K net domestic
    migrants 2020-25, most of any metro (Census Vintage 2025 / Census story May-26). Fort Worth +19,512, 10th largest.
  Jobs: ~4.36M nonfarm, +1.0% YoY, unemployment 4.4% (Aug-26, BLS via Dallas Fed); U.S. unemployment 4.2% (Sep-26).
  Income: DFW median household income ~$92.7K vs. U.S. ~$83.7K (2024; Social Explorer ACS / SmartAsset; confirm B19013).
  Supply/demand: 2024 completions 41,650 vs. absorption 21,024 (MMG); 2026 YTD Q3 absorption 26,700+, ~5,600 above
    completions (Cushman & Wakefield Q3-26); vacancy 9.4%, effective rent $1,461, -2.0% YoY (C&W).
  Investor confidence (Matthews Q3-26, CoStar): TTM sales volume ~$10.8B, +3% YoY; average cap rate 5.66%; $171K per unit,
    12% below the Q3-22 peak of $193K. Northmarq: 2025 volume +3%, median $172.3K/unit.
  East Fort Worth: Transwestern Q3-25 (38 bldgs, 8,068 units, 90.3%, $1,088, -1.4%) and Q1-26 (39 bldgs, 8,359 units,
    91.3%, 252 absorbed, $1,050, -4.6%, 200 under construction); GREA winter 2025: 666 under construction, 0 delivering
    in the next four quarters; Matthews 2024: ~670 under way, most in 20 years; Apartments.com Jan-26: 1BR $983,
    2BR $1,277, -3.6% YoY; ZIP 76112 ~51% renter-occupied (ACS 2019-23).
  Key employers (prior team's logos): American Airlines HQ, UT Arlington, GM Arlington Assembly; IC deck: residents'
    top employers Walmart, Amazon, FedEx.
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, R1 = g["L0"], g["R1"]
C1, CW = L0, 6.0                 # left module column
C2 = 6.83                        # right module column
CW2 = R1 - C2
IMG = os.path.join(HERE, "img")
pos_m = drop_slide_titled(prs, "Market Overview", layout="6_CONTENT")
pos_s = drop_slide_titled(prs, "Submarket Overview", layout="5_CONTENT")
old_sub = drop_slide_titled(prs, "Submarket: East Fort Worth and the Comp Set", layout="CONTENT")


def stat_rows(s, x, y, w, rows, name, num_w=1.25, size=12):
    """Prior team's 'big number + sentence' callouts, restyled: bold number left, sentence right, hairlines."""
    for k, (num, txt, colr) in enumerate(rows):
        textbox(s, x, y, num_w, 0.42, [[(num, {"b": True, "c": colr, "size": 17})]], size=17,
                anchor=MSO_ANCHOR.MIDDLE, name=f"{name} num {k+1}")
        textbox(s, x + num_w + 0.1, y, w - num_w - 0.1, 0.42, [[(txt, {})]], size=size, anchor=MSO_ANCHOR.MIDDLE,
                name=f"{name} txt {k+1}")
        y += 0.5
        if k < len(rows) - 1:
            line(s, x, y - 0.04, x + w, y - 0.04, "hair", 0.5, name=f"{name} rule {k+1}")
    return y


def hbar_pair(s, x, y, label, items, vmax, name, fmt, bar_w=2.6):
    textbox(s, x, y, 1.55, 0.46, [[(label, {"b": True})]], size=12, anchor=MSO_ANCHOR.MIDDLE, name=f"{name} label")
    for j, (nm, v, fill, tc) in enumerate(items):
        w = bar_w * abs(v) / vmax
        yy = y + 0.02 + j * 0.23
        box(s, x + 1.6, yy, w, 0.19, fill=fill, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"{name} bar {j+1}")
        textbox(s, x + 1.6 + w + 0.05, yy - 0.02, 1.7, 0.22, [[(fmt(v), {"b": True, "c": tc}), (" " + nm, {"c": "grey_txt"})]],
                size=11, name=f"{name} val {j+1}")


# ============================================================================ Market Overview
s = new_slide(prs, "Market Overview: Dallas-Fort Worth",
              "DFW keeps adding people and jobs and demand now outruns supply, but rents and sale prices are still "
              "below their peaks")
add_slide_number(s)

# --- module 1: population and economy (prior team's stat callouts, corrected)
header(s, C1, g["ROW1"], CW, "Population and Economy", "A header", sup="(1)")
stat_rows(s, C1, 1.7, CW, [
    ("No. 2", "U.S. metro for population growth: +123,557 people in 2025", "ink"),
    ("~270K", "Net domestic movers into DFW since 2020, the most of any metro", "ink"),
    ("4.36M", "Jobs, +1.0% year over year; unemployment 4.4% (Aug-26)", "ink"),
], "A")

# --- module 2: employment and income vs. U.S. (prior team's chart, with verified figures)
header(s, C2, g["ROW1"], CW2, "Employment and Income vs. U.S.", "B header", sup="(2)")
hbar_pair(s, C2, 1.78, "Unemployment", [("DFW", 4.4, "ink", "ink"), ("U.S.", 4.2, "grey_lt", "ink")], 5.0, "B unemp",
          lambda v: f"{v:.1f}%")
hbar_pair(s, C2, 2.36, "Median income", [("DFW", 92.7, "ink", "ink"), ("U.S.", 83.7, "grey_lt", "ink")], 100.0,
          "B inc", lambda v: f"${v:.1f}K")
textbox(s, C2, 2.95, CW2, 0.5, [[("Incomes run ~11% above the U.S.; ", {}),
                                 ("Fort Worth is now the 10th-largest U.S. city", {"b": True})]],
        size=12, name="B note")

# --- module 3: apartment fundamentals (supply vs. demand)
header(s, C1, 3.68, CW, "Apartment Fundamentals (000s of units)", "C header", sup="(3)")
cats = ["2024", "2026 through Q3"]
comp, absn = [41.7, 21.1], [21.0, 26.7]
P = Plot(C1 + 0.15, 4.42, 2.95, 1.2, n=2, vmin=0, vmax=46)
chart(s, P, cats, [("Completions", comp), ("Net absorption", absn)],
      [bar_ser(0, "Completions", cats, comp, "grey_lt"), bar_ser(1, "Net absorption", cats, absn, "ink")],
      "C supply chart", grouping="clustered", gap=70, overlap=-8)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="C baseline")
bw = P.slot / (2 + 0.08 + 0.7)
for k in range(2):
    for j, v in enumerate([comp[k], absn[k]]):
        bx = P.center(k) + (j - 0.5) * bw * 1.08
        textbox(s, bx - 0.4, P.v(v) - 0.24, 0.8, 0.22, [[(f"{v:.1f}" + ("+" if (k, j) == (1, 1) else ""), {"b": True})]],
                size=11, align=PP_ALIGN.CENTER, name=f"C label {k+1}-{j+1}")
    textbox(s, P.center(k) - 0.7, P.y + P.h + 0.04, 1.4, 0.22, [[(cats[k], {"b": True})]], size=11,
            align=PP_ALIGN.CENTER, name=f"C cat {k+1}")
for k, (lab, colr) in enumerate([("Completions", "grey_lt"), ("Absorbed", "ink")]):
    box(s, C1 + 0.2 + k * 1.3, 4.0, 0.13, 0.13, fill=colr, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"C key {k+1}")
    textbox(s, C1 + 0.38 + k * 1.3, 3.96, 1.1, 0.2, [[(lab, {})]], size=11, name=f"C keylab {k+1}")
stat_rows(s, C1 + 3.3, 4.02, CW - 3.3, [("9.4%", "Vacancy (peak 10.0%)", "ink"),
                                        ("(2.0%)", "Rent, year over year", "accent"),
                                        ("$1,461", "Effective rent", "ink")], "C stats", num_w=1.0, size=12)

# --- module 4: investor confidence (prior team's module, current data)
header(s, C2, 3.68, CW2, "Investor Confidence", "D header", sup="(4)")
stat_rows(s, C2, 4.02, CW2, [("$10.8B", "Sales volume, trailing 12 months, +3% year over year", "ink"),
                             ("5.66%", "Average cap rate on DFW apartment trades", "ink"),
                             ("$171K", "Per unit, 12% below the 2022 peak of $193K", "accent")], "D", num_w=1.05)
closing(s, [("Buyers and renters are back, but ", {}),
            ("rents and pricing are still below peak, and older Class C trails the metro", {"b": True})])
footer_notes(s, "Notes: (1) Census Vintage 2025; BLS via Dallas Fed. (2) Aug/Sep-26 unemployment; 2024 median household "
                "income (confirm ACS B19013). (3) MMG for 2024; Cushman & Wakefield Q3-26. (4) Matthews Q3-26 (CoStar).",
             "Sources: U.S. Census Bureau; Dallas Fed; MMG; Cushman & Wakefield; Matthews; Northmarq.")
speaker_notes(s, """Presenter: TBD

We kept the structure the earlier team set up for this slide and updated every number.

Population and economy: DFW added about 124,000 people last year, second only to Houston, and since 2020 it has drawn about 270,000 net domestic movers, more than any metro in the country. It has about 4.4 million jobs, growing about 1% a year, with unemployment around 4.4%, slightly above the national 4.2%. Household incomes run about 11% above the U.S., and Fort Worth itself added almost 20,000 people and is now the tenth-largest city in the country.

Apartments: in 2024 the metro built more than 41,000 units and absorbed only about half. This year, through the third quarter, renters absorbed more than 26,700 units against about 21,000 completed. Vacancy is down to 9.4% from a 10% peak, but effective rents are still 2% below a year ago.

Investor confidence: trading has come back, about $10.8 billion over the last twelve months, up 3%, at an average cap rate around 5.7%. But pricing is about $171,000 a unit on average, still 12% below the 2022 peak, and older Class C product like Taylor Commons trades well below that average.

Reference notes:
- Matthews Q3-26 (CoStar): TTM volume ~$10.8B (+3%), average cap 5.66%, $171K/unit vs. $193K Q3-22 peak. Northmarq: 2025 volume +3%, median $172,300/unit; transaction cap rates average ~5.25% (different method).
- Prior team's slide said 'DFW led the nation for absorption in Q2-25' and recommended holding until sales volume recovers; both points still hold.
- Income figures differ by source; confirm DFW and U.S. medians in ACS table B19013 if asked.""")
place(prs, s, pos_m if pos_m is not None else 7)

# ============================================================================ Submarket Overview
s = new_slide(prs, "Submarket Overview: East Fort Worth",
              "East Fort Worth has almost no new supply coming, but its rents are falling 4.6% a year, more than twice "
              "the metro's pace")
add_slide_number(s)
LW = 6.0
# --- supply pipeline (prior team's module; Arlington data replaced)
header(s, C1, g["ROW1"], LW, "Supply Pipeline", "A header", sup="(1)")
cats = ["2024", "Winter 2025", "Q1-26"]
uc = [670, 666, 200]
P = Plot(C1 + 0.3, 2.0, LW - 0.6, 2.3, n=3, vmin=0, vmax=780)
caption(s, C1, 1.68, LW, "East Fort Worth units under construction", "A caption")
chart(s, P, cats, [("Under construction", uc)],
      [bar_ser(0, "Under construction", cats, uc, "grey_lt", dpt={2: "ink"})], "A pipeline chart", gap=90)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="A baseline")
for k, v in enumerate(uc):
    textbox(s, P.center(k) - 0.5, P.v(v) - 0.25, 1.0, 0.22, [[(f"{v:,}", {"b": True})]], size=12,
            align=PP_ALIGN.CENTER, name=f"A label {k+1}")
    textbox(s, P.center(k) - 0.8, P.y + P.h + 0.04, 1.6, 0.22, [[(cats[k], {"b": True})]], size=11,
            align=PP_ALIGN.CENTER, name=f"A cat {k+1}")
numbered_items(s, C1, 4.72, LW, [
    [[("Pipeline down to ~200 units, ", {"b": True}), ("2.4% of an 8,359-unit stock", {})]],
    [[("Occupancy 91.3%, ", {"b": True}), ("up 1.0 point since Q3-25; 252 units absorbed in Q1-26", {})]],
    [[("New supply is not the threat; ", {"b": True}), ("older, discounted product is", {})]],
], [0.32, 0.32, 0.32], name="A item", size=12, num_d=0.24, gap=0.1)

# --- rent dynamics (prior team's module; restated for East Fort Worth)
header(s, C2, g["ROW1"], CW2, "Rent Dynamics", "B header", sup="(2)")
rows = [
    ["Average rent", "$1,088", "$1,050"],
    ["Rent, year over year", "(1.4%)", "(4.6%)"],
    ["DFW, year over year", "", "(2.0%)"],
]
yt = grid_table(s, C2, 1.72, [2.6, 1.3, 2.1], ["", "Q3-25", "Q1-26 / latest"], rows, "B table", size=12, row_h=0.31,
                hair_rules=True, accent_cols={(0, 2), (1, 2)})
numbered_items(s, C2, yt + 0.12, CW2, [
    [[("Asking rents: ", {"b": True}), ("1 bed $983, 2 bed $1,277, down 3.6% a year", {})]],
    [[("Renters: ", {"b": True}), ("about half of homes in ZIP 76112 are rented", {})]],
    [[("Taylor Commons: ", {"b": True}), ("$970 market rent, 74.9% occupied (Nov-26)", {"c": "accent"})]],
], [0.32, 0.32, 0.32], name="B item", size=12, num_d=0.24, gap=0.08)

# --- key employers (prior team's logos)
header(s, C2, 4.62, CW2, "Key Employers Within ~20 Minutes", "C header", sup="(3)")
add_image(s, os.path.join(IMG, "logo_american.png"), C2 + 0.05, 5.08, 1.7, 0.63, name="C logo American")
add_image(s, os.path.join(IMG, "logo_uta.png"), C2 + 2.05, 5.12, 1.75, 0.56, name="C logo UTA")
add_image(s, os.path.join(IMG, "logo_gm.png"), C2 + 4.2, 5.03, 0.95, 0.72, name="C logo GM")
textbox(s, C2, 5.8, CW2, 0.24, [[("Residents' top employers at IC: ", {"i": True, "c": "grey_txt"}),
                                 ("Walmart, Amazon, FedEx", {"i": True, "b": True, "c": "grey_txt"})]], size=11,
        name="C tenants")
closing(s, [("Supply is not the problem here; ", {}),
            ("falling submarket rents cap what an older asset can charge (comps next)", {"b": True})])
footer_notes(s, "Notes: (1) Matthews 2024; GREA winter 2025; Transwestern Q1-26. (2) Transwestern Q3-25 and Q1-26; "
                "Cushman & Wakefield; Apartments.com Jan-26; ACS 2019-23. (3) Logos from the prior team's slide; IC "
                "presentation.",
             "Sources: Transwestern; GREA; Matthews; Cushman & Wakefield; Apartments.com; U.S. Census; TexSun.")
speaker_notes(s, """Presenter: TBD

We kept the earlier team's three blocks for the submarket, supply pipeline, rent dynamics and key employers, and replaced the Arlington data that had been carried over from another deal with east Fort Worth.

Supply first. East Fort Worth had about 670 units under construction in 2024, the most in twenty years, and by early 2026 that is down to about 200, only 2.4% of an 8,400-unit stock. Occupancy is about 91% and improving. New supply is not what is hurting Taylor Commons.

Rents are. Average rent in the submarket fell from about $1,088 to $1,050, and the year-over-year decline widened to 4.6%, more than twice the metro's 2%. Asking rents run about $983 for a one-bedroom and $1,277 for a two-bedroom. About half the homes around the property are rented, so the renter base is there; the competition is filling units by cutting price. Taylor Commons markets at $970 and is 74.9% occupied.

Employment is a strength: American Airlines' headquarters, UT Arlington and GM's Arlington assembly plant are all within about twenty minutes, and at IC the residents' top employers were Walmart, Amazon and FedEx.

Reference notes:
- Transwestern Q3-25: 38 buildings, 8,068 units, 90.3%, $1,088, -1.4%. Q1-26: 39 buildings, 8,359 units, 91.3%, 252 absorbed, $1,050, -4.6%, 200 under construction (extract partly garbled; confirm).
- GREA winter 2025: 666 under construction, none delivering in the next four quarters. Matthews 2024: ~670 under way, the highest in 20 years.
- Prior team's slide: the chart and pipeline text (7,300 units delivered, 1,700 under construction, 'Mid-Cities') described Arlington, not east Fort Worth, so they were replaced.""")
place(prs, s, (pos_m + 1) if pos_m is not None else (old_sub if old_sub is not None else 8))

prs.save(DST)
print("saved", DST)
