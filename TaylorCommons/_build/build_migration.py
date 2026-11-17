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

# ------------------------------------------------------------------ left: components of change
header(s, L0, g["ROW1"], LW, "Tarrant County Population Change, Jul-24 to Jul-25 (000s)", "A header", sup="(1)")
caption(s, L0, g["ROW1"] + 0.36, LW, "Components of change, Census Vintage 2025 estimates", "A caption")
cats = ["Births less deaths", "International", "Domestic", "Net change"]
base = [0, NAT, NAT + INTL + DOM, 0]
pos = [NAT, INTL, 0, TOT]
neg = [0, 0, -DOM, 0]   # drawn as a positive span sitting on the base
P = Plot(0.85, 2.2, 6.7, 2.95, n=4, vmin=0, vmax=28)
chart(s, P, cats, [("Base", base), ("Up", pos), ("Down", neg)],
      [nofill_ser(0, "Base", cats, base),
       bar_ser(1, "Up", cats, pos, "grey_lt", dpt={3: "ink"}),
       bar_ser(2, "Down", cats, neg, "accent_fill")],
      "A components chart", grouping="stacked", gap=70)
bw = P.bar_width(70)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="A baseline")
for k, lvl in [(0, NAT), (1, NAT + INTL), (2, NAT + INTL + DOM)]:
    line(s, P.center(k) + bw / 2, P.v(lvl), P.center(k + 1) - bw / 2, P.v(lvl), "grey_mid", 0.75,
         dash=MSO_LINE_DASH_STYLE.DASH, name=f"A connector {k+1}")
for k, top, txt, colr in [(0, NAT, "+12.4", "ink"), (1, NAT + INTL, "+12.8", "ink"),
                          (2, NAT + INTL, "(7.0)", "accent"), (3, TOT, "+18.4", "ink")]:
    textbox(s, P.center(k) - 0.6, P.v(top) - 0.27, 1.2, 0.22, [[(txt, {"b": True, "c": colr})]], size=12,
            align=PP_ALIGN.CENTER, name=f"A label {k+1}")
sub = ["Natural increase", "Net arrivals from abroad", "Net moves within the U.S.", "Total growth"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, P.center(k) - 0.85, P.y + P.h + 0.06, 1.7, 0.46,
            [[(cat, {"b": True})], [(sb, {"c": "grey_txt", "i": True})]], size=11, align=PP_ALIGN.CENTER,
            name=f"A category {k+1}")

# ------------------------------------------------------------------ right: city and ZIP
header(s, R0, g["ROW1"], RW, "City of Fort Worth", "B header", sup="(2)")
rows_c = [
    ["Population, Jul-24", "1,008,106"],
    ["Growth, 2023 to 2024", "+23,442"],
    ["Growth, 2024 to 2025", "+19,512"],
    ["Growth since 2020", "+9.7%"],
]
y = 1.72
for k, (a, b) in enumerate(rows_c):
    textbox(s, R0, y, 3.0, 0.26, [[(a, {})]], size=12, name=f"B k{k+1}", anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, R0 + 3.0, y, 1.63, 0.26, [[(b, {"b": True})]], size=12, align=PP_ALIGN.RIGHT, name=f"B v{k+1}",
            anchor=MSO_ANCHOR.MIDDLE)
    y += 0.3
    line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"B rule {k+1}")
header(s, R0, y + 0.16, RW, "ZIP 76112 Around the Property", "C header", sup="(3)")
y += 0.6
rows_z = [
    ("Median household income", "$56,179"),
    ("Renter-occupied homes", "~51%"),
    ("Median gross rent", "$1,215"),
    ("Taylor Commons market rent", "$970"),
]
for k, (a, b) in enumerate(rows_z):
    hl = k == 3
    textbox(s, R0, y, 3.0, 0.26, [[(a, {"b": hl, "c": "accent" if hl else "ink"})]], size=12, name=f"C k{k+1}",
            anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, R0 + 3.0, y, 1.63, 0.26, [[(b, {"b": True, "c": "accent" if hl else "ink"})]], size=12,
            align=PP_ALIGN.RIGHT, name=f"C v{k+1}", anchor=MSO_ANCHOR.MIDDLE)
    y += 0.3
    if k < len(rows_z) - 1:
        line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"C rule {k+1}")
callout(s, R0, y + 0.14, RW, 0.62,
        [("Half the area rents, and rents run ", {}), ("below what local incomes support", {"b": True})],
        "C callout", size=12)

closing(s, [("Renter demand is there, but not from new domestic arrivals: ", {}),
            ("Taylor Commons must win local renters on price and condition", {"b": True})])
footer_notes(s, "Notes: (1) Census Vintage 2025 components as published by USAFacts; international and domestic "
                "are rounded. (2) Census city estimates, May 2025 and 2026 releases. (3) ACS 2019-2023 5-year; "
                "gross rent includes utilities.",
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
move_slide(prs, s, idx if idx is not None else 8)
prs.save(DST)
print("saved", DST, "at position", (idx if idx is not None else 8) + 1)
