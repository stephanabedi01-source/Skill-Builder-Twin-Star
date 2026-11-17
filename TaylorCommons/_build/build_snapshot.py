"""Slide: Then vs. Now (IC underwriting vs. today). Replaces the old Asset Overview content.

Usage: python3 build_snapshot.py <in.pptx> <out.pptx> [position]

IC column (Apr-23):
  Occupancy 86%                          IC presentation, transaction details
  Rent: $975 avg in place at close       2025 Q2 sponsor asset report (unit table, 'at close')
        2026 underwritten ~$1,086        975 x 1.04 x 1.04 x 1.03 (IC model Growth Assumptions 2024-2026)
  Year 3 NOI $2.21M                      IC model Transaction Summary DSCR table, Year 3 NOI 2,206,017
  Year 2 NOI $2.00M                      same table, Year 2 1,998,007
  Senior $14.1M at 3.57% fixed, 36 mo.   IC presentation; Capital Structure tab (14,144,000 at close)
  Pref $5.6M, 8% current / 12.5% IRR     Control Sheet 5,575,000; IC memo terms
  Property taxes 2025 UW $640K           IC model Growth Assumptions, Real Estate Taxes 2025 639,631
  Net return to SMU 17.6% IRR, 1.9x      IC memo
Today column:
  Occupancy 74.9%, leased 78.5%          AM model v1_1 'Rent Roll' tab (box score 11/15/2026)
  Avg market rent $970                   same, total row
  T12 NOI $1.43M (Mar-25 to Feb-26)      Income_Statement_T12
  Year 2 actual NOI $0.37M               Data drop Apr-24..Mar-25 (31,372 ... 104,416) = 370,249
  Senior $16.5M at ~9.0% I/O due Mar-27  Deed of Trust; annual report
  Pref $5.03M, 10% accruing              Capital Structure tab; annual report
  Property taxes T12 $28K (HFC)          Income_Statement_T12
  Rent rules 50% <80% AMI, 90% <140%     Regulatory Agreement
  Renovated layouts 12 units             box score 'Renovated' block (A2R 3, A3R 1, A4R 2, B1R 2, B2R 4)
  SMU recovery $0                        Fund annual report 2025-26
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 2

UW_RENT_26 = 975 * 1.04 * 1.04 * 1.03
assert round(UW_RENT_26) == 1086
UW_NOI = [1.998, 2.206]
ACT_NOI = [0.370, 1.429]
NOI_TAXED = 0.937

prs = Presentation(SRC)
g = THEME["grid"]
s = new_content_slide(
    prs, "Then vs. Now: IC Underwriting vs. Today",
    "Three years in, occupancy is 75% vs. 86% at IC and T12 NOI is $1.43M vs. $2.21M underwritten, even with "
    "the tax break")
add_slide_number(s)

# ------------------------------------------------------------------ left: the comparison table
L0, LW = g["L0"], g["LP1"] - g["L0"]
section_header(s, L0, g["ROW1"], LW, "What Was Underwritten vs. Where It Stands", "A header")
rows = [
    ["Occupancy", "86%", [("74.9%; 78.5% leased", {}), ("(1)", {"sup": True})]],
    ["Average rent", "$1,086 by 2026 (+11%)", "$970 market (Nov-26)"],
    ["Year 3 NOI", "$2.21M", [("$1.43M T12, with tax break", {}), ("(2)", {"sup": True})]],
    ["Property taxes", "$640K (2025)", "$28K T12; exemption ends YE2026"],
    ["Senior loan", "$14.1M agency, 3.57% fixed", "$16.5M bridge, ~9.0% I/O, due Mar-27"],
    ["Preferred equity", "$5.6M, 8% current pay", "$5.0M+, accruing at 10%"],
    ["Rent and income rules", "None", "50% <80% AMI, 90% <140% AMI"],
    ["Renovated units", "$1.8M upgrade budget", "12 units in renovated layouts"],
    ["Operator", "TexSun with Asset Living", "Cottonwood running operations"],
    [[("Net return to SMU", {"b": True})], [("17.6% IRR, 1.9x", {"b": True})],
     [("$0 recovery (fund mark)", {"b": True})]],
]
grid_table(s, L0, 1.74, [1.75, 2.35, 3.25], ["", "IC (Apr-23)", "Today (Nov-26)"], rows, "A table",
           size=10.5, row_h=0.375, hair_rules=True, left_cols=(0, 1, 2), rule_before=(9,),
           accent_cols={(0, 2), (2, 2), (4, 2), (9, 2)})

# ------------------------------------------------------------------ right: NOI vs. plan
R0, RW = g["RP0"], g["R1"] - g["RP0"]
section_header(s, R0, g["ROW1"], RW, "Underwritten vs. Actual NOI ($M)", "B header", sup="(3)")
for k, (lab, colr) in enumerate([("IC underwriting", "grey_lt"), ("Actual", "accent_fill")]):
    lx = R0 + k * 1.75
    box(s, lx, 1.78, 0.14, 0.14, fill=colr, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"B legend key {k+1}")
    textbox(s, lx + 0.2, 1.74, 1.5, 0.22, [[(lab, {})]], size=10, name=f"B legend {k+1}")
cats = ["Year 2", "Year 3"]
P = Plot(R0 + 0.15, 2.3, RW - 0.3, 2.3, n=2, vmin=0, vmax=2.5)
chart(s, P, cats, [("IC underwriting", UW_NOI), ("Actual", ACT_NOI)],
      [bar_ser(0, "IC underwriting", cats, UW_NOI, "grey_lt"), bar_ser(1, "Actual", cats, ACT_NOI, "accent_fill")],
      "B noi chart", grouping="clustered", gap=70, overlap=-8)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="B baseline")
bw = P.slot / (2 + 0.08 + 0.70)
for k in range(2):
    for j, (v, colr) in enumerate([(UW_NOI[k], "ink"), (ACT_NOI[k], "accent")]):
        bx = P.center(k) + (j - 0.5) * bw * 1.08
        textbox(s, bx - 0.45, P.v(v) - 0.23, 0.9, 0.2, [[(f"${v:.2f}M", {"b": True, "c": colr})]], size=10.5,
                align=PP_ALIGN.CENTER, name=f"B label {k+1}-{j+1}")
sub = ["Apr-24 to Mar-25", "T12 to Feb-26"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, P.center(k) - 1.0, P.y + P.h + 0.06, 2.0, 0.42,
            [[(cat, {"b": True})], [(sb, {"c": "grey_txt", "i": True})]], size=10, align=PP_ALIGN.CENTER,
            name=f"B category {k+1}")
callout(s, R0, 5.2, RW, 0.62,
        [("With taxes paid, Year 3 NOI is ", {}), (f"${NOI_TAXED:.2f}M", {"b": True}),
         (f", {NOI_TAXED / UW_NOI[1]:.0%} of plan", {})], "B callout", size=10.5)

# ------------------------------------------------------------------ closing, footer, notes
closing_line(s, [("The plan missed on operations before the refinancing; ", {}),
                 ("the March 2025 HFC trade then added cost the property could not carry", {"b": True})])
footer_notes(s,
             "Notes: (1) Box score 11/15/2026. (2) T12 Mar-25 to Feb-26 with property taxes of $28K under the HFC. "
             "(3) IC model project years run from an assumed Apr-23 close (actual close 6/15/2023); Year 2 actual "
             "is Apr-24 to Mar-25 from sponsor statements. 2026 rent is $975 at close grown at the IC model's "
             "4%, 4%, 3%.",
             "Sources: IC presentation and memo (Apr/May-23); IC model (3/18/23); TexSun T12 and box score; "
             "Regulatory Agreement (3/24/25); Fund annual report 2025-26.")

speaker_notes(s, """Presenter: TBD

Here is the deal we approved next to the deal we have today.

At IC, Taylor Commons was 86% occupied and the plan was to push rents about 11% by 2026 and reach $2.2 million of NOI in year three. Today it is 75% occupied, rents are about $970, roughly where they were at closing, and trailing NOI is $1.43 million. And that $1.43 million only looks that good because the property is paying almost no property tax under the HFC. With taxes paid it is about $940 thousand, 42% of the plan.

The capital stack changed too. The 3.57% agency loan is gone, replaced by a $16.5 million bridge loan at about 9%, due in March. The pref is accruing at 10%. The property now carries rent and income rules from the HFC. Only 12 units sit in renovated floor plans, against a $1.8 million upgrade budget, and the pref investor, Cottonwood, is now running operations instead of TexSun.

The bottom row is the one that matters for us: underwritten at a 17.6% net IRR, marked today at zero.

Year 2 is worth noting on the right: before any refinancing, the property made $370 thousand against $2 million of plan. Operations missed first. Then the refinancing made it worse. The next slides walk through both.

Reference notes:
- IC presentation: 86% occupancy, $25.9M price ($115,625/unit), 17.6% net IRR / 1.9x; assumed agency loan $14.2M at 3.57% fixed, 36 months.
- Year 2 actual NOI includes a Dec-24 property tax true-up of $127.6K (NOI that month was -$85.8K).
- Box score 11/15/2026: 167 of 223 rentable units occupied (74.9%), 78.5% leased, 56 vacant of which 49 not ready.
- Fixes the old Asset Overview line '24 affordable units': the Regulatory Agreement requires at least 112 units (50%) for households under 80% AMI and 202 (90%) under 140% AMI.""")

move_slide(prs, s, POS)
prs.save(DST)
print("saved", DST, "at position", POS + 1)
