"""Slide: The March 2025 Refinancing Trade (HFC block, slide 2 of 4).

Usage: python3 build_refi_trade.py <in.pptx> <out.pptx> [position]

Numbers (all $000s unless noted), T12 before = Mar-24..Feb-25, T12 after = Mar-25..Feb-26:
  Interest before   499.5   Operating Stmt. Data Drop, 'Debt Service - 1st Lien Interest' (v1 AM model)
  Interest after  1,539.6   Income_Statement_T12 (Mar-25..Feb-26), 'Debt Service - 1st Lien Interest'
  Taxes before      519.9   Data Drop, 'Taxes & Dues - Property Taxes' Mar-24..Feb-25
  Taxes after        28.0   Income_Statement_T12, 'Taxes & Dues - Property Taxes'
  HFC fees           64.1   T12 'Housing Authority - Annual Exp' 51.0 + 'Admin fees' 13.1 (non-operating)
  NOI after       1,429.3   Income_Statement_T12 'Net Operating Income'
  Old balance    ~13,690    Jan-25 interest 42,084 / (3.57% x 31/360)  (estimate)
  New loan        16,500    Deed of Trust 3/24/2025 (W Financial REIT); annual report; Capital Structure tab
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 11

INT_B, INT_A = 499.5, 1539.6
TAX_B, TAX_A = 519.9, 28.0
FEES = 64.1
NOI_A = 1429.3
TAX_SAVED = round(TAX_B - TAX_A, 1)                 # 491.9
ADD_INT = round(INT_A - INT_B, 1)                   # 1,040.1
NET = round(TAX_SAVED - FEES - ADD_INT, 1)          # -612.3
NOI_AGENCY = NOI_A - TAX_SAVED                      # 937.4: same operations, taxes at pre-HFC level
COV_HFC = NOI_A / INT_A                             # 0.93x
COV_AGENCY = NOI_AGENCY / INT_B                     # 1.88x
COV_MKT = NOI_AGENCY / (13690 * 0.0525)             # 1.30x at a ~5.25% agency refinancing
assert abs(NET - (-612.3)) < 0.5, NET

prs = Presentation(SRC)
g = THEME["grid"]
s = new_content_slide(
    prs, "The March 2025 Refinancing Trade",
    "The HFC refinancing added $1.04M of annual interest to save $428K of taxes net of fees, a $612K yearly loss")
add_slide_number(s)

# ------------------------------------------------------------------ left panel: waterfall
LW = g["LP1"] - g["L0"]
section_header(s, g["L0"], g["ROW1"], LW, "Annual Cash Flow Effect of the Trade ($000s)", "A header", sup="(1)")
caption(s, g["L0"], g["ROW1"] + 0.36, LW,
        "T12 to Feb-26 under the HFC and bridge loan vs. T12 to Feb-25 under the agency loan", "A caption")

cats = ["Property taxes avoided", "HFC fees", "Added interest", "Net annual cost"]
base = [0, TAX_SAVED - FEES, 0, 0]
pos = [TAX_SAVED, FEES, TAX_SAVED - FEES, 0]
neg = [0, 0, -(ADD_INT - (TAX_SAVED - FEES)), NET]
P = Plot(0.75, 2.14, 6.9, 3.05, n=4, vmin=-700, vmax=560)
chart(s, P, cats, [("Base", base), ("Increase", pos), ("Decrease", neg)],
      [nofill_ser(0, "Base", cats, base),
       bar_ser(1, "Increase", cats, pos, "ink", dpt={1: "grey_mid", 2: "accent_fill"}),
       bar_ser(2, "Decrease", cats, neg, "accent_fill")],
      "A waterfall chart", grouping="stacked", gap=75)
bw = P.bar_width(75)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="A zero line")
# connectors at running totals
for k, lvl in [(0, TAX_SAVED), (1, TAX_SAVED - FEES), (2, NET)]:
    line(s, P.center(k) + bw / 2, P.v(lvl), P.center(k + 1) - bw / 2, P.v(lvl), "grey_mid", 0.75,
         dash=MSO_LINE_DASH_STYLE.DASH, name=f"A connector {k+1}")
# value labels
labels = [(0, TAX_SAVED, "+$492", "above"), (1, TAX_SAVED, "($64)", "above"),
          (2, NET, "($1,040)", "below"), (3, NET, "($612)", "below")]
for k, v, txt, where in labels:
    y = P.v(v) - 0.25 if where == "above" else P.v(v) + 0.04
    col_ = "accent" if k >= 2 else "ink"
    textbox(s, P.center(k) - 0.6, y, 1.2, 0.21, [[(txt, {"b": True, "c": col_})]], size=11, align=PP_ALIGN.CENTER,
            name=f"A label {k+1}")
# category labels
sub = ["T12 $520 to $28", "Pecos HFC, booked below NOI", "T12 $500 to $1,540", "While the exemption holds"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, P.center(k) - 0.85, 5.45, 1.7, 0.44, [[(cat, {"b": True})], [(sb, {"c": "grey_txt", "i": True})]],
            size=10, align=PP_ALIGN.CENTER, name=f"A category {k+1}")
# pill: the only positive bar is the one HB 21 removes
px, py = P.center(0) - 0.95, P.v(-330)
pill(s, px, py, 1.90, 0.34, [("Ends at YE2026", {"b": True, "c": "accent"}), (" under HB 21", {})],
     "A pill HB21", size=10)
line(s, P.center(0), P.v(0) + 0.02, P.center(0), py, "accent", 0.75, dash=MSO_LINE_DASH_STYLE.ROUND_DOT,
     name="A pill leader")

# ------------------------------------------------------------------ right panel: old loan vs. bridge
RW = g["R1"] - g["RP0"]
section_header(s, g["RP0"], g["ROW1"], RW, "Agency Loan vs. HFC Bridge", "B header")
rows = [
    ["Lender", "Agency (assumed)", "W Financial REIT"],
    ["Loan balance", [("$13.7M", {}), ("(2)", {"sup": True})], "$16.5M"],
    ["Rate", "3.57% fixed", [("9.50% I/O", {}), ("(3)", {"sup": True})]],
    ["Term", "Fixed to ~Jun-26", "Due Mar-27"],
    ["T12 interest", "$500K", "$1,540K"],
    ["T12 property taxes", "$520K", "$28K"],
    ["Rent restrictions", "None", "50% / 90% of units"],
    [[("Interest coverage", {}), ("(4)", {"sup": True})], "1.9x", "0.93x"],
]
yb = grid_table(s, g["RP0"], 1.74, [1.62, 1.44, 1.52], ["", "Agency loan", "HFC bridge"], rows, "B table",
                size=10.5, row_h=0.31, bold_rows=(7,), rule_before=(7,), hair_rules=True,
                accent_cols={(2, 2), (4, 2), (6, 2), (7, 2)})
callout(s, g["RP0"], yb + 0.24, RW, 0.66,
        [("Even after the 3.57% lock expired, a ~5.25% agency loan on the same NOI covers interest ", {}),
         ("1.3x", {"b": True}), ("(5)", {"sup": True})],
        "B callout", size=10.5)

# ------------------------------------------------------------------ closing, notes, sources
closing_line(s, [("The exemption never paid for the debt it required, and ", {}),
                 ("HB 21 ends the exemption at YE2026 while the 9% debt remains", {"b": True})])
footer_notes(s,
             "Notes: (1) T12 to Feb-26 vs. T12 to Feb-25. (2) Estimate from Jan-25 interest. (3) 9.00% from Mar-26. "
             "(4) Same NOI with pre-HFC taxes. (5) $937K / ($13.7M x 5.25%).",
             "Sources: TexSun T12s; Deed of Trust (3/24/25); IC deck; Fund annual report 2025-26.")

speaker_notes(s, """Presenter: TBD

In March 2025 the sponsor did two things on the same day. They moved the land into a housing finance corporation out of Pecos, Texas, which takes the property off the tax rolls, and they paid off our 3.57% agency loan with a $16.5 million bridge loan at 9.5%, interest only.

Here is what that did to a year of cash flow. Taxes went from $520 thousand to $28 thousand, so about $492 thousand saved. The HFC charges about $64 thousand a year in fees. But interest went from $500 thousand to more than $1.5 million. Net, the property gives up about $612 thousand a year, and that is while the exemption is still working.

On the right, the same story by loan. Coverage on the old structure, using this year's NOI with taxes added back, is about 1.9 times. On the new one it is 0.93. The 3.57% rate was locked only through mid-2026, but even a normal agency refinancing in the low fives would still cover interest about 1.3 times.

And the one bar that helped, the tax savings, goes away at the end of this year under HB 21. That is the next slide.

Reference notes:
- Deed of Trust recorded 3/25/2025, dated 3/24/2025: 5600 Cotswold SPE LLC and Pecos HFC to W Financial REIT, Ltd.
- Loan $16.5M at 9.5% I/O; extended Mar-26 for 12 months at about 9.0% (annual report, TexSun Q1-26).
- New loan was about $2.8M larger than the agency balance it repaid (~$13.7M); use of proceeds not yet explained by TexSun. Mar-25 T12 shows $803K of partner loan interest.
- The T12 to Feb-26 also accrued $137K of 2026 property taxes (Dec-25 to Feb-26); confirm with TexSun whether that is a reserve against losing the exemption.
- IC deck (Apr-23): assumed agency loan $14.2M, 3.57% fixed, 36 months.""")

move_slide(prs, s, POS)
prs.save(DST)
print("saved", DST, "at position", POS + 1)
