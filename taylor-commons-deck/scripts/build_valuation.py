"""Slide: Valuation and Recovery.

Usage: python3 build_valuation.py <in.pptx> <out.pptx> [position]

Value ($M):
  Income approach, cap 6.0% / 6.5% / 7.0%:
    No relief NOI 937.4   -> 15.62 / 14.42 / 13.39
    Exemption kept 1,429.3 -> 23.82 / 21.99 / 20.42
  Fund annual report (as of 3/31/2026): ~$19.2M at a 6.5% cap on recast NOI ($85,714/unit)
  Sales comps (annual report): $91K to $137K per unit x 224 = $20.38M to $30.69M (stabilized trades)
Capital stack: senior $16.5M (W Financial bridge); pref $5.03M (Capital Structure tab / annual report; accrues 10%
  when unpaid, so the true balance is likely higher); common $9.875M incl. SMU $350K.
Selling costs 3.5% (model Primary Assumptions 'Est. Closing Costs').
Breakeven NOI for $1 to common: (16.5 + 5.03) / (1 - 3.5%) x 6.5% = $1.45M (55% above the 2027 base NOI
  of $0.94M; just above the exempt T12 NOI of $1.43M).
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 14

LOAN, PREF, COMMON, SMU = 16.5, 5.03, 9.875, 0.35
COST = 0.035
NOI_B, NOI_K = 0.9374, 1.4293
val = lambda noi, cap: noi / cap  # noqa: E731
VB = [val(NOI_B, c) for c in (0.07, 0.065, 0.06)]
VK = [val(NOI_K, c) for c in (0.07, 0.065, 0.06)]
COMP = (0.091 * 224, 0.137 * 224)
BE_NOI = (LOAN + PREF) / (1 - COST) * 0.065
assert round(VB[0], 1) == 13.4 and round(VB[2], 1) == 15.6 and round(BE_NOI, 2) == 1.45


def waterfall(value):
    net = value * (1 - COST)
    sr = min(net, LOAN); rem = net - sr
    pf = min(rem, PREF); rem -= pf
    cm = min(rem, COMMON)
    return net, sr, pf, cm


WB, WK = waterfall(VB[1]), waterfall(VK[1])

prs = Presentation(SRC)
g = THEME["grid"]
s = new_slide(
    prs, "Valuation and Recovery",
    "Without relief the asset is worth $13.4M to $15.6M, below the $16.5M loan, so SMU's $350K recovers $0")
add_slide_number(s)

# ------------------------------------------------------------------ left: value range vs. the stack
L0, LW = g["L0"], g["LP1"] - g["L0"]
header(s, L0, g["ROW1"], LW, "Value Range vs. the Capital Stack ($M)", "A header", sup="(1)")
cats = ["Income: no relief", "Income: exemption kept", "Fund mark (Mar-26)", "Sales comps"]
lo = [VB[0], VK[0], None, COMP[0]]
span = [VB[2] - VB[0], VK[2] - VK[0], None, COMP[1] - COMP[0]]
P = Plot(2.55, 1.85, 5.05, 3.45, n=4, vmin=10, vmax=32, horizontal=True)
chart(s, P, cats, [("Low", lo), ("Range", span)],
      [nofill_ser(0, "Low", cats, lo),
       bar_ser(1, "Range", cats, span, "grey_lt", dpt={0: "accent_fill"})],
      "A football field", grouping="stacked", gap=70)
bh = P.bar_width(70)
# fund mark as a diamond
ym = P.center(2)
box(s, P.v(19.2) - 0.09, ym - 0.09, 0.18, 0.18, fill="ink", line=None, shape=MSO_SHAPE.DIAMOND, name="A fund mark")
textbox(s, P.v(19.2) - 0.75, ym + 0.12, 1.5, 0.2, [[("$19.2M", {"b": True})]], size=12, align=PP_ALIGN.CENTER,
        name="A fund mark label")
# range labels
textbox(s, P.v(VB[2]) - 1.75, P.center(0) - bh / 2 - 0.27, 1.7, 0.22,
        [[(f"${VB[0]:.1f} to ${VB[2]:.1f}", {"b": True, "c": "accent"})]], size=12, align=PP_ALIGN.RIGHT,
        name="A range 1")
rng = [(1, VK[0], VK[2], "ink"), (3, COMP[0], COMP[1], "ink")]
for k, a, b, colr in rng:
    yy = P.center(k) - 0.1
    textbox(s, P.v(a) - 0.72, yy, 0.66, 0.2, [[(f"${a:.1f}", {"b": True, "c": colr})]], size=12,
            align=PP_ALIGN.RIGHT, name=f"A low {k+1}")
    textbox(s, P.v(b) + 0.06, yy, 0.66, 0.2, [[(f"${b:.1f}", {"b": True, "c": colr})]], size=12,
            name=f"A high {k+1}")
# row labels (left of plot)
sub = ["6.0% to 7.0% cap on $0.94M", "6.0% to 7.0% cap on $1.43M", "6.5% cap, recast NOI", "$91K to $137K per unit"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, L0, P.center(k) - 0.21, 1.95, 0.42, [[(cat, {"b": True, "c": "accent" if k == 0 else "ink"})],
                                                    [(sb, {"c": "grey_txt", "i": True})]],
            size=11, align=PP_ALIGN.RIGHT, name=f"A row {k+1}")
# claim lines: loan, loan + pref
for v, lab, colr, nm, al in [(LOAN, "Loan $16.5M", "accent", "loan", PP_ALIGN.RIGHT),
                             (LOAN + PREF, "Loan + pref $21.5M", "ink", "pref", PP_ALIGN.LEFT)]:
    x = P.v(v)
    line(s, x, P.y - 0.08, x, P.y + P.h, colr, 1.25, dash=MSO_LINE_DASH_STYLE.DASH, name=f"A {nm} line")
    lx = x - 1.45 if al == PP_ALIGN.RIGHT else x + 0.05
    textbox(s, lx, P.y + P.h + 0.04, 1.4, 0.2, [[(lab, {"b": True, "c": colr})]], size=12,
            align=al, name=f"A {nm} label")
caption(s, L0, P.y + P.h + 0.36, LW,
        "Common equity is in the money only to the right of the loan + pref line", "A caption")

# ------------------------------------------------------------------ right: who gets paid
R0, RW = g["RP0"], g["R1"] - g["RP0"]
header(s, R0, g["ROW1"], RW, "Sale at a 6.5% Cap: Who Gets Paid ($M)", "B header", sup="(2)")
pct = lambda a, b: f"{a / b:.0%}"  # noqa: E731
rows = [
    ["Value", f"${VB[1]:.1f}", f"${VK[1]:.1f}"],
    ["Net of 3.5% selling costs", f"${WB[0]:.1f}", f"${WK[0]:.1f}"],
    ["Senior loan ($16.5)", f"${WB[1]:.1f} ({pct(WB[1], LOAN)})", f"${WK[1]:.1f} ({pct(WK[1], LOAN)})"],
    [[("Pref ($5.0)", {}), ("(3)", {"sup": True})], f"${WB[2]:.1f} ({pct(WB[2], PREF)})",
     f"${WK[2]:.1f} ({pct(WK[2], PREF)})"],
    ["Common ($9.9)", f"${WB[3]:.1f} ({pct(WB[3], COMMON)})", f"${WK[3]:.1f} ({pct(WK[3], COMMON)})"],
    [[("SMU ($0.35)", {"b": True})], [("$0", {"b": True})], [("$0", {"b": True})]],
]
yb = grid_table(s, R0, 1.74, [1.95, 1.3, 1.38], ["", "No relief", "Exemption kept"], rows, "B table",
                size=12, row_h=0.31, hair_rules=True, rule_before=(2, 5),
                accent_cols={(2, 1), (3, 1), (4, 1), (4, 2), (5, 1), (5, 2)})
callout(s, R0, yb + 0.22, RW, 0.66,
        [("The common breaks even near ", {}), (f"${BE_NOI:.2f}M", {"b": True}), (" of NOI", {}),
         ("(4)", {"sup": True}), (", 55% above the 2027 base case of $0.94M", {})], "B callout", size=12)

# ------------------------------------------------------------------ closing, footer, notes
closing(s, [("SMU's $350K recovers nothing on income value; ", {}),
                 ("the common needs a sale above ~$100K per unit, stabilized comp pricing", {"b": True})])
footer_notes(s,
             "Notes: (1) Comps are stabilized trades per the fund annual report. (2) Net of 3.5% selling costs. (3) $5.03M; accrues at 10%, so likely higher. (4) ($16.5M + $5.0M) / 96.5% x 6.5%.",
             "Sources: TexSun T12; Fund annual report 2025-26; AM model; team analysis.")

speaker_notes(s, """Presenter: TBD

So what is the building worth, and who gets paid if it sells?

On the left, four ways to value it, against two lines: the $16.5 million loan in red and the loan plus the preferred equity at about $21.5 million. Our common equity only has value to the right of that second line.

In red, our base case. With the exemption gone, NOI is about $940 thousand. At a 6 to 7 percent cap rate, that is $13.4 to $15.6 million, below the loan. If the exemption survives, value is $20.4 to $23.8 million, which clears the loan but mostly not the pref. The fund marked it at $19.2 million in March. Recent sales of stabilized properties ran $91 to $137 thousand a unit, $20 to $31 million. The common breaks even at about $100 thousand a unit, so it takes stabilized pricing to reach us, and this building is 74% occupied, not stabilized.

On the right, a sale at a 6.5 cap. With no relief the lender takes about 84 cents on the dollar and the pref and common get nothing. Even with the exemption, the pref takes a small haircut and the common gets nothing. The common only sees a dollar once NOI reaches about $1.45 million. That is 55% above our 2027 base case, and slightly above what the property made this year even with the tax exemption.

So the fund's $350 thousand is at zero in every case we can underwrite.

Reference notes:
- Annual report: 6.5% cap on recast NOI -> ~$19.2M ($85,714/unit); submarket average cap 5.9% for stabilized assets; comps $91K to $137K/unit; recovery considered zero.
- Pref: $5.03M as of the Mar-25 capital structure; 10% accruing when unpaid per the annual report (IC memo terms were 8% current / 12.5% IRR lookback with Electra: reconcile with TexSun).
- Selling costs 3.5% from the AM model's Primary Assumptions.
- Breakeven NOI: (16.5 + 5.03) / 0.965 x 0.065 = $1.45M; any pref accrual raises it.
- Sales comps are from the annual report; the team's updated sales comp outreach should replace or confirm them.""")

place(prs, s, POS)
prs.save(DST)
print("saved", DST)
