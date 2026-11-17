"""Slide: The 2027 Cliff and Refinancing Gap (HFC block, slide 4 of 4).

Usage: python3 build_cliff.py <in.pptx> <out.pptx> [position]

NOI ($000s): base (no relief) 937.4 = T12 Feb-26 NOI 1,429.3 less tax increase 491.9 (see build_refi_trade.py);
  exemption kept 1,429.3.
Sizing: rate 6.75% (TexSun Q1-26: refi target mid-6s to 7%), 30-yr amortization -> constant 7.783%;
  DSCR 1.25x; debt yield 8.0%; LTV 65% of value at a 6.5% cap (annual report uses 6.5%).
  No relief:  DSCR 9,635  DY 11,718  LTV 9,374  -> binding 9,374, gap 7,126
  Kept:       DSCR 14,691 DY 17,866  LTV 14,293 -> binding 14,293, gap 2,207
Current loan: $16.5M at ~9.0% I/O after the Mar-26 extension -> interest 1,485; due Mar-27.
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 13

LOAN, INT9 = 16500, 16500 * 0.09
NOI = {"base": 937.4, "kept": 1429.3}
r = 0.0675 / 12
K = 12 * r / (1 - (1 + r) ** -360)


def sizes(noi):
    return [noi / 1.25 / K, noi / 0.08, 0.65 * noi / 0.065]


SB, SK = sizes(NOI["base"]), sizes(NOI["kept"])
GAP_B, GAP_K = LOAN - min(SB), LOAN - min(SK)
assert round(min(SB)) == 9374 and round(GAP_B) == 7126 and round(GAP_K) == 2207

prs = Presentation(SRC)
g = THEME["grid"]
s = new_content_slide(
    prs, "The 2027 Cliff and Refinancing Gap",
    "At $0.94M of 2027 NOI a new loan sizes to about $9.4M against $16.5M due in March 2027, a $7.1M gap")
add_slide_number(s)

# ------------------------------------------------------------------ left: maximum loan by sizing test
L0, LW = g["L0"], g["LP1"] - g["L0"]
section_header(s, L0, g["ROW1"], LW, "Maximum New Loan by Lender Test ($M)", "A header", sup="(1)(2)")
caption(s, L0, g["ROW1"] + 0.36, LW, "2027 NOI with no relief vs. with the exemption kept; 6.75% rate", "A caption")
# legend
for k, (lab, colr) in enumerate([("No relief (base case)", "accent_fill"), ("Exemption kept", "grey_lt")]):
    lx = L0 + 4.05 + k * 1.85
    box(s, lx, 2.02, 0.14, 0.14, fill=colr, line=None, shape=MSO_SHAPE.RECTANGLE, name=f"A legend key {k+1}")
    textbox(s, lx + 0.2, 1.98, 1.6, 0.22, [[(lab, {})]], size=10, name=f"A legend {k+1}")

cats = ["Debt coverage", "Debt yield", "Loan to value"]
vb = [round(v / 1000, 1) for v in SB]
vk = [round(v / 1000, 1) for v in SK]
P = Plot(0.75, 2.45, 5.85, 2.75, n=3, vmin=0, vmax=19.5)
GAP = 80
chart(s, P, cats, [("No relief", vb), ("Exemption kept", vk)],
      [bar_ser(0, "No relief", cats, vb, "accent_fill"), bar_ser(1, "Exemption kept", cats, vk, "grey_lt")],
      "A sizing chart", grouping="clustered", gap=GAP, overlap=-8)
line(s, P.x, P.v(0), P.x + P.w, P.v(0), "grey_mid", 0.75, name="A baseline")
# bar geometry: two bars per slot
bw = P.slot / (2 + 0.08 + GAP / 100)   # clustered, 2 series, overlap -8
for k in range(3):
    cx = P.center(k)
    for j, (v, colr) in enumerate([(vb[k], "accent"), (vk[k], "ink")]):
        bx = cx + (j - 0.5) * bw * 1.08
        textbox(s, bx - 0.45, P.v(v) - 0.23, 0.9, 0.2, [[(f"${v:.1f}M", {"b": True, "c": colr})]], size=10.5,
                align=PP_ALIGN.CENTER, name=f"A label {k+1}-{j+1}")
# claim line: amount owed
yl = P.v(LOAN / 1000)
line(s, P.x - 0.1, yl, P.x + P.w + 0.1, yl, "accent", 1.25, dash=MSO_LINE_DASH_STYLE.DASH, name="A owed line")
textbox(s, P.x - 0.1, yl - 0.27, 3.0, 0.22, [[("$16.5M owed, due Mar-27", {"b": True, "c": "accent"})]], size=10.5,
        name="A owed label", fill="white")
# gap brace on the binding test (loan to value, no relief), drawn right of the bars
brace_x = P.x + P.w + 0.12
line(s, P.center(2) + 0.54 * bw + bw / 2 + 0.03, P.v(vb[2]), brace_x, P.v(vb[2]), "accent", 0.75,
     dash=MSO_LINE_DASH_STYLE.ROUND_DOT, name="A gap leader")
brace(s, brace_x, yl, P.v(vb[2]), "A gap brace", width=0.1, color="accent")
textbox(s, brace_x + 0.17, (yl + P.v(vb[2])) / 2 - 0.2, 1.05, 0.42,
        [[("$7.1M gap", {"b": True, "c": "accent"})], [("lowest test", {"c": "grey_txt", "i": True})]], size=10.5,
        name="A gap label")
sub = ["1.25x at 6.75%, 30-yr amortization", "8.0% of NOI", "65% of value at a 6.5% cap"]
for k, (cat, sb) in enumerate(zip(cats, sub)):
    textbox(s, P.center(k) - 1.1, 5.3, 2.2, 0.44, [[(cat, {"b": True})], [(sb, {"c": "grey_txt", "i": True})]],
            size=10, align=PP_ALIGN.CENTER, name=f"A category {k+1}")

# ------------------------------------------------------------------ right: coverage on the current loan
R0, RW = g["RP0"], g["R1"] - g["RP0"]
section_header(s, R0, g["ROW1"], RW, "2027 on the Current Loan ($000s)", "B header", sup="(3)")
rows = [
    ["2027 NOI", "$937", "$1,429"],
    ["Interest at 9.0%", "$1,485", "$1,485"],
    ["Shortfall after interest", "($548)", "($56)"],
    ["Interest coverage", "0.63x", "0.96x"],
    ["Value at a 6.5% cap", "$14,422", "$21,989"],
    ["Loan to value", "114%", "75%"],
]
yb = grid_table(s, R0, 1.74, [2.05, 1.25, 1.33], ["", "No relief", "Exemption kept"], rows, "B table",
                size=10.5, row_h=0.30, hair_rules=True, bold_rows=(3,), rule_before=(3,),
                accent_cols={(2, 1), (3, 1), (5, 1)})

section_header(s, R0, yb + 0.2, RW, "How the Gap Gets Filled", "C header")
items = [
    [[("New money: ", {"b": True}), ("about $7.1M from the common or pref holders", {})]],
    [[("Another extension: ", {"b": True}), ("only if W Financial agrees; interest still exceeds NOI", {})]],
    [[("Sale or hand-back: ", {"b": True}), ("a $14.4M value sits below the $16.5M loan", {})]],
]
numbered_items(s, R0, yb + 0.66, RW, items, [0.36, 0.36, 0.36], name="C item", size=10.5, num_d=0.22, gap=0.08)

# ------------------------------------------------------------------ closing, footer, notes
closing_line(s, [("Even with the exemption a $2.2M gap remains; without it, ", {}),
                 ("the loan exceeds the value of the property", {"b": True})])
footer_notes(s,
             "Notes: (1) No relief = T12 NOI less pre-HFC taxes. (2) 6.75% is TexSun's target; 1.25x, 8.0% and 65% are team assumptions. (3) $16.5M at ~9.0% interest only.",
             "Sources: TexSun T12 and Q1-26 commentary; Fund annual report 2025-26; team analysis.")

speaker_notes(s, """Presenter: TBD

The bridge loan comes due in March 2027, so the real question is what a new lender would lend against this property.

On the left, three standard tests a lender uses. In red, our base case where the exemption ends and NOI is about $940 thousand. A lender that wants 1.25 times coverage at 6.75% lends about $9.6 million. One that wants an 8% debt yield lends $11.7 million. One that lends 65% of value lends $9.4 million. The lowest test wins, so call it $9.4 million against $16.5 million owed. That is a $7.1 million hole.

In grey, the good case where the exemption survives. It helps, but even then the loan-to-value test leaves about a $2.2 million gap.

On the right, the year itself. At 9% the loan costs about $1.5 million a year. With no relief, NOI covers 63% of the interest and the property runs about $550 thousand short. At a 6.5% cap the building is worth about $14.4 million, so the loan is already 114% of value.

There are three ways to close that gap: someone puts in about $7 million of new money, the lender agrees to extend again, or the property is sold or handed back, and a sale does not cover the loan. That is why we look at value and recovery next.

Reference notes:
- Sizing: 6.75% rate, 30-year amortization (7.78% constant); DSCR 1.25x -> $9.6M; debt yield 8.0% -> $11.7M; 65% LTV at 6.5% cap -> $9.4M. Exemption kept: $14.7M, $17.9M, $14.3M.
- Interest-only sizing at 1.25x and 6.75% would be $11.1M (no relief) and $16.9M (kept): still short in the base case.
- TexSun Q1-26: talking to Trevian and others for a mid-6s to 7% bridge; agency not available yet (agency in the low 5s).
- The annual report values the property at ~$19.2M on a 6.5% cap of recast NOI; our base case is lower because it removes the exemption.""")

move_slide(prs, s, POS)
prs.save(DST)
print("saved", DST, "at position", POS + 1)
