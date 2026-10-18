"""Add the Liquidation Analysis slide (slide 11, ahead of the Recommendation) to Project Print deck v13.

Usage: python3 build_liquidation.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v16 ($XRX Liquidation Analysis mid case; $XRX Valuation FF; $XRX Transaction Analysis).
"""
import sys

_argv = list(sys.argv)
sys.argv = sys.argv[:1]
from pp_helpers import *  # noqa: E402,F401,F403
from pp_helpers import textbox, section_header  # noqa: E402

SRC, OUT = _argv[1], _argv[2]
HAIR = "D1D1D1"
GREY_FILL = "D6D6D6"
BAND = "E9E9E9"

prs = Presentation(SRC)
ref = prs.slides[2]
slide = prs.slides.add_slide(ref.slide_layout)
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 20:
        ph._element.getparent().remove(ph._element)
slide.placeholders[21].text_frame.text = "Liquidation Analysis"
slide.placeholders[16].text_frame.text = (
    "Liquidation yields $2.1B to $2.9B and 2c to 13c for unsecured creditors: Xerox is worth more as a going concern")

spTree = slide.shapes._spTree
for sp in ref.shapes._spTree.iterchildren(qn("p:sp")):
    ph = sp.find(".//" + qn("p:ph"))
    if ph is not None and ph.get("type") in ("sldNum", "ftr"):
        spTree.append(copy.deepcopy(sp))
ftr = [s for s in slide.placeholders if s.placeholder_format.idx == 18][0]
ftr.top = Inches(6.93)
tf = ftr.text_frame
for p in list(tf.paragraphs)[1:]:
    p._p.getparent().remove(p._p)
for r in list(tf.paragraphs[0].runs):
    r._r.getparent().remove(r._r)


def ftr_para(p, text):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(10)


ftr_para(tf.paragraphs[0],
         "Notes: (1) Hypothetical Chapter 7, 6/30/26 book values, mid case (low / high: $2,089M / $2,887M); (2) Selected "
         "EV mid; (3) Costs 11% to 14% of gross proceeds; (4) ABL 60% drawn at filing plus accrued; IPCo JV debt repaid "
         "from brand IP held in the JV; (5) Prepack plan, Case B")
ftr_para(tf.add_paragraph(),
         "Sources: Xerox Q2'26 10-Q; Project Print model ($XRX Liquidation Analysis, $XRX Transaction Analysis)")

L0, R1 = 0.23, 13.11
LP1 = 8.33              # left panel right edge
RP0 = 8.62              # right panel left edge
ROW1 = 1.30


def callout(x, y, w, h, runs, name, size=10.5):
    c = box(slide, x, y, w, h, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
            radius=0.08, name=name)
    c.text_frame.margin_left = c.text_frame.margin_right = Inches(0.08)
    c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(c.text_frame, [runs], size, align=PP_ALIGN.CENTER)
    return c


def nofill_ser(i, name, cats, vals):
    n = len(cats)
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col(i)}$1</c:f>{str_cache([name])}</c:strRef></c:tx>'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>'
            f'{dlbls("", None, 9, INK, show=False)}'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n + 1}</c:f>{str_cache(cats)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!${col(i)}$2:${col(i)}${n + 1}</c:f>{num_cache(vals)}</c:numRef></c:val>'
            f'</c:ser>')


NOLBL = dlbls("", None, 9, INK, show=False)

# ================================================================= LEFT: book value to net proceeds
textbox(slide, L0, ROW1, LP1 - L0, 0.24, [[("From $9.7B of Book Assets to $2.5B of Net Proceeds (Mid, $M)", {"b": True}),
                                           ("1", {"sup": True, "b": True})]], size=14, name="A header")
line(slide, L0, ROW1 + 0.29, LP1, ROW1 + 0.29, INK, 1.0, name="A header rule")

BOOK = 9724
# (two-line label, recovery range, book value, mid recovery)
steps = [(["Goodwill,", "intangibles"], "0% to 1%", 3091, 21.425),
         (["Other", "long-term"], "0% to 9%", 1543, 72.85),
         (["Inventory,", "other current"], "17% to 33%", 1535, 393.39),
         (["Receivables", "(trade, XFS)"], "60% to 80%", 2399, 1676.4),
         (["PP&E, lease", "equipment"], "23% to 37%", 661, 198.3)]
GROSS = 2857.365
COSTS = 357.171
NET = GROSS - COSTS
GC = 3357.25
cats = ["Book assets"] + [" ".join(s[0]).replace("&", "and") for s in steps] + ["Gross proceeds", "Liquidation costs", "Net proceeds",
                                                           "Going concern"]
N = len(cats)
base, down, tot = [None] * N, [None] * N, [None] * N
tot[0] = BOOK
run = BOOK
for k, (_, _, bv, rec) in enumerate(steps, start=1):
    loss = bv - rec
    run -= loss
    base[k], down[k] = run, loss
tot[6] = GROSS
base[7], down[7] = NET, COSTS
tot[8] = NET
tot[9] = GC
assert abs(run - GROSS) < 1, run
C = dict(x=L0, y=1.92, w=LP1 - L0, h=2.85, vmin=0, vmax=10400)
slot = C["w"] / N
xc = [C["x"] + slot * (k + 0.5) for k in range(N)]
yv = lambda v: C["y"] + C["h"] * (C["vmax"] - v) / (C["vmax"] - C["vmin"])


def wf_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
            + nofill_ser(0, "Base", cats, base)
            + bar_ser(1, "Value lost", cats, down, GREY_MID, NOLBL)
            + bar_ser(2, "Value", cats, tot, INK, NOLBL, dpt_fills={0: GREY_LT, 8: RED_BAR})
            + '<c:gapWidth val="40"/><c:overlap val="100"/><c:axId val="2701"/><c:axId val="2702"/></c:barChart>'
            + cat_ax(2701, 2702, delete=True) + val_ax(2702, 2701, C["vmin"], C["vmax"]))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, C["x"], C["y"], C["w"], C["h"], cats, [("Base", base), ("Value lost", down), ("Value", tot)], wf_xml,
          "A liquidation waterfall chart")
line(slide, C["x"], yv(0), C["x"] + C["w"], yv(0), GREY_MID, 0.75, name="A baseline")

vals = [(f"${BOOK:,}", BOOK, INK)]
vals += [(f"(${s[2] - s[3]:,.0f})", base[k + 1] + down[k + 1], GREY_TXT) for k, s in enumerate(steps)]
vals += [(f"${GROSS:,.0f}", GROSS, INK), (f"(${COSTS:,.0f})", NET + COSTS, GREY_TXT), (f"${NET:,.0f}", NET, RED),
         (f"${GC:,.0f}", GC, INK)]
for k, (t, top, c) in enumerate(vals):
    runs = [(t, {"b": True, "c": c})] + ([("3", {"sup": True, "c": c})] if k == 7 else []) \
        + ([("2", {"sup": True, "c": c})] if k == 9 else [])
    textbox(slide, xc[k] - 0.45, yv(top) - 0.23, 0.9, 0.2, [runs], size=10.5, align=PP_ALIGN.CENTER,
            name=f"A value {k + 1}")
lbls = [(["Book", "assets"], "6/30/26")] + [(s[0], s[1]) for s in steps] + \
       [(["Gross", "proceeds"], None), (["Wind-down,", "fees"], "12.5%"), (["Net", "proceeds"], None),
        (["Going", "concern"], None)]
for k, (two, rec) in enumerate(lbls):
    hot = k in (8, 9)
    c = RED if k == 8 else INK
    paras = [[(two[0], {"b": hot, "c": c if hot else GREY_TXT})], [(two[1], {"b": hot, "c": c if hot else GREY_TXT})]]
    if rec:
        paras.append([(rec, {"i": True, "c": GREY_TXT, "size": 9})])
    textbox(slide, xc[k] - slot / 2 - 0.05, yv(0) + 0.05, slot + 0.1, 0.5, paras, size=9.5, align=PP_ALIGN.CENTER,
            name=f"A category {k + 1}")

# value preserved by reorganizing
line(slide, xc[8] + slot * 0.3, yv(NET), xc[9] - slot * 0.3, yv(NET), INK, 0.75, dash=MSO_LINE_DASH_STYLE.DASH,
     name="A preserved link")
pp = box(slide, xc[9] - 0.95, yv(GC) - 0.62, 1.32, 0.30, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.ROUND_DOT,
         line_w=1.0, radius=0.5, name="A preserved pill")
pp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
pp.text_frame.margin_left = pp.text_frame.margin_right = Inches(0.02)
fill_tf(pp.text_frame, [[(f"+${GC - NET:,.0f}M", {"b": True}), (" preserved", {})]], 10, align=PP_ALIGN.CENTER)

callout(L0, 5.58, LP1 - L0, 0.62,
        [("71%", {"b": True, "c": RED}), (" of book value disappears in a fire sale: goodwill and intangibles are worth "
                                          "almost nothing, and ", {}),
         ("cash and receivables", {"b": True}), (" make up three-quarters of what is left", {})],
        "A callout", size=10.5)

# ================================================================= RIGHT TOP: who gets the proceeds
section_header(slide, RP0, ROW1, R1 - RP0, "Who Gets the $2.5B in Chapter 7", "B header")
ADMIN, SECURED = 125, 1807.6
POOL = NET - ADMIN - SECURED
NOTES, PENS, TRADE = 2345.5, 1175, 3682
CLAIMS = NOTES + PENS + TRADE
bcats = ["Net proceeds", "Unsecured claims"]
p1, p2, p3 = [ADMIN, NOTES], [SECURED, PENS], [POOL, TRADE]
B = dict(x=RP0, y=1.86, w=R1 - RP0, h=1.60)
bs = B["h"] / 2
BH = bs / 2.4                    # bar thickness at gapWidth 140
bar_top = [B["y"] + bs * (k + 0.5) - BH / 2 for k in range(2)]
fx = lambda frac: B["x"] + B["w"] * frac


def split_ser(i, name, vals, fills):
    d = "".join(f'<c:dPt><c:idx val="{k}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
                f'<c:spPr><a:solidFill><a:srgbClr val="{f}"/></a:solidFill>'
                f'<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>'
                for k, f in enumerate(fills))
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:f>Sheet1!${col(i)}$1</c:f>'
            f'{str_cache([name])}</c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="{fills[0]}"/></a:solidFill>'
            f'</c:spPr><c:invertIfNegative val="0"/>{d}{NOLBL}'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f>{str_cache(bcats)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!${col(i)}$2:${col(i)}$3</c:f>{num_cache(vals)}</c:numRef></c:val></c:ser>')


def split_xml(rid):
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>'
            + split_ser(0, "First", p1, [GREY_MID, INK])
            + split_ser(1, "Second", p2, [INK, GREY_MID])
            + split_ser(2, "Third", p3, [RED_BAR, GREY_LT])
            + '<c:gapWidth val="140"/><c:overlap val="100"/><c:axId val="2801"/><c:axId val="2802"/></c:barChart>'
            + f'<c:catAx><c:axId val="2801"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
              f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
              f'<c:crossAx val="2802"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
              f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            + f'<c:valAx><c:axId val="2802"/><c:scaling><c:orientation val="minMax"/><c:max val="1"/><c:min val="0"/>'
              f'</c:scaling><c:delete val="1"/><c:axPos val="b"/><c:numFmt formatCode="0%" sourceLinked="0"/>'
              f'<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
              f'<c:crossAx val="2801"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, B["x"], B["y"], B["w"], B["h"], bcats, [("First", p1), ("Second", p2), ("Third", p3)], split_xml,
          "B distribution chart")
# row 1 labels
textbox(slide, RP0, bar_top[0] - 0.24, 3.0, 0.2, [[("Net proceeds ", {"c": GREY_TXT}), (f"${NET:,.0f}", {"b": True})]],
        size=10, name="B row1 title")
f_adm, f_sec = ADMIN / NET, (ADMIN + SECURED) / NET
textbox(slide, fx(f_adm) + 0.05, bar_top[0], fx(f_sec) - fx(f_adm) - 0.1, BH,
        [[("Secured claims ", {"c": WHITE}), (f"${SECURED:,.0f}", {"b": True, "c": WHITE}), (" (100%)", {"c": WHITE}),
          ("4", {"sup": True, "c": WHITE})]], size=9.5, anchor=MSO_ANCHOR.MIDDLE, name="B secured label")
textbox(slide, fx(f_sec), bar_top[0], fx(1) - fx(f_sec), BH, [[(f"${POOL:,.0f}", {"b": True, "c": WHITE})]], size=9.5,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name="B pool label")
textbox(slide, fx(1) - 2.2, bar_top[0] - 0.24, 2.2, 0.2,
        [[("Admin ", {"c": GREY_TXT}), (f"${ADMIN}", {"b": True}), (" | Unsecured pool ", {"c": GREY_TXT}),
          (f"${POOL:,.0f}", {"b": True, "c": RED})]], size=10, align=PP_ALIGN.RIGHT, name="B row1 legend")
# funnel from the unsecured pool to the claims bar
y_a, y_b = bar_top[0] + BH, bar_top[1]
line(slide, fx(f_sec), y_a, fx(0), y_b, GREY_LT, 0.75, name="B funnel left")
line(slide, fx(1), y_a, fx(1), y_b, GREY_LT, 0.75, name="B funnel right")
textbox(slide, RP0, y_a + 0.04, 1.9, 0.2, [[("Shared by ", {"c": GREY_TXT}), (f"${CLAIMS:,.0f}", {"b": True}),
                                            (" of claims", {"c": GREY_TXT})]], size=10, name="B row2 title")
# row 2 labels
segs = [(NOTES, "Notes", WHITE), (PENS, "Pension", WHITE), (TRADE, "Trade and other", INK)]
x0 = 0
for v, nm, c in segs:
    a, b = fx(x0 / CLAIMS), fx((x0 + v) / CLAIMS)
    paras = ([[(nm, {"c": c, "size": 9})], [(f"${v:,.0f}", {"b": True, "c": c, "size": 9})]] if nm == "Pension"
             else [[(nm + " ", {"c": c}), (f"${v:,.0f}", {"b": True, "c": c})]])
    tb = textbox(slide, a, bar_top[1], b - a, BH, paras, size=9.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 name=f"B claims {nm}")
    tb.text_frame.word_wrap = False
    x0 += v
textbox(slide, RP0, bar_top[1] + BH + 0.06, R1 - RP0, 0.2,
        [[("Every unsecured class recovers ", {}), (f"~{POOL / CLAIMS * 100:.0f}c", {"b": True, "c": RED}),
          (" on the dollar (2c to 13c across cases)", {})]], size=10.5, align=PP_ALIGN.CENTER, name="B result")

# ================================================================= RIGHT BOTTOM: best-interests test
section_header(slide, RP0, 3.86, R1 - RP0, "Best-Interests Test: Every Class Does Better", "C header")
TX, TY = RP0, 4.28
CW = [1.95, 1.27, 1.27]
hdrs = [[("Class", {})], [("Chapter 7", {})], [("Plan", {}), ("5", {"sup": True})]]
rows = [("Secured (ABL, 1L, 2L)", "100%", "100%"),
        ("Holdings notes", "7.9%", "51.9%"),
        ("Legacy Xerox Corp. notes", "7.9%", "26.0%"),
        ("Pension and trade", "7.9%", "Unimpaired"),
        ("Preferred and common", "0%", "0%")]
x = TX
for j, h in enumerate(hdrs):
    textbox(slide, x, TY, CW[j], 0.2, [[(t, {"b": True, **o}) for t, o in h]], size=10.5,
            align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, name=f"C header {j + 1}")
    x += CW[j]
TW = sum(CW)
line(slide, TX, TY + 0.25, TX + TW, TY + 0.25, INK, 0.75, name="C header rule")
RH = 0.27
y = TY + 0.31
for k, (cls, a, b) in enumerate(rows):
    x = TX
    better = k in (1, 2, 3)
    cells = [[(cls, {})], [(a, {"c": RED if better else INK})], [(b, {"b": True})]]
    for j, runs in enumerate(cells):
        textbox(slide, x, y, CW[j], 0.2, [runs], size=10.5, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"C r{k + 1} c{j + 1}")
        x += CW[j]
    y += RH
    if k < len(rows) - 1:
        line(slide, TX, y - 0.04, TX + TW, y - 0.04, HAIR, 0.5, name=f"C rule {k + 1}")
box(slide, TX + CW[0] + CW[1] + 0.06, TY - 0.03, CW[2] - 0.02, y - TY - 0.02, fill=None, line=INK, line_w=0.75,
    radius=0.06, name="C plan column outline")

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Liquidation conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Liquidation is the floor, not an option: it destroys ", {}), ("~$0.9B of value", {"b": True, "c": RED}),
    (" and cuts noteholder recoveries from ~52c to ~8c, so ", {}),
    ("the plan clears the best-interests test with room to spare", {"i": True})]], 11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 11 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(10, new)
prs.save(OUT)
print("saved", OUT, "pool", round(POOL, 1), "claims", CLAIMS, "rec", round(POOL / CLAIMS, 4))
