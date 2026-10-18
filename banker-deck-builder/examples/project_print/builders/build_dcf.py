"""Add the DCF slide (slide 10, ahead of the Recommendation) to Project Print deck v12.

Usage: python3 build_dcf.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v16 ($XRX DCF: FCF build, reorganized WACC, EV sensitivity grid).
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
slide.placeholders[21].text_frame.text = "Discounted Cash Flow"
slide.placeholders[16].text_frame.text = (
    "Even crediting the full base case, the DCF values Xerox at $2.9B to $3.2B, $1.2B to $1.5B short of its debt")

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
         "Notes: (1) Only the 10/16 to 12/31/26 stub (21%) is valued; (2) Exit 5.0x of 2030E Adj. EBITDA ($670M); "
         "(3) Claims at face, 6/30/26; (4) Reorganized company: cost of equity 14.9% (beta 1.27, 1.5% size and 2.0% "
         "company-specific premia), after-tax cost of debt 7.6% (B / B2, 5.0% spread), 55% equity / 45% debt")
ftr_para(tf.add_paragraph(),
         "Sources: Project Print model ($XRX DCF); FactSet as of 10/14/26; Kroll cost of capital data")

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


# ================================================================= LEFT TOP: cash flows
section_header(slide, L0, ROW1, LP1 - L0, "Unlevered Free Cash Flow, Base Case ($M)", "A header")
yrs = ["2026E", "2027E", "2028E", "2029E", "2030E"]
ebitda = [711.6, 727.6, 718.8, 693.8, 670.0]
ufcf = [302.1, 306.9, 344.6, 340.5, 328.8]
A = dict(x=1.55, y=1.98, w=6.60, h=1.30, vmin=0, vmax=860)
aslot = A["w"] / len(yrs)
ax = [A["x"] + aslot * (k + 0.5) for k in range(len(yrs))]
ay = lambda v: A["y"] + A["h"] * (A["vmax"] - v) / (A["vmax"] - A["vmin"])

# legend
line(slide, 4.60, 1.76, 4.90, 1.76, INK, 1.5, name="A legend line")
textbox(slide, 4.97, 1.66, 1.2, 0.2, [[("Adj. EBITDA", {"c": GREY_TXT})]], size=10, name="A legend ebitda")
box(slide, 6.25, 1.70, 0.14, 0.14, fill=INK, line=None, shape=MSO_SHAPE.RECTANGLE, name="A legend bar")
textbox(slide, 6.46, 1.66, 1.8, 0.2, [[("Unlevered free cash flow", {"c": GREY_TXT})]], size=10,
        name="A legend ufcf")

dpt = {0: GREY_LT}


def fcf_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "Unlevered FCF", yrs, ufcf, INK, dlbls("", None, 9, INK, show=False), dpt_fills=dpt)
            + '<c:gapWidth val="150"/><c:axId val="2501"/><c:axId val="2502"/></c:barChart>'
            + '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            + line_ser(1, "Adj. EBITDA", yrs, ebitda, INK, dlbls("", None, 9, INK, show=False), marker="circle",
                       msize=6, mfill=WHITE)
            + '<c:marker val="1"/><c:axId val="2501"/><c:axId val="2502"/></c:lineChart>'
            + cat_ax(2501, 2502, delete=True) + val_ax(2502, 2501, A["vmin"], A["vmax"]))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, A["x"], A["y"], A["w"], A["h"], yrs, [("Unlevered FCF", ufcf), ("Adj. EBITDA", ebitda)], fcf_xml,
          "A cash flow chart")
line(slide, A["x"], ay(0), A["x"] + A["w"], ay(0), GREY_MID, 0.75, name="A baseline")
for k in range(len(yrs)):
    textbox(slide, ax[k] - 0.4, ay(ebitda[k]) - 0.25, 0.8, 0.18, [[(f"${ebitda[k]:,.0f}", {"b": True})]], size=10,
            align=PP_ALIGN.CENTER, name=f"A ebitda value {k + 1}")
    textbox(slide, ax[k] - 0.4, ay(ufcf[k]) + 0.05, 0.8, 0.18,
            [[(f"${ufcf[k]:,.0f}", {"b": True, "c": INK if k == 0 else WHITE})]], size=10, align=PP_ALIGN.CENTER,
            name=f"A ufcf value {k + 1}")
    runs = [(yrs[k], {"b": True})] + ([("1", {"sup": True})] if k == 0 else [])
    textbox(slide, ax[k] - 0.5, ay(0) + 0.04, 1.0, 0.18, [runs], size=10, align=PP_ALIGN.CENTER,
            name=f"A year {k + 1}")
    textbox(slide, ax[k] - 0.5, ay(0) + 0.27, 1.0, 0.18, [[(f"{ufcf[k] / ebitda[k]:.0%}", {"i": True})]], size=10,
            align=PP_ALIGN.CENTER, name=f"A conversion {k + 1}")
line(slide, L0, ay(0) + 0.25, LP1, ay(0) + 0.25, HAIR, 0.5, name="A conversion rule")
textbox(slide, L0, ay(0) + 0.27, 1.3, 0.18, [[("FCF / EBITDA", {"i": True, "c": GREY_TXT})]], size=10,
        name="A conversion label")
textbox(slide, L0, A["y"] + 0.15, 1.25, 0.8,
        [[("Peaks in 2027", {"b": True})], [("and less than half of EBITDA becomes free cash flow", {"c": GREY_TXT})]],
        size=10, name="A side note")

# ================================================================= LEFT BOTTOM: value vs claims
B0 = 3.95
section_header(slide, L0, B0, LP1 - L0, "DCF Value vs. Claims ($M)", "B header")
PVF, TVX, TVG = 1096, 2079, 1800
SEC, UNS, TOT = 2049, 2306, 4355
Bc = dict(x=2.25, y=B0 + 0.78, w=5.60, h=1.40, vmin=0, vmax=4700)
bx = lambda v: Bc["x"] + Bc["w"] * (v - Bc["vmin"]) / (Bc["vmax"] - Bc["vmin"])
rows = [("DCF: exit multiple", "2", "5.0x of 2030E EBITDA", [(PVF, INK), (TVX, GREY_LT)]),
        ("DCF: perpetuity growth", "", "0% growth", [(PVF, INK), (TVG, GREY_LT)]),
        ("Claims", "3", "Debt at face", [(SEC, GREY_MID), (UNS, None)])]
bslot = Bc["h"] / len(rows)
byc = [Bc["y"] + bslot * (k + 0.5) for k in range(len(rows))]
BH = bslot * 0.58
cats = [r[0] for r in rows]
s1 = [r[3][0][0] for r in rows]
s2 = [r[3][1][0] for r in rows]


def stack_xml(rid):
    first = bar_ser(0, "Senior piece", cats, s1, INK, dlbls("", None, 9, INK, show=False),
                    dpt_fills={2: GREY_MID})
    second = (f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f>'
              f'{str_cache(["Second piece"])}</c:strRef></c:tx>'
              f'<c:spPr><a:solidFill><a:srgbClr val="{GREY_LT}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
              f'<c:invertIfNegative val="0"/>'
              f'<c:dPt><c:idx val="2"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:noFill/>'
              f'<a:ln w="15875"><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill><a:prstDash val="dash"/></a:ln>'
              f'</c:spPr></c:dPt>{dlbls("", None, 9, INK, show=False)}'
              f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f>{str_cache(cats)}</c:strRef></c:cat>'
              f'<c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>{num_cache(s2)}</c:numRef></c:val></c:ser>')
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>' + first + second
            + '<c:gapWidth val="72"/><c:overlap val="100"/><c:axId val="2601"/><c:axId val="2602"/></c:barChart>'
            + f'<c:catAx><c:axId val="2601"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
              f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
              f'<c:crossAx val="2602"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
              f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            + f'<c:valAx><c:axId val="2602"/><c:scaling><c:orientation val="minMax"/><c:max val="{Bc["vmax"]}"/>'
              f'<c:min val="{Bc["vmin"]}"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
              f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
              f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2601"/>'
              f'<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, Bc["x"], Bc["y"], Bc["w"], Bc["h"], cats, [("Senior piece", s1), ("Second piece", s2)], stack_xml,
          "B value vs claims chart")
line(slide, Bc["x"], Bc["y"] - 0.02, Bc["x"], Bc["y"] + Bc["h"] + 0.02, GREY_MID, 0.75, name="B axis")
# total debt line
line(slide, bx(TOT), Bc["y"] - 0.20, bx(TOT), Bc["y"] + Bc["h"] + 0.02, RED, 1.25, dash=MSO_LINE_DASH_STYLE.DASH,
     name="B total debt line")
textbox(slide, bx(TOT) - 0.9, Bc["y"] - 0.40, 1.8, 0.18, [[(f"Total debt ${TOT:,}", {"b": True, "c": RED})]],
        size=10, align=PP_ALIGN.CENTER, name="B total debt label")

seg_lbl = {0: [("PV of FCF ", {}), (f"${PVF:,}", {"b": True})],
           1: [("PV of terminal value ", {}), (f"${TVX:,}", {"b": True}), (f" ({TVX / (PVF + TVX):.0%})", {})],
           2: [("Secured ", {}), (f"${SEC:,}", {"b": True})]}
for k, (name, sup, sub, segs) in enumerate(rows):
    textbox(slide, L0, byc[k] - 0.2, 1.95, 0.4,
            [[(name, {"b": True}), (sup, {"sup": True})], [(sub, {"c": GREY_TXT, "size": 9})]], size=10.5,
            anchor=MSO_ANCHOR.MIDDLE, name=f"B label {k + 1}")
    a, b = segs[0][0], segs[1][0]
    first = seg_lbl[0] if k < 2 else seg_lbl[2]
    textbox(slide, bx(0) + 0.06, byc[k] - 0.1, bx(a) - bx(0) - 0.1, 0.2,
            [[(t, {**o, "c": WHITE}) for t, o in first]], size=9.5, name=f"B seg1 {k + 1}")
    if k == 0:
        second = seg_lbl[1]
    elif k == 1:
        second = [("PV of terminal value ", {}), (f"${TVG:,}", {"b": True}), (f" ({TVG / (PVF + TVG):.0%})", {})]
    else:
        second = [("Unsecured notes ", {"c": RED}), (f"${UNS:,}", {"b": True, "c": RED})]
    textbox(slide, bx(a) + 0.06, byc[k] - 0.1, bx(a + b) - bx(a) - 0.1, 0.2, [second], size=9.5,
            name=f"B seg2 {k + 1}")
    if k < 2:
        ev = a + b
        textbox(slide, bx(ev) + 0.05, byc[k] - 0.1, 0.62, 0.2, [[(f"${ev:,}", {"b": True})]], size=10.5,
                name=f"B total {k + 1}")
        # shortfall bracket from EV to total debt
        y = byc[k] + BH / 2 + 0.05
        line(slide, bx(ev), y, bx(TOT), y, RED, 1.0, name=f"B shortfall {k + 1}")
        textbox(slide, bx(TOT) + 0.06, byc[k] - 0.1, 0.9, 0.2, [[(f"(${TOT - ev:,})", {"b": True, "c": RED})]],
                size=10.5, name=f"B shortfall label {k + 1}")

# ================================================================= RIGHT TOP: sensitivity
section_header(slide, RP0, ROW1, R1 - RP0, "No Reasonable Case Covers the Debt", "C header")
textbox(slide, RP0, 1.66, R1 - RP0, 0.2, [[("Enterprise value ($M) by WACC and exit multiple", {"i": True, "c": GREY_TXT})]],
        size=10, name="C caption")
mults = [4.0, 4.5, 5.0, 5.5, 6.0]
waccs = [0.14, 0.13, 0.12, 0.11, 0.10]
grid = [[2602, 2795, 2988, 3181, 3374],
        [2679, 2879, 3080, 3280, 3480],
        [2759, 2967, 3175, 3383, 3591],
        [2843, 3059, 3275, 3491, 3707],
        [2930, 3155, 3379, 3603, 3827]]
GX0 = RP0 + 0.62
GW = (R1 - GX0) / 5
GY0 = 1.98
GH = 0.25
for j, m in enumerate(mults):
    textbox(slide, GX0 + GW * j, GY0, GW, 0.2, [[(f"{m:.1f}x", {"b": True})]], size=10, align=PP_ALIGN.CENTER,
            name=f"C col {j + 1}")
line(slide, RP0, GY0 + 0.24, R1, GY0 + 0.24, INK, 0.75, name="C hdr rule")
textbox(slide, RP0, GY0, 0.6, 0.2, [[("WACC", {"b": True})]], size=10, name="C wacc hdr")
for i, w in enumerate(waccs):
    y = GY0 + 0.30 + GH * i
    base_row = abs(w - 0.12) < 1e-9
    textbox(slide, RP0, y, 0.6, GH - 0.03, [[(f"{w:.1%}", {"b": base_row})]], size=10, anchor=MSO_ANCHOR.MIDDLE,
            name=f"C row {i + 1}")
    for j in range(5):
        base = base_row and j == 2
        c = box(slide, GX0 + GW * j + 0.02, y, GW - 0.04, GH - 0.03, fill=INK if base else BAND, line=None,
                shape=MSO_SHAPE.RECTANGLE, name=f"C cell {i + 1}-{j + 1}")
        c.text_frame.margin_top = c.text_frame.margin_bottom = 0
        c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        fill_tf(c.text_frame, [[(f"${grid[i][j]:,}", {"b": base, "c": WHITE if base else INK})]], 10,
                align=PP_ALIGN.CENTER)
yg = GY0 + 0.30 + GH * 5
textbox(slide, RP0, yg + 0.03, R1 - RP0, 0.2,
        [[("Every cell is below the ", {"c": GREY_TXT}), (f"${TOT:,}", {"b": True, "c": RED}),
          (" of debt; base case highlighted", {"c": GREY_TXT})]], size=10, name="C grid note")
callout(RP0, yg + 0.30, R1 - RP0, 0.50,
        [("Covering the debt takes a ", {}), ("7.8x exit multiple", {"b": True, "c": RED}),
         (", above every peer and every print deal since 2020", {})], "C callout", size=10.5)

# ================================================================= RIGHT BOTTOM: assumptions
section_header(slide, RP0, 4.45, R1 - RP0, "Key Assumptions", "D header")
CW = [1.55, 0.70, 2.24]
rows_a = [("WACC", "12.0%", [("Reorganized company", {}), ("4", {"sup": True})]),
          ("Exit multiple", "5.0x", [("2020-24 precedent median 4.7x", {})]),
          ("Perpetuity growth", "0.0%", [("Shrinking print market", {})]),
          ("Terminal value / EV", "65%", [("62% under perpetuity growth", {})]),
          ("Valuation date", "10/16/26", [("Mid-year convention, 25% tax", {})])]
y = 4.86
for k, (a, b, c) in enumerate(rows_a):
    textbox(slide, RP0, y, CW[0], 0.2, [[(a, {})]], size=10.5, anchor=MSO_ANCHOR.MIDDLE, name=f"D r{k + 1} c1")
    textbox(slide, RP0 + CW[0], y, CW[1], 0.2, [[(b, {"b": True})]], size=10.5, align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE, name=f"D r{k + 1} c2")
    textbox(slide, RP0 + CW[0] + CW[1] + 0.15, y, CW[2] - 0.15, 0.2, [[(t, {**o, "c": GREY_TXT}) for t, o in c]],
            size=10, anchor=MSO_ANCHOR.MIDDLE, name=f"D r{k + 1} c3")
    if k < len(rows_a) - 1:
        line(slide, RP0, y + 0.245, R1, y + 0.245, HAIR, 0.5, name=f"D rule {k + 1}")
    y += 0.29

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="DCF conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Even with full credit for the plan, the DCF leaves Xerox ", {}), ("$1.2B to $1.5B short", {"b": True, "c": RED}),
    (" and breaks even only at a 7.8x exit: ", {}), ("the balance sheet has to be resized to the business", {"i": True})]],
    11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 10 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(9, new)
prs.save(OUT)
print("saved", OUT)
