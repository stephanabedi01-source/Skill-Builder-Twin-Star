"""Add the Valuation Summary slide (slide 7, ahead of the Recommendation) to Project Print deck v9.

Usage: python3 build_valsum.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v15 ($XRX Valuation FF, Capital Structure).
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
slide.placeholders[21].text_frame.text = "Valuation Summary"
slide.placeholders[16].text_frame.text = (
    "Xerox is worth $3.1B to $3.7B against $4.4B of debt: equity is out of the money and value breaks in the "
    "unsecured notes")

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
         "Notes: (1) Peer Q1 to median, 7 print peers; (2) Q1 to median, 10 print deals; (3) WACC 12.0% +/- 1.0%, "
         "exit 5.0x +/- 0.5x; (4) 0% growth +/- 0.5%; (5) Ch. 7 net proceeds; (6) Debt at market plus 52-week "
         "equity range, less cash; (7) Weighted 30/20/25/25; (8) Excl. $105mm IEEPA benefit")
ftr_para(tf.add_paragraph(),
         "Sources: FactSet as of 10/14/26; Xerox Q2'26 10-Q; ARC, EFI and Lexmark merger proxies; Project Print model")

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


def pill(x, y, w, h, runs, line_c, name, size=10):
    p = box(slide, x, y, w, h, fill=WHITE, line=line_c, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
            radius=0.5, name=name)
    p.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p.text_frame.margin_left = p.text_frame.margin_right = Inches(0.04)
    fill_tf(p.text_frame, [runs], size, align=PP_ALIGN.CENTER)
    return p


# ================================================================= LEFT: football field vs claims
section_header(slide, L0, ROW1, LP1 - L0, "Enterprise Value Range vs. Claims ($M)", "A header")
rows = [
    ("Trading comps", "1", "5.9x to 6.4x LTM EBITDA", 3635, 3937, True),
    ("Precedent transactions", "2", "4.9x to 6.4x LTM EBITDA", 3016, 3930, True),
    ("DCF: exit multiple", "3", "4.7x to 5.7x implied", 2879, 3491, True),
    ("DCF: perpetuity growth", "4", "4.3x to 5.3x implied", 2617, 3253, True),
    None,
    ("Liquidation value", "5", "3.4x to 4.7x, a floor", 2089, 2887, False),
    ("Market-implied TEV", "6", "4.1x to 4.7x at market", 2527, 2867, False),
]
cats = [r[0] if r else " " for r in rows]
lo = [r[3] if r else None for r in rows]
span = [r[4] - r[3] if r else None for r in rows]
C = dict(x=2.55, y=2.17, w=5.55, h=3.50, vmin=1500, vmax=4800)
N = len(rows)
slot = C["h"] / N
yc = [C["y"] + slot * (k + 0.5) for k in range(N)]
xv = lambda v: C["x"] + C["w"] * (v - C["vmin"]) / (C["vmax"] - C["vmin"])
SEL_LO, SEL_HI = 3050, 3664
SEC, TOT = 2049, 4355

# selected-range band sits behind the chart
band = box(slide, xv(SEL_LO), C["y"] - 0.30, xv(SEL_HI) - xv(SEL_LO), C["h"] + 0.30, fill=BAND, line=None,
           shape=MSO_SHAPE.RECTANGLE, name="A selected range band")
dpt = {k: GREY_LT for k, r in enumerate(rows) if r and not r[5]}


def ff_xml(rid):
    base = (f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>{str_cache(["Low"])}'
            f'</c:strRef></c:tx><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>'
            f'{dlbls("", None, 9, INK, show=False)}'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${N + 1}</c:f>{str_cache(cats)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!$B$2:$B${N + 1}</c:f>{num_cache(lo)}</c:numRef></c:val></c:ser>')
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>' + base
            + bar_ser(1, "Range", cats, span, INK, dlbls("", None, 9, INK, show=False), dpt_fills=dpt)
            + '<c:gapWidth val="55"/><c:overlap val="100"/><c:axId val="2101"/><c:axId val="2102"/></c:barChart>'
            + f'<c:catAx><c:axId val="2101"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
              f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
              f'<c:crossAx val="2102"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
              f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            + f'<c:valAx><c:axId val="2102"/><c:scaling><c:orientation val="minMax"/><c:max val="{C["vmax"]}"/>'
              f'<c:min val="{C["vmin"]}"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
              f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
              f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2101"/>'
              f'<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, C["x"], C["y"], C["w"], C["h"], cats, [("Low", lo), ("Range", span)], ff_xml,
          "A football field chart")

# row labels and end values
for k, r in enumerate(rows):
    if r is None:
        textbox(slide, L0, yc[k] - 0.1, 2.2, 0.2, [[("Reference points (not weighted)", {"i": True, "c": GREY_TXT})]],
                size=9.5, name="A reference header")
        line(slide, L0, yc[k] + 0.13, LP1, yc[k] + 0.13, HAIR, 0.5, name="A reference rule")
        continue
    name, sup, sub, a, b, core = r
    textbox(slide, L0, yc[k] - 0.2, 2.25, 0.4,
            [[(name, {"b": core, "c": INK if core else GREY_TXT}), (sup, {"sup": True, "c": INK if core else GREY_TXT})],
             [(sub, {"c": GREY_TXT, "size": 9})]], size=10.5, anchor=MSO_ANCHOR.MIDDLE, name=f"A label {k + 1}")
    c = INK if core else GREY_TXT
    if core:
        textbox(slide, xv(a) - 0.72, yc[k] - 0.1, 0.66, 0.2, [[(f"${a:,}", {"b": True, "c": c})]], size=10,
                align=PP_ALIGN.RIGHT, name=f"A low value {k + 1}")
    else:   # reference bars: low value inside the bar, clear of the secured-debt line
        textbox(slide, xv(a) + 0.05, yc[k] - 0.1, 0.66, 0.2, [[(f"${a:,}", {"c": INK})]], size=10,
                name=f"A low value {k + 1}")
    textbox(slide, xv(b) + 0.06, yc[k] - 0.1, 0.66, 0.2, [[(f"${b:,}", {"b": core, "c": c})]], size=10,
            name=f"A high value {k + 1}")

# claim lines
top = C["y"] - 0.30
bot = C["y"] + C["h"] + 0.04
for v, c, lbl, nm in ((SEC, GREY_MID, f"Secured debt ${SEC:,}", "secured"), (TOT, RED, f"Total debt ${TOT:,}", "total")):
    line(slide, xv(v), top, xv(v), bot, c, 1.25, dash=MSO_LINE_DASH_STYLE.DASH, name=f"A {nm} debt line")
    textbox(slide, xv(v) - 0.9, top - 0.21, 1.8, 0.18, [[(lbl, {"b": True, "c": c})]], size=10,
            align=PP_ALIGN.CENTER, name=f"A {nm} debt label")
textbox(slide, xv(SEL_LO), top + 0.02, xv(SEL_HI) - xv(SEL_LO), 0.36,
        [[("Selected range", {"b": True})], [(f"${SEL_LO:,} to ${SEL_HI:,}", {"size": 9.5}), ("7", {"sup": True, "size": 9.5})]],
        size=10, align=PP_ALIGN.CENTER, name="A selected range label")
# unsecured bracket under the chart
by = bot + 0.12
line(slide, xv(SEC), by, xv(TOT), by, RED, 1.0, name="A unsecured bracket")
for v in (SEC, TOT):
    line(slide, xv(v), by - 0.06, xv(v), by + 0.06, RED, 1.0, name="A unsecured bracket tick")
pill(xv(SEC) + (xv(TOT) - xv(SEC)) / 2 - 1.65, by + 0.1, 3.3, 0.27,
     [("Unsecured notes ($2,306 face): ", {}), ("the fulcrum", {"b": True, "c": RED})], RED, "A fulcrum pill")

# ================================================================= RIGHT TOP: selected range and multiples
section_header(slide, RP0, ROW1, R1 - RP0, "Selected Range and Implied Multiples", "B header")
CW = [1.84, 0.87, 0.87, 0.88]
TX = RP0


def table(y0, hdr, body, name, bold_rows=(), red_rows=(), rule_before=()):
    x = TX
    for j, h in enumerate(hdr):
        textbox(slide, x, y0, CW[j], 0.2, [[(h, {"b": True})]], size=10.5,
                align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, name=f"{name} hdr {j + 1}")
        x += CW[j]
    line(slide, TX, y0 + 0.25, TX + sum(CW), y0 + 0.25, INK, 0.75, name=f"{name} hdr rule")
    y = y0 + 0.31
    for k, row in enumerate(body):
        if k in rule_before:
            line(slide, TX, y - 0.03, TX + sum(CW), y - 0.03, INK, 0.75, name=f"{name} total rule")
        x = TX
        for j, cell in enumerate(row):
            runs = cell if isinstance(cell, list) else [(cell, {})]
            o = {"b": k in bold_rows, "c": RED if (k in red_rows and j > 0) else INK}
            runs = [(t, {**o, **oo}) for t, oo in runs]
            textbox(slide, x, y, CW[j], 0.2, [runs], size=10.5, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT,
                    anchor=MSO_ANCHOR.MIDDLE, name=f"{name} r{k + 1} c{j + 1}")
            x += CW[j]
        y += 0.26
    return y


yb = table(1.72, ["($M)", "Low", "Mid", "High"],
           [["Enterprise value", "$3,050", "$3,357", "$3,664"],
            [[("EV / LTM Adj. EBITDA", {}), ("8", {"sup": True})], "5.0x", "5.5x", "6.0x"],
            ["EV / 2027E Adj. EBITDA", "4.2x", "4.6x", "5.0x"],
            ["EV / total debt", "0.70x", "0.77x", "0.84x"]],
           "B table", bold_rows=(0,), red_rows=(3,))

# ================================================================= RIGHT BOTTOM: equity is out of the money
section_header(slide, RP0, 3.20, R1 - RP0, "Equity Is Out of the Money", "C header")
ye = table(3.62, ["($M)", "Low", "Mid", "High"],
           [["Enterprise value", "$3,050", "$3,357", "$3,664"],
            ["(-) Total debt", "(4,355)", "(4,355)", "(4,355)"],
            ["(+) Cash", "495", "495", "495"],
            ["(-) Series A preferred", "(180)", "(180)", "(180)"],
            ["Implied equity value", "($990)", "($683)", "($376)"]],
           "C table", bold_rows=(4,), red_rows=(4,), rule_before=(4,))
callout(RP0, ye + 0.08, R1 - RP0, 0.62,
        [("Market cap is ", {}), ("$324M", {"b": True}), (" but value is ", {}), ("negative", {"b": True, "c": RED}),
         (": each class of claims is paid from enterprise value, not from the share price", {})],
        "C callout", size=10)

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Valuation conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Value covers secured debt ", {}), ("~1.5x", {"b": True}), (" even at the low end, but falls ", {}),
    ("$0.7B to $1.3B short", {"b": True, "c": RED}), (" of total debt: ", {}),
    ("the unsecured notes are the fulcrum and should own the reorganized company", {"i": True})]],
    11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 7 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(6, new)
prs.save(OUT)
print("saved", OUT)
