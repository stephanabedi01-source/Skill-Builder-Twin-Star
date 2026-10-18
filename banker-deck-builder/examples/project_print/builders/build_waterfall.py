"""Add the Recovery Waterfall slide (slide 12, ahead of the Recommendation) to Project Print deck v14.

Usage: python3 build_waterfall.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v16 ($XRX Recovery Waterfall; Capital Structure bond prices as of 10/14/26).
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
slide.placeholders[21].text_frame.text = "Recovery Waterfall"
slide.placeholders[16].text_frame.text = (
    "Secured claims are covered in full, value breaks in the unsecured notes at 18c to 36c, and equity receives nothing")

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
         "Notes: (1) Mid EV $3,357M less $125M admin; claims include 3 months accrued, ABL 60% drawn; low / high EV "
         "$3,050M / $3,664M; (2) Repaid first from the JV's 17% share of EV incl. 3.875% premium; $72M residual to the "
         "estate; (3) Guarantee plus intercompany loan claims at Xerox Corp., capped at par; (4) FactSet, 10/14/26")
ftr_para(tf.add_paragraph(),
         "Sources: Xerox Q2'26 10-Q; indentures; FactSet; Project Print model ($XRX Recovery Waterfall)")

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


NOLBL = dlbls("", None, 9, INK, show=False)


def hbar_axes(a1, a2, vmax):
    return (f'<c:catAx><c:axId val="{a1}"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
            f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'<c:crossAx val="{a2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
            f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            f'<c:valAx><c:axId val="{a2}"/><c:scaling><c:orientation val="minMax"/><c:max val="{vmax}"/>'
            f'<c:min val="0"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
            f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
            f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{a1}"/>'
            f'<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')


# ================================================================= LEFT: waterfall by priority
textbox(slide, L0, ROW1, LP1 - L0, 0.24, [[("Waterfall by Priority at Mid Plan Value ($M)", {"b": True}),
                                           ("1", {"sup": True, "b": True})]], size=14, name="A header")
line(slide, L0, ROW1 + 0.29, LP1, ROW1 + 0.29, INK, 1.0, name="A header rule")
# (label, sup, claim, mid recovery $, mid %, low-to-high)
rows = [("Admin and priority claims", "", 125.0, 125.0, "100%", ""),
        ("IPCo JV debt (brand IP)", "2", 477.8, 477.8, "100%", ""),
        ("ABL revolver", "", 167.8, 167.8, "100%", ""),
        ("1L Term Loan B", "", 718.9, 718.9, "100%", ""),
        ("1L 10.25% notes", "", 410.3, 410.3, "100%", ""),
        ("2L 13.50% notes", "", 510.7, 510.7, "100%", ""),
        None,
        ("Holdings notes", "3", 1736.6, 625.5, "36%", "24% to 48%"),
        ("Legacy Xerox Corp. notes", "", 608.9, 109.7, "18%", "12% to 24%"),
        ("Pension / OPEB", "", 1175.0, 211.6, "18%", "12% to 24%"),
        ("Series A preferred", "", 180.0, 0.0, "0%", ""),
        ("Common equity", "", None, None, "0%", "")]
N = len(rows)
cats = [r[0] if r else " " for r in rows]
claim = [r[2] if r else None for r in rows]
rec = [r[3] if r else None for r in rows]
C = dict(x=2.40, y=1.95, w=3.15, h=0.30 * N, vmax=1850)
slot = C["h"] / N
yc = [C["y"] + slot * (k + 0.5) for k in range(N)]
xv = lambda v: C["x"] + C["w"] * v / C["vmax"]
COLS = [(5.60, 0.62, "Claim"), (6.22, 0.70, "Recovery"), (6.92, 0.52, "Mid %"), (7.44, 0.89, "Low to high")]

textbox(slide, C["x"], 1.66, 3.1, 0.2, [[("Claim (outline) vs. recovery (fill)", {"i": True, "c": GREY_TXT})]],
        size=10, name="A caption")
for x, w, h in COLS:
    textbox(slide, x, 1.66, w, 0.2, [[(h, {"b": True})]], size=10, align=PP_ALIGN.RIGHT, name=f"A col {h}")
dfill = {k: (INK if r[4] == "100%" else RED_BAR) for k, r in enumerate(rows) if r and r[3]}


def wf_xml(rid):
    claim_ser = (f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f>'
                 f'{str_cache(["Claim"])}</c:strRef></c:tx><c:spPr><a:noFill/><a:ln w="9525"><a:solidFill>'
                 f'<a:srgbClr val="{GREY_MID}"/></a:solidFill></a:ln></c:spPr><c:invertIfNegative val="0"/>{NOLBL}'
                 f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${N + 1}</c:f>{str_cache(cats)}</c:strRef></c:cat>'
                 f'<c:val><c:numRef><c:f>Sheet1!$C$2:$C${N + 1}</c:f>{num_cache(claim)}</c:numRef></c:val></c:ser>')
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "Recovery", cats, rec, INK, NOLBL, dpt_fills=dfill) + claim_ser
            + '<c:gapWidth val="45"/><c:overlap val="100"/><c:axId val="2901"/><c:axId val="2902"/></c:barChart>'
            + hbar_axes(2901, 2902, C["vmax"]))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, C["x"], C["y"], C["w"], C["h"], cats, [("Recovery", rec), ("Claim", claim)], wf_xml,
          "A waterfall chart")
line(slide, C["x"], C["y"], C["x"], C["y"] + C["h"], GREY_MID, 0.75, name="A axis")
for k, r in enumerate(rows):
    if r is None:
        continue
    lbl, sup, cl, rv, pct, rng = r
    unsec = k > 6
    hot = pct not in ("100%", "0%")
    textbox(slide, L0, yc[k] - 0.1, 2.15, 0.2, [[(lbl, {"b": hot}), (sup, {"sup": True})]], size=10.5,
            anchor=MSO_ANCHOR.MIDDLE, name=f"A label {k + 1}")
    vals = [f"${cl:,.0f}" if cl else "n.a.", f"${rv:,.0f}" if rv is not None else "n.a.", pct, rng]
    for j, (x, w, _) in enumerate(COLS):
        o = {"b": j == 2, "c": RED if (j == 2 and (hot or pct == "0%")) else (GREY_TXT if j == 3 else INK)}
        textbox(slide, x, yc[k] - 0.1, w, 0.2, [[(vals[j], o)]], size=10 if j == 3 else 10.5, align=PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"A r{k + 1} c{j + 1}")
    if k < N - 1 and rows[k + 1] is not None:
        line(slide, L0, yc[k] + slot / 2, LP1, yc[k] + slot / 2, HAIR, 0.5, name=f"A rule {k + 1}")
# value breaks here
yb = yc[6]
line(slide, L0, yb, LP1, yb, RED, 1.25, dash=MSO_LINE_DASH_STYLE.DASH, name="A break line")
bl = textbox(slide, 1.35, yb - 0.11, 5.6, 0.22,
             [[("Value breaks here: ", {"b": True, "c": RED}), ("$947M", {"b": True}), (" left for ", {}),
               ("$3,521M", {"b": True}), (" of unsecured claims", {})]], size=10.5, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, name="A break label")
bl.fill.solid()
bl.fill.fore_color.rgb = rgb(WHITE)

callout(L0, 5.72, LP1 - L0, 0.56,
        [("The preferred and common are ", {}), ("wiped out at every value", {"b": True, "c": RED}),
         (" in the range: even at $3,664M, the Holdings notes recover only ", {}), ("48c", {"b": True})],
        "A callout", size=10.5)

# ================================================================= RIGHT TOP: recovery vs market
textbox(slide, RP0, ROW1, R1 - RP0, 0.24, [[("Plan-Value Recovery vs. Market Price", {"b": True})]], size=14,
        name="B header")
line(slide, RP0, ROW1 + 0.29, R1, ROW1 + 0.29, INK, 1.0, name="B header rule")
tr = [("1L Term Loan B", 1.0, 1.0, 1.0, 0.774, "100%"),
      ("1L notes", 1.0, 1.0, 1.0, 0.876, "100%"),
      ("2L notes", 1.0, 1.0, 1.0, 0.677, "100%"),
      ("Holdings notes", 0.243, 0.360, 0.477, 0.463, "24% to 48%"),
      ("Legacy notes", 0.122, 0.180, 0.239, 0.320, "12% to 24%")]
D = dict(x=RP0 + 2.08, y=1.95, w=R1 - RP0 - 2.16, h=0.30 * len(tr))
dslot = D["h"] / len(tr)
dy = [D["y"] + dslot * (k + 0.5) for k in range(len(tr))]
dx = lambda v: D["x"] + D["w"] * v
dcats = [t[0] for t in tr]
dmid = [t[2] for t in tr]

# legend
box(slide, RP0, 1.70, 0.14, 0.12, fill=INK, line=None, shape=MSO_SHAPE.RECTANGLE, name="B legend bar")
textbox(slide, RP0 + 0.19, 1.65, 1.3, 0.2, [[("Recovery at mid EV", {"c": GREY_TXT})]], size=9.5, name="B legend t1")
line(slide, RP0 + 1.55, 1.76, RP0 + 1.80, 1.76, INK, 1.0, name="B legend whisker")
textbox(slide, RP0 + 1.85, 1.65, 0.9, 0.2, [[("Low to high", {"c": GREY_TXT})]], size=9.5, name="B legend t2")
lg = box(slide, RP0 + 2.78, 1.69, 0.14, 0.14, fill=WHITE, line=INK, line_w=1.0, shape=MSO_SHAPE.DIAMOND,
         name="B legend diamond")
textbox(slide, RP0 + 2.97, 1.65, 1.5, 0.2, [[("Market price", {"c": GREY_TXT}), ("4", {"sup": True, "c": GREY_TXT})]],
        size=9.5, name="B legend t3")


def mk_xml(rid):
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "Recovery at mid EV", dcats, dmid, INK, NOLBL, dpt_fills={3: RED_BAR, 4: RED_BAR})
            + '<c:gapWidth val="80"/><c:axId val="3001"/><c:axId val="3002"/></c:barChart>' + hbar_axes(3001, 3002, 1))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, D["x"], D["y"], D["w"], D["h"], dcats, [("Recovery at mid EV", dmid)], mk_xml, "B recovery chart")
line(slide, D["x"], D["y"], D["x"], D["y"] + D["h"], GREY_MID, 0.75, name="B axis")
for k, (nm, lo, mid, hi, px, rng) in enumerate(tr):
    unsec = mid < 1
    textbox(slide, RP0, dy[k] - 0.1, 1.25, 0.2, [[(nm, {"b": unsec})]], size=10.5, anchor=MSO_ANCHOR.MIDDLE,
            name=f"B label {k + 1}")
    textbox(slide, RP0 + 1.12, dy[k] - 0.1, 0.88, 0.2, [[(rng, {"b": unsec, "c": RED if unsec else INK})]], size=10,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, name=f"B range {k + 1}")
    if unsec:
        line(slide, dx(lo), dy[k], dx(hi), dy[k], INK, 1.0, name=f"B whisker {k + 1}")
        for v in (lo, hi):
            line(slide, dx(v), dy[k] - 0.06, dx(v), dy[k] + 0.06, INK, 1.0, name=f"B cap {k + 1}")
    d = box(slide, dx(px) - 0.07, dy[k] - 0.07, 0.14, 0.14, fill=WHITE, line=INK if unsec else WHITE, line_w=1.0,
            shape=MSO_SHAPE.DIAMOND, name=f"B market {k + 1}")
    lx = max(dx(px), dx(hi)) + 0.10 if unsec else dx(px) - 0.55
    textbox(slide, lx, dy[k] - 0.09, 0.45, 0.18, [[(f"{px * 100:.0f}c", {"b": True, "c": INK if unsec else WHITE})]],
            size=9.5, align=PP_ALIGN.LEFT if unsec else PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
            name=f"B market label {k + 1}")
textbox(slide, RP0, D["y"] + D["h"] + 0.06, R1 - RP0, 0.2,
        [[("Secured trades below full recovery; legacy notes trade ", {"i": True, "c": GREY_TXT}),
          ("above", {"i": True, "b": True, "c": RED}), (" their range", {"i": True, "c": GREY_TXT})]], size=10,
        name="B note")

# ================================================================= RIGHT BOTTOM: who holds the fulcrum value
section_header(slide, RP0, 3.92, R1 - RP0, "Who Holds the Fulcrum Value (Mid, $M)", "C header")
HOLD, PENS, LEG = 625.5, 211.6, 109.7
UV = HOLD + PENS + LEG
textbox(slide, RP0, 4.30, R1 - RP0, 0.2, [[("Unsecured value at mid plan value: ", {"c": GREY_TXT}),
                                         (f"${UV:,.0f}", {"b": True})]], size=10, name="C caption")
E = dict(x=RP0, y=4.50, w=R1 - RP0, h=0.40)
ecats = ["Unsecured value"]


def one_ser(i, name, v, fill):
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:f>Sheet1!${col(i)}$1</c:f>'
            f'{str_cache([name])}</c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>{NOLBL}'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$2</c:f>{str_cache(ecats)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!${col(i)}$2:${col(i)}$2</c:f>{num_cache([v])}</c:numRef></c:val></c:ser>')


def split_xml(rid):
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>'
            + one_ser(0, "Holdings notes", HOLD, INK) + one_ser(1, "Pension", PENS, GREY_MID)
            + one_ser(2, "Legacy notes", LEG, GREY_LT)
            + '<c:gapWidth val="0"/><c:overlap val="100"/><c:axId val="3101"/><c:axId val="3102"/></c:barChart>'
            + hbar_axes(3101, 3102, 1))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, E["x"], E["y"], E["w"], E["h"], ecats, [("Holdings notes", [HOLD]), ("Pension", [PENS]),
                                                         ("Legacy notes", [LEG])], split_xml, "C split chart")
ex = lambda v: E["x"] + E["w"] * v / UV
segs = [(0, HOLD, f"Holdings notes ${HOLD:,.0f} ({HOLD / UV:.0%})", WHITE),
        (HOLD, PENS, f"${PENS:,.0f} ({PENS / UV:.0%})", WHITE),
        (HOLD + PENS, LEG, f"${LEG:,.0f}", INK)]
for a, v, t, c in segs:
    tb = textbox(slide, ex(a), E["y"], ex(a + v) - ex(a), E["h"], [[(t, {"b": True, "c": c})]], size=10,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name=f"C seg {t[:8]}")
    tb.text_frame.word_wrap = False
textbox(slide, RP0, E["y"] + E["h"] + 0.05, 2.6, 0.2, [[("36c on $1,737 of claims", {"c": GREY_TXT})]], size=9.5,
        name="C note holdings")
textbox(slide, R1 - 2.0, E["y"] + E["h"] + 0.05, 2.0, 0.2, [[("Pension and legacy: 18c", {"c": GREY_TXT})]],
        size=9.5, align=PP_ALIGN.RIGHT, name="C note others")
callout(RP0, 5.72, R1 - RP0, 0.56,
        [("Holdings notes recover ", {}), ("2x", {"b": True}), (" the legacy notes: they claim twice against Xerox "
                                                                 "Corp. (guarantee and intercompany loan)", {})],
        "C callout", size=10.5)

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Waterfall conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("The Holdings notes are the fulcrum and hold ", {}), ("two-thirds of unsecured value", {"b": True, "c": RED}),
    (": they should ", {}), ("lead the restructuring and own the reorganized Xerox", {"i": True})]],
    11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 12 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(11, new)
prs.save(OUT)
print("saved", OUT, "UV", round(UV, 1))
