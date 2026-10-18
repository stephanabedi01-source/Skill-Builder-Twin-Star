"""Add the Public Comps slide (slide 8, ahead of the Recommendation) to Project Print deck v10.

Usage: python3 build_comps.py <input.pptx> <output.pptx>
Numbers: FactSet comps as of 10/14/26 ($XRX Comparable Companies), Capital Structure, Project Print model.
Logos are copied from the Company Snapshot slide (slide 3); Pitney Bowes has no logo there yet.
"""
import sys

_argv = list(sys.argv)
sys.argv = sys.argv[:1]
from pp_helpers import *  # noqa: E402,F401,F403
from pp_helpers import textbox, section_header  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402

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
slide.placeholders[21].text_frame.text = "Public Comps"
slide.placeholders[16].text_frame.text = (
    "Even at peer multiples of 5.9x to 6.4x, Xerox is worth $0.4B to $0.7B less than its $4.4B of debt")

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
         "Notes: (1) FactSet EV and LTM EBITDA; Quadient excluded (no LTM EBITDA); (2) Xerox: $614M LTM Adj. "
         "EBITDA excl. $105mm IEEPA benefit on $7.8B revenue; (3) Bonds at 10/14/26 prices ($2,686M, ~62% of "
         "face); (4) Debt at face ($4,355M); (5) Xerox net debt at face")
ftr_para(tf.add_paragraph(),
         "Sources: FactSet as of 10/14/26; Xerox Q2'26 10-Q; peer filings; Project Print model")

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


def pill(x, y, w, h, text, fill, line_c, color, name, size=10.5):
    p = box(slide, x, y, w, h, fill=fill, line=line_c, line_w=0.75, radius=0.3, name=name)
    p.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p.text_frame.margin_left = p.text_frame.margin_right = Inches(0.02)
    p.text_frame.margin_top = p.text_frame.margin_bottom = 0
    fill_tf(p.text_frame, [[(text, {"b": True, "c": color})]], size, align=PP_ALIGN.CENTER)
    return p


# ================================================================= LEFT: multiples with logos (mirrors slide 3)
section_header(slide, L0, ROW1, LP1 - L0, "Trading Multiples vs. Print Peers", "A header")
# (logo on slide 3, EV / LTM EBITDA, EBITDA margin)
peers = [
    ("Pitney Bowes", None, 7.64, 0.296),
    ("Canon", "Picture 90", 6.77, 0.151),
    ("HP", "Picture 87", 6.59, 0.085),
    ("Ricoh", "Picture 89", 6.41, 0.079),
    ("Epson", "Picture 94", 6.2514, 0.112),
    ("Brother", "Picture 96", 5.59, 0.182),
    ("Konica Minolta", "Graphic 95", 5.44, 0.106),
]
XRX_MKT, XRX_FACE, XRX_MGN = 4.39, 7.11, 0.079
Q1, MED, MED_MGN = 5.92, 6.41, 0.112
rows = peers + [("Xerox", "Picture 93", XRX_MKT, XRX_MGN)]
N = len(rows)
cats = [r[0] for r in rows]
vals = [r[2] for r in rows]
gap = [None] * (N - 1) + [round(XRX_FACE - XRX_MKT, 2)]

C = dict(x=1.60, y=1.92, w=5.30, h=0.44 * N, vmin=0.0, vmax=8.0)
slot = C["h"] / N
yc = [C["y"] + slot * (k + 0.5) for k in range(N)]
xv = lambda v: C["x"] + C["w"] * (v - C["vmin"]) / (C["vmax"] - C["vmin"])
MX = 7.62               # margin column left edge
MW = 0.62

textbox(slide, C["x"], 1.66, 2.4, 0.2, [[("EV / LTM Adj. EBITDA", {"i": True, "c": GREY_TXT}),
                                         ("1", {"sup": True, "i": True, "c": GREY_TXT})]], size=10.5,
        name="A caption multiple")
textbox(slide, MX - 0.3, 1.66, MW + 0.6, 0.2, [[("EBITDA margin", {"i": True, "c": GREY_TXT}),
                                                ("2", {"sup": True, "i": True, "c": GREY_TXT})]], size=10.5,
        align=PP_ALIGN.CENTER, name="A caption margin")

# peer Q1 to median band behind the bars
box(slide, xv(Q1), C["y"] - 0.04, xv(MED) - xv(Q1), C["h"] + 0.04, fill=BAND, line=None, shape=MSO_SHAPE.RECTANGLE,
    name="A peer band")
textbox(slide, xv(Q1) - 1.2, 1.66, xv(MED) - xv(Q1) + 2.4, 0.2,
        [[("Peer Q1 to median: ", {"b": True, "size": 9.5}), ("5.9x to 6.4x", {"size": 9.5})]], size=9.5,
        align=PP_ALIGN.CENTER, name="A peer band label")

# Xerox highlight row (same treatment as slide 3)
callout(L0, yc[-1] - 0.21, LP1 - L0, 0.42, [("", {})], "A xerox row")


def mult_xml(rid):
    gap_ser = (f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f>'
               f'{str_cache(["Xerox at face"])}</c:strRef></c:tx>'
               f'<c:spPr><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill>'
               f'<a:prstDash val="dash"/></a:ln></c:spPr><c:invertIfNegative val="0"/>'
               f'{dlbls("", None, 9, INK, show=False)}'
               f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${N + 1}</c:f>{str_cache(cats)}</c:strRef></c:cat>'
               f'<c:val><c:numRef><c:f>Sheet1!$C$2:$C${N + 1}</c:f>{num_cache(gap)}</c:numRef></c:val></c:ser>')
    body = ('<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
            + bar_ser(0, "EV / LTM EBITDA", cats, vals, GREY_LT, dlbls("", None, 9, INK, show=False),
                      dpt_fills={N - 1: RED_BAR})
            + gap_ser
            + '<c:gapWidth val="70"/><c:overlap val="100"/><c:axId val="2201"/><c:axId val="2202"/></c:barChart>'
            + f'<c:catAx><c:axId val="2201"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
              f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
              f'<c:crossAx val="2202"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
              f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            + f'<c:valAx><c:axId val="2202"/><c:scaling><c:orientation val="minMax"/><c:max val="{C["vmax"]}"/>'
              f'<c:min val="{C["vmin"]}"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
              f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
              f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2201"/>'
              f'<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, C["x"], C["y"], C["w"], C["h"], cats, [("EV / LTM EBITDA", vals), ("Xerox at face", gap)],
          mult_xml, "A multiples chart")
line(slide, C["x"], C["y"] - 0.04, C["x"], C["y"] + C["h"] + 0.04, GREY_MID, 0.75, name="A axis")

for k, (name, logo, m, mg) in enumerate(rows):
    xerox = k == N - 1
    if xerox:
        textbox(slide, xv(m) + 0.06, yc[k] - 0.1, 1.5, 0.2,
                [[(f"{m:.1f}x", {"b": True, "c": RED}), (" at market", {"c": INK}), ("3", {"sup": True})]],
                size=10.5, name="A xerox market label")
        textbox(slide, xv(XRX_FACE) + 0.06, yc[k] - 0.1, 1.2, 0.2,
                [[(f"{XRX_FACE:.1f}x", {"b": True}), (" at face", {}), ("4", {"sup": True})]], size=10.5,
                name="A xerox face label")
        pill(MX, yc[k] - 0.12, MW, 0.24, f"{mg:.1%}", RED, RED, WHITE, "A xerox margin")
    else:
        textbox(slide, xv(m) + 0.06, yc[k] - 0.1, 0.6, 0.2, [[(f"{m:.1f}x", {"b": True})]], size=10.5,
                name=f"A value {k + 1}")
        textbox(slide, MX, yc[k] - 0.1, MW, 0.2, [[(f"{mg:.1%}", {"b": True})]], size=10.5,
                align=PP_ALIGN.CENTER, name=f"A margin {k + 1}")
    if logo is None:    # no logo on slide 3: drop the Pitney Bowes logo over this label
        textbox(slide, L0, yc[k] - 0.11, 1.25, 0.22, [[(name, {"b": True})]], size=11, align=PP_ALIGN.CENTER,
                name=f"Logo placeholder {name}")

# peer median row
ym = C["y"] + C["h"] + 0.06
line(slide, L0, ym, LP1, ym, HAIR, 0.5, name="A median rule")
textbox(slide, L0, ym + 0.05, 4.0, 0.22, [[("Peer median (7 peers): ", {"i": True, "c": GREY_TXT}),
                                           ("6.4x", {"b": True, "i": True, "c": GREY_TXT})]], size=10.5,
        name="A median label")
pill(MX, ym + 0.05, MW, 0.23, f"{MED_MGN:.1%}", WHITE, INK, INK, "A median margin")

callout(L0, 5.88, LP1 - L0, 0.44,
        [("At face value Xerox screens at ", {}), ("7.1x", {"b": True}), (", above peers; with bonds at ", {}),
         ("~62% of par", {"b": True}), (", the market values it at ", {}), ("4.4x, below every peer", {"b": True, "c": RED})],
        "A callout", size=10.5)

# ================================================================= RIGHT TOP: leverage
section_header(slide, RP0, ROW1, R1 - RP0, "Leverage Explains the Discount", "B header")
textbox(slide, RP0, 1.66, 3.0, 0.2, [[("Net debt / LTM EBITDA", {"i": True, "c": GREY_TXT}),
                                      ("5", {"sup": True, "i": True, "c": GREY_TXT})]], size=10.5,
        name="B caption")
lev = [("Brother", -1.04), ("Epson", -0.35), ("Canon", 0.46), ("HP", 1.42), ("Ricoh", 1.49), ("Konica", 2.48),
       ("PBI", 3.37), ("Xerox", 6.29)]
LEV_MED = 1.42
lc = [a for a, _ in lev]
lvv = [b for _, b in lev]
D = dict(x=RP0 + 0.05, y=1.98, w=R1 - RP0 - 0.05, h=1.42, vmin=-1.9, vmax=7.4)
dslot = D["w"] / len(lev)
dx = [D["x"] + dslot * (k + 0.5) for k in range(len(lev))]
yv = lambda v: D["y"] + D["h"] * (D["vmax"] - v) / (D["vmax"] - D["vmin"])


def lev_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "Net debt / LTM EBITDA", lc, lvv, GREY_LT, dlbls("", None, 9, INK, show=False),
                      dpt_fills={len(lev) - 1: RED_BAR})
            + '<c:gapWidth val="55"/><c:axId val="2301"/><c:axId val="2302"/></c:barChart>'
            + cat_ax(2301, 2302, delete=True) + val_ax(2302, 2301, D["vmin"], D["vmax"]))
    return chart_space((0, 0, 1, 1), body, rid)


add_chart(slide, D["x"], D["y"], D["w"], D["h"], lc, [("Net debt / LTM EBITDA", lvv)], lev_xml, "B leverage chart")
line(slide, D["x"], yv(0), D["x"] + D["w"], yv(0), GREY_MID, 0.75, name="B zero line")
line(slide, D["x"], yv(LEV_MED), D["x"] + dslot * 7, yv(LEV_MED), GREY_MID, 1.0, dash=MSO_LINE_DASH_STYLE.DASH,
     name="B median line")
textbox(slide, D["x"], yv(LEV_MED) - 0.21, 1.6, 0.18, [[("Peer median 1.4x", {"i": True, "c": GREY_TXT})]],
        size=9.5, name="B median label")
for k, (nm, v) in enumerate(lev):
    xerox = k == len(lev) - 1
    txt = f"{v:.1f}x" if v >= 0 else f"({-v:.1f}x)"
    ty = yv(v) - 0.2 if v >= 0 else yv(v) + 0.02
    vb = textbox(slide, dx[k] - 0.2, ty, 0.4, 0.18, [[(txt, {"b": True, "c": RED if xerox else INK})]], size=10,
                 align=PP_ALIGN.CENTER, name=f"B value {k + 1}")
    if 0 <= v < LEV_MED:    # label sits on the median line: white fill breaks the line around it
        vb.fill.solid()
        vb.fill.fore_color.rgb = rgb(WHITE)
    textbox(slide, dx[k] - 0.3, D["y"] + D["h"] + 0.06, 0.6, 0.18,
            [[(nm, {"b": xerox, "c": RED if xerox else GREY_TXT})]], size=9.5, align=PP_ALIGN.CENTER,
            name=f"B category {k + 1}")
textbox(slide, dx[-1] - dslot * 2.6, yv(5.6) - 0.1, dslot * 2.2, 0.2,
        [[("~4x the peer median", {"b": True, "c": RED})]], size=10, align=PP_ALIGN.RIGHT, name="B xerox note")

# ================================================================= RIGHT BOTTOM: shortfall at peer multiples
section_header(slide, RP0, 3.80, R1 - RP0, "Peer Multiples Still Leave a Shortfall", "C header")
CW = [2.25, 1.12, 1.12]
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
            o = {"b": k in bold_rows, "c": RED if (k in red_rows and j > 0) else INK}
            textbox(slide, x, y, CW[j], 0.2, [[(cell, o)]], size=10.5, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT,
                    anchor=MSO_ANCHOR.MIDDLE, name=f"{name} r{k + 1} c{j + 1}")
            x += CW[j]
        y += 0.26
    return y


yt = table(4.20, ["($M)", "Peer Q1", "Peer median"],
           [["EV / LTM Adj. EBITDA", "5.9x", "6.4x"],
            ["Implied enterprise value", "$3,635", "$3,937"],
            ["(-) Total debt", "(4,355)", "(4,355)"],
            ["Shortfall to total debt", "($720)", "($418)"]],
           "C table", bold_rows=(3,), red_rows=(3,), rule_before=(3,))
callout(RP0, yt + 0.08, R1 - RP0, 0.66,
        [("The selected ", {}), ("5.0x to 6.0x", {"b": True}), (" sits below the peer band: Xerox carries ", {}),
         ("6.3x", {"b": True, "c": RED}), (" net leverage vs. 1.4x for peers, a bottom-of-group margin and print "
                                           "revenue down 4% to 10% a year", {})],
        "C callout", size=10)

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Comps conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Peer multiples are a ceiling, not a target: even the peer median leaves Xerox ", {}),
    ("$0.4B short", {"b": True, "c": RED}), (" of its debt, and ", {}),
    ("its bonds already price it at 4.4x", {"i": True})]], 11.5, align=PP_ALIGN.CENTER)

# ================================================================= logos copied from slide 3
next_id = max(int(e.get("id")) for e in spTree.iter(qn("p:cNvPr"))) + 1
pics = {e.find(".//" + qn("p:cNvPr")).get("name"): e for e in ref.shapes._spTree.iter(qn("p:pic"))}
LOGO_CX = L0 + 0.62
for k, (name, logo, _, _) in enumerate(rows):
    if logo is None:
        continue
    el = copy.deepcopy(pics[logo])
    for node in el.iter():
        rid = node.get(qn("r:embed"))
        if rid:
            node.set(qn("r:embed"), slide.part.relate_to(ref.part.related_part(rid), RT.IMAGE))
    nv = el.find(".//" + qn("p:cNvPr"))
    nv.set("id", str(next_id))
    nv.set("name", f"Logo {name}")
    next_id += 1
    for ext in nv.findall(qn("a:extLst")):
        nv.remove(ext)
    xf = el.find(".//" + qn("a:xfrm"))
    ext = xf.find(qn("a:ext"))
    w, h = int(ext.get("cx")), int(ext.get("cy"))
    xf.find(qn("a:off")).set("x", str(int(Inches(LOGO_CX)) - w // 2))
    xf.find(qn("a:off")).set("y", str(int(Inches(yc[k])) - h // 2))
    spTree.append(el)

# ---------------------------------------------------------------- move to slide 8 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(7, new)
prs.save(OUT)
print("saved", OUT)
