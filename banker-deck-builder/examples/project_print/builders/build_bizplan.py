"""Add the Business Plan slide (slide 6, ahead of the Recommendation) to Project Print deck v8.

Usage: python3 build_bizplan.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v15 ($XRX Normalized EBITDA, $XRX DCF Operating Assumptions, $XRX Liquidity Analysis).
"""
import sys

_argv = list(sys.argv)
sys.argv = sys.argv[:1]
from pp_helpers import *  # noqa: E402,F401,F403
from pp_helpers import textbox, section_header  # noqa: E402

SRC, OUT = _argv[1], _argv[2]
HAIR = "D1D1D1"
GREY_FILL = "D6D6D6"

prs = Presentation(SRC)
ref = prs.slides[2]
slide = prs.slides.add_slide(ref.slide_layout)
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 20:
        ph._element.getparent().remove(ph._element)
slide.placeholders[21].text_frame.text = "Business Plan"
slide.placeholders[16].text_frame.text = (
    "Synergies lift Adj. EBITDA to ~$730M by 2027, but cash flow covers only interest, not $2.9B of principal "
    "due by 2030")

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
         "Notes: (1) Excl. $105mm IEEPA benefit; (2) Est. $115mm of synergies and 40% of Transformation savings in "
         "LTM; (3) 90% of remaining savings realized; (4) Pro forma; (5) Adj. EBITDA less depreciation; "
         "(6) FCF before interest")
ftr_para(tf.add_paragraph(),
         "Sources: Xerox Q2'26 earnings slides, FY2025 10-K, Q2'26 10-Q; FactSet consensus as of 10/14/26; "
         "S&P (6/16/26); Project Print model")

L0, L1, R0, R1 = 0.23, 6.53, 6.81, 13.11
ROW1, ROW2 = 1.30, 4.13


def callout(x, y, w, h, runs, name, size=10.5):
    c = box(slide, x, y, w, h, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
            radius=0.08, name=name)
    c.text_frame.margin_left = c.text_frame.margin_right = Inches(0.08)
    c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(c.text_frame, [runs], size, align=PP_ALIGN.CENTER)
    return c


def geom(C, n):
    x0 = C["x"] + C["px"] * C["w"]
    w = C["pw"] * C["w"]
    y0 = C["y"] + C["py"] * C["h"]
    h = C["ph"] * C["h"]
    slot = w / n
    centers = [x0 + slot * (k + 0.5) for k in range(n)]
    yv = lambda v: y0 + h - h * (v - C["vmin"]) / (C["vmax"] - C["vmin"])
    return centers, yv, slot


def line_ser_d(i, name, cats, vals, color, labels, marker="none", msize=6, mfill=WHITE, dash=None, width=2.0,
               show_line=True):
    n = len(cats)
    d = f'<a:prstDash val="{dash}"/>' if dash else ""
    ln = (f'<a:ln w="{int(width * 12700)}" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{d}</a:ln>'
          if show_line else '<a:ln><a:noFill/></a:ln>')
    mk = (f'<c:marker><c:symbol val="none"/></c:marker>' if marker == "none" else
          f'<c:marker><c:symbol val="{marker}"/><c:size val="{msize}"/><c:spPr>'
          + (f'<a:solidFill><a:srgbClr val="{mfill}"/></a:solidFill>' if mfill else '<a:noFill/>')
          + f'<a:ln w="15875"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr></c:marker>')
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col(i)}$1</c:f>{str_cache([name])}</c:strRef></c:tx>'
            f'<c:spPr>{ln}</c:spPr>{mk}{labels}'
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n + 1}</c:f>{str_cache(cats)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>Sheet1!${col(i)}$2:${col(i)}${n + 1}</c:f>{num_cache(vals)}</c:numRef></c:val>'
            f'<c:smooth val="0"/></c:ser>')


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

# ================================================================= MODULE 1: EBITDA bridge
section_header(slide, L0, ROW1, L1 - L0, "Path to Normalized Adj. EBITDA (LTM to 2027E)", "M1 header")
w_cats = ["LTM 2Q'26", "Remaining synergies", "Transformation", "Execution haircut", "Revenue decline",
          "XFS run-off, other", "2027E base case"]
w_lbl = [["LTM", "2Q'26"], ["Remaining", "synergies"], ["Transform-", "ation"], ["Execution", "haircut"],
         ["Revenue", "decline"], ["XFS run-off,", "other"], ["2027E", "base case"]]
w_base = [None, 614, 849, 879.5, 732.1, 727.6, None]
w_up = [None, 235, 60, None, None, None, None]
w_dn = [None, None, None, 29.5, 147.4, 4.5, None]
w_tot = [614, None, None, None, None, None, 727.6]
C1 = dict(x=L0, y=1.62, w=5.02, h=2.10, px=0.01, pw=0.98, py=0.10, ph=0.86, vmin=0, vmax=1000)


def m1_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
            + nofill_ser(0, "Base", w_cats, w_base)
            + bar_ser(1, "Increase", w_cats, w_up, GREY_MID, NOLBL)
            + bar_ser(2, "Decrease", w_cats, w_dn, RED_BAR, NOLBL)
            + bar_ser(3, "Adj. EBITDA", w_cats, w_tot, INK, NOLBL)
            + '<c:gapWidth val="45"/><c:overlap val="100"/><c:axId val="1101"/><c:axId val="1102"/></c:barChart>'
            + cat_ax(1101, 1102, lbl_pos="none") + val_ax(1102, 1101, C1["vmin"], C1["vmax"]))
    return chart_space((C1["px"], C1["py"], C1["pw"], C1["ph"]), body, rid)


add_chart(slide, C1["x"], C1["y"], C1["w"], C1["h"], w_cats,
          [("Base", w_base), ("Increase", w_up), ("Decrease", w_dn), ("Adj. EBITDA", w_tot)], m1_xml,
          "M1 EBITDA bridge chart")
c1, y1, s1 = geom(C1, 7)
lbls = [("$614M", 614, INK, "1"), ("+$235M", 849, INK, "2"), ("+$60M", 909, INK, None), ("($30M)", 909, RED, "3"),
        ("($147M)", 879.5, RED, None), ("($5M)", 732.1, RED, None), ("$728M", 727.6, INK, None)]
for k, (t, top, c, sup) in enumerate(lbls):
    runs = [(t, {"b": True, "c": c})] + ([(sup, {"c": c, "sup": True})] if sup else [])
    textbox(slide, c1[k] - 0.45, y1(top) - 0.23, 0.9, 0.2, [runs], size=10.5, align=PP_ALIGN.CENTER,
            name=f"M1 label {k + 1}")
for k, (a, b) in enumerate(w_lbl):
    textbox(slide, c1[k] - 0.4, y1(0) + 0.05, 0.8, 0.34, [[(a, {"c": GREY_TXT})], [(b, {"c": GREY_TXT})]], size=9,
            align=PP_ALIGN.CENTER, name=f"M1 category {k + 1}")
# growth pill over the end bar
gp = box(slide, c1[6] - 0.42, y1(727.6) - 0.52, 0.84, 0.26, fill=WHITE, line=INK,
         dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0, radius=0.5, name="M1 growth pill")
gp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(gp.text_frame, [[("+18.5%", {"b": True})]], 10, align=PP_ALIGN.CENTER)
callout(5.36, 2.05, 1.17, 1.30,
        [("$295M", {"b": True}), (" of remaining savings outweighs ", {}), ("$152M", {"b": True, "c": RED}),
         (" of revenue and XFS drag", {})], "M1 callout", size=10)

# ================================================================= MODULE 2: cases vs consensus
section_header(slide, R0, ROW1, R1 - R0, "Base Case vs. Management Plan and Consensus", "M2 header")
e_cats = ["FY23", "FY24", "FY25", "LTM", "26E", "27E", "28E", "29E", "30E"]
e_hist = [651, 555, 512, 614, None, None, None, None, None]
e_base = [None, None, None, 614, 712, 728, 719, 694, 670]
e_bull = [None, None, None, 614, 761, 791, 807, 788, 770]
e_bear = [None, None, None, 614, 646, 609, 595, 567, 541]
e_cons = [None] * 5 + [781] + [None] * 3
e_sp = [None] * 5 + [511] + [None] * 3
C2 = dict(x=R0, y=1.66, w=4.62, h=2.42, px=0.02, pw=0.96, py=0.10, ph=0.74, vmin=380, vmax=880)
CUR = '\\$#,##0"M"'


def m2_xml(rid):
    body = ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            + line_ser_d(0, "Actual", e_cats, e_hist, INK,
                         dlbls(CUR, "t", 9.5, INK, False), marker="circle", msize=5, mfill=INK, width=2.0)
            + line_ser_d(1, "Mgmt plan", e_cats, e_bull, GREY_MID,
                         dlbls(CUR, "t", 9.5, GREY_MID, True, deleted_idx=range(8)), dash="dash", width=1.5)
            + line_ser_d(2, "2025-style miss", e_cats, e_bear, RED,
                         dlbls(CUR, "b", 9.5, RED, True, deleted_idx=range(8)), dash="sysDot", width=1.75)
            + line_ser_d(3, "Base case", e_cats, e_base, INK,
                         dlbls(CUR, "b", 10, INK, True, deleted_idx=[0, 1, 2, 3, 4, 6, 7]), marker="diamond",
                         msize=6, mfill=WHITE, width=2.25)
            + line_ser_d(4, "FactSet consensus", e_cats, e_cons, GREY_MID, NOLBL, marker="circle", msize=8,
                         mfill=WHITE, show_line=False)
            + line_ser_d(5, "S&amp;P default case", e_cats, e_sp, RED, NOLBL, marker="square", msize=7,
                         mfill=WHITE, show_line=False)
            + '<c:marker val="1"/><c:axId val="1201"/><c:axId val="1202"/></c:lineChart>'
            + cat_ax(1201, 1202) + val_ax(1202, 1201, C2["vmin"], C2["vmax"]))
    return chart_space((C2["px"], C2["py"], C2["pw"], C2["ph"]), body, rid)


add_chart(slide, C2["x"], C2["y"], C2["w"], C2["h"], e_cats,
          [("Actual", e_hist), ("Mgmt plan", e_bull), ("2025-style miss", e_bear), ("Base case", e_base),
           ("FactSet consensus", e_cons), ("S&P default case", e_sp)], m2_xml, "M2 EBITDA cases chart")
c2, y2, s2 = geom(C2, 9)
# projection shading divider and marker labels
line(slide, (c2[3] + c2[4]) / 2, C2["y"] + 0.12, (c2[3] + c2[4]) / 2, y2(C2["vmin"]), GREY_LT, 0.75,
     dash=MSO_LINE_DASH_STYLE.DASH, name="M2 projection divider")
textbox(slide, (c2[3] + c2[4]) / 2 + 0.05, C2["y"] + 0.06, 1.0, 0.17, [[("Projected", {"i": True, "c": GREY_TXT})]],
        size=9, name="M2 projected label")
# legend strip
LX, LY = 11.55, 1.78
leg = [("Base case", INK, None, "diamond"), ("Mgmt plan", GREY_MID, MSO_LINE_DASH_STYLE.DASH, None),
       ("2025-style miss", RED, MSO_LINE_DASH_STYLE.ROUND_DOT, None)]
for k, (t, c, d, mk) in enumerate(leg):
    yy = LY + k * 0.24
    line(slide, LX, yy + 0.09, LX + 0.24, yy + 0.09, c, 1.75, dash=d, name=f"M2 legend line {k + 1}")
    if mk:
        box(slide, LX + 0.075, yy + 0.04, 0.09, 0.1, fill=WHITE, line=c, line_w=0.75, shape=MSO_SHAPE.DIAMOND,
            name=f"M2 legend marker {k + 1}")
    textbox(slide, LX + 0.32, yy, 1.3, 0.18, [[(t, {"c": GREY_TXT})]], size=9.5, name=f"M2 legend {k + 1}")
# consensus and S&P points drawn as shapes on the chart (single-point series do not render reliably)
for v, c, shp, nm in ((781, GREY_MID, MSO_SHAPE.OVAL, "consensus"), (511, RED, MSO_SHAPE.RECTANGLE, "S&P")):
    box(slide, c2[5] - 0.06, y2(v) - 0.06, 0.12, 0.12, fill=WHITE, line=c, line_w=1.25, shape=shp,
        name=f"M2 {nm} marker")
for k, (t, c, shp) in enumerate([("Consensus $781M", GREY_MID, MSO_SHAPE.OVAL),
                                 ("S&P default $511M", RED, MSO_SHAPE.RECTANGLE)]):
    yy = LY + (3 + k) * 0.24
    box(slide, LX + 0.07, yy + 0.04, 0.1, 0.1, fill=WHITE, line=c, line_w=1.0, shape=shp,
        name=f"M2 legend marker {4 + k}")
    textbox(slide, LX + 0.32, yy, 1.28, 0.18, [[(t, {"c": c if c == RED else GREY_TXT})]], size=9.5,
            name=f"M2 legend {4 + k}")
callout(11.52, 3.06, 1.59, 0.98,
        [("Base sits ", {}), ("~7% below", {"b": True, "c": RED}), (" consensus for 2027E, haircutting a "
                                                                      "management that missed 2025 guidance", {})],
        "M2 callout", size=10)

# ================================================================= MODULE 3: assumptions vs track record
section_header(slide, L0, ROW2, L1 - L0, "Base Case Assumptions vs. Track Record", "M3 header")
TX, TY = L0, 4.52
CW = [1.95, 0.87, 0.87, 0.87]
hdrs = ["Driver", "FY25A", "Base 27E", "Mgmt 27E"]
rows = [
    ("Post-sale revenue growth", "4", "(9.8%)", "(4.5%)", "(3.5%)"),
    ("Equipment revenue growth", "4", "(8.5%)", "(6.0%)", "(5.0%)"),
    ("Adj. operating margin", "5", "3.8%", "6.3%", "7.0%"),
    ("Remaining savings realized", None, "n.a.", "90%", "100%"),
    ("Adj. EBITDA", "1", "$512M", "$728M", "$791M"),
]
RH = 0.27
x = TX
for j, h in enumerate(hdrs):
    textbox(slide, x + (0 if j == 0 else 0), TY, CW[j], 0.2, [[(h, {"b": True})]], size=10.5,
            align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, name=f"M3 header {j + 1}")
    x += CW[j]
TW = sum(CW)
line(slide, TX, TY + 0.25, TX + TW, TY + 0.25, INK, 0.75, name="M3 header rule")
y = TY + 0.31
for k, (drv, sup, a, b, m) in enumerate(rows):
    x = TX
    runs0 = [(drv, {})] + ([(sup, {"sup": True})] if sup else [])
    cells = [runs0, [(a, {"c": RED if a.startswith("(") else INK})], [(b, {"b": True})], [(m, {"c": GREY_TXT})]]
    for j, runs in enumerate(cells):
        textbox(slide, x, y, CW[j], 0.2, [runs], size=10.5, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"M3 r{k + 1} c{j + 1}")
        x += CW[j]
    y += RH
    if k < len(rows) - 1:
        line(slide, TX, y - 0.04, TX + TW, y - 0.04, HAIR, 0.5, name=f"M3 rule {k + 1}")
# highlight the base-case column
bc_x = TX + CW[0] + CW[1] + 0.06
hl = box(slide, bc_x, TY - 0.03, CW[2] - 0.02, y - TY - 0.02, fill=None, line=INK, line_w=0.75, radius=0.06,
         name="M3 base column outline")
callout(4.95, 4.62, 1.58, 1.28,
        [("Management cut 2025 margin guidance ", {}), ("twice", {"b": True}), (", from at least 5.0% to ", {}),
         ("3.5%", {"b": True, "c": RED}), (", and delivered ", {}), ("$133M", {"b": True, "c": RED}),
         (" of FCF vs. ~$375M guided", {})], "M3 callout", size=10)

# ================================================================= MODULE 4: cash generation vs debt service
section_header(slide, R0, ROW2, R1 - R0, "Cash Generation vs. Debt Service, $M (Status Quo)", "M4 header")
d_cats = ["2027E", "2028E", "2029E", "2030E"]
d_cash = [418, 382, 338, 293]
d_serv = [452, 985, 1333, 1255]
C4 = dict(x=R0, y=4.50, w=4.62, h=1.88, px=0.02, pw=0.96, py=0.13, ph=0.70, vmin=0, vmax=1500)


def m4_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "Cash generation", d_cats, d_cash, INK, dlbls('\\$#,##0', "outEnd", 10.5, INK, True))
            + bar_ser(1, "Debt service", d_cats, d_serv, RED_BAR, dlbls('\\$#,##0', "outEnd", 10.5, RED, True))
            + '<c:gapWidth val="90"/><c:overlap val="-12"/><c:axId val="1401"/><c:axId val="1402"/></c:barChart>'
            + cat_ax(1401, 1402) + val_ax(1402, 1401, C4["vmin"], C4["vmax"]))
    return chart_space((C4["px"], C4["py"], C4["pw"], C4["ph"]), body, rid)


add_chart(slide, C4["x"], C4["y"], C4["w"], C4["h"], d_cats,
          [("Cash generation", d_cash), ("Debt service", d_serv)], m4_xml, "M4 cash vs debt service chart")
LX4, LY4 = 11.55, 4.58
box(slide, LX4, LY4 + 0.035, 0.11, 0.11, fill=INK, shape=MSO_SHAPE.RECTANGLE, name="M4 legend cash key")
textbox(slide, LX4 + 0.2, LY4, 1.4, 0.18, [[("Cash generation", {"c": GREY_TXT}), ("6", {"c": GREY_TXT, "sup": True})]],
        size=9.5, name="M4 legend cash")
box(slide, LX4, LY4 + 0.275, 0.11, 0.11, fill=RED_BAR, shape=MSO_SHAPE.RECTANGLE, name="M4 legend service key")
textbox(slide, LX4 + 0.2, LY4 + 0.24, 1.4, 0.18, [[("Interest + principal", {"c": GREY_TXT})]], size=9.5,
        name="M4 legend service")
callout(11.52, 5.10, 1.59, 1.18,
        [("2027E to 2030E: ", {}), ("$1.4B", {"b": True}), (" of cash vs. ", {}), ("$4.0B", {"b": True, "c": RED}),
         (" of debt service ($1.2B interest, $2.9B principal)", {})], "M4 callout", size=10)

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Business plan conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("The base case supports ", {}), ("~$730M of normalized Adj. EBITDA", {"b": True}),
    (", the basis for the valuation, but not ", {}), ("$4.3B of debt", {"b": True, "c": RED}),
    (": the business works, the balance sheet does not", {"i": True})]], 11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 6 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(5, new)
prs.save(OUT)
print("saved", OUT)
