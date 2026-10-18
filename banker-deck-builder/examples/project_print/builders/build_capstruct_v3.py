"""Capital Structure Overview, v2 (slide 5) for the Project Print deck.

Usage: python3 build_capstruct_v2.py <input.pptx> <output.pptx>
"""
import sys
from decimal import Decimal, ROUND_HALF_UP

_argv = list(sys.argv)
sys.argv = sys.argv[:1]
from pp_helpers import *  # noqa: E402,F401
from pp_helpers import textbox, section_header, style_run  # noqa: E402

SRC, OUT = _argv[1], _argv[2]

LIGHT = "F2F2F2"
CALLOUT_FILL = "D6D6D6"
HAIR = "D1D1D1"
DARK_GREY = "595959"
RED_TINT = "E9A5AA"


def r0(v):
    return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def money(v):
    return f"{r0(v):,}"


# ---------------------------------------------------------------- data (6/30/26 principal; 10/14/26 prices)
EBITDA_EXT = 614.0
groups = [
    ("IPCo JV: structurally senior on contributed IP", [
        ("IPCo JV Term Loan", "IPCo (TPG JV)", "S+812.5", "Feb-31", 405, 106.00, "10.2%"),
        ("IPCo Class A Preferred", "IPCo (TPG JV)", "S+1,187.5", "Feb-31", 45, None, "n/a"),
    ]),
    ("First Lien Secured", [
        ("ABL ($425M facility)", "Corp; Holdings gtd.", "S+150", "May-28", None, None, "n/a"),
        ("Term Loan B", "Corp; Holdings gtd.", "S+400", "Nov-29", 705, 77.38, "18.7%"),
        ("10.25% 1L Notes", "Corp; Holdings gtd.", "10.250%", "Oct-30", 400, 87.59, "14.4%"),
    ]),
    ("Second Lien Secured", [
        ("13.50% 2L Notes", "Corp; Holdings gtd.", "13.500%", "Apr-31", 494, 67.70, "26.1%"),
    ]),
    ("Holdings Unsecured (Corp guarantee)", [
        ("5.50% Notes", "Holdings; Corp gtd.", "5.500%", "Aug-28", 556, 55.88, "42.3%"),
        ("8.875% Notes", "Holdings; Corp gtd.", "8.875%", "Nov-29", 500, 38.93, "48.8%"),
        ("3.75% Convertible Notes", "Holdings; Corp gtd.", "3.750%", "Mar-30", 400, 35.38, "40.0%"),
        ("13.00% Step-up Notes", "Holdings; Corp gtd.", "13.000%", "Jul-30", 250, 57.00, "35.8%"),
    ]),
    ("Corp Legacy Unsecured (no Holdings claim)", [
        ("4.80% Notes", "Corp only", "4.800%", "Mar-35", 250, 32.00, "24.0%"),
        ("6.75% Notes", "Corp only", "6.750%", "Dec-39", 350, 32.01, "23.8%"),
    ]),
]
OTHER_DEBT = 2
PREF, MCAP, CASH = 180, 354, 529


def mv(face, price):
    if face is None:
        return None
    return face * (price / 100.0) if price is not None else float(face)


# ---------------------------------------------------------------- deck + slide frame
prs = Presentation(SRC)
ref = prs.slides[2]
slide = prs.slides.add_slide(ref.slide_layout)
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 20:
        ph._element.getparent().remove(ph._element)
slide.placeholders[21].text_frame.text = "Capital Structure Overview"
slide.placeholders[16].text_frame.text = (
    "Debt trades at 68 to 88 through the 2L but 32 to 57 below it, implying a ~$2.7B TEV and "
    "impaired unsecured notes")

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
         "Notes: (1) 6/30/26 principal incl. $2mm other; ABL undrawn, $112mm LCs; (2) FactSet 10/14/26; Class A at "
         "par; total is wtd. avg.; (3) Cum. debt / $614mm LTM Adj. EBITDA ex-tariff; (4) Mkt. debt + $180mm pref. + "
         "$354mm mkt. cap less $529mm cash; (5) Face at final maturity; (6) IPCo credit agmt., SSLA \u00a76.3(b); "
         "(7) Share of 2024 PF sales")
ftr_para(tf.add_paragraph(),
         "Sources: Xerox Q2 2026 10-Q; IPCo credit agreement and SSLA (8-K, 2/17/26); S&P (6/16/26); Octus; "
         "FactSet as of 10/14/26")

L0, L1, R0, R1 = 0.23, 7.95, 8.25, 13.11
ROW1 = 1.30

# ================================================================= LEFT: debt table
section_header(slide, L0, ROW1, L1 - L0, "Debt Priority, Liens and Trading Levels ($M)", "T header")

COLS = [1.70, 1.36, 0.72, 0.60, 0.58, 0.56, 0.62, 0.55, 0.74]
T_X, T_Y = 0.52, 1.70
HGT = {"hdr": 0.27, "grp": 0.21, "tr": 0.20, "tot": 0.24, "tev": 0.22}

rows = [("hdr", None)]
for gname, trs in groups:
    rows.append(("grp", gname))
    for t in trs:
        rows.append(("tr", t))
rows.append(("tot", None))
rows.append(("tev", None))

gf = slide.shapes.add_table(len(rows), len(COLS), Inches(T_X), Inches(T_Y),
                            Inches(sum(COLS)), Inches(sum(HGT[k] for k, _ in rows)))
gf.name = "Debt table"
tbl = gf.table
tblPr = gf._element.graphic.graphicData.tbl.tblPr
sid = tblPr.find(qn("a:tableStyleId"))
if sid is None:
    sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # No Style, No Grid
tbl.first_row = False
tbl.horz_banding = False
for i, w in enumerate(COLS):
    tbl.columns[i].width = Inches(w)
for i, (kind, _) in enumerate(rows):
    tbl.rows[i].height = Inches(HGT[kind])


def ln_xml(tag, color=None, w=0.5):
    if color is None:
        return f'<a:{tag} w="0"><a:noFill/></a:{tag}>'
    return (f'<a:{tag} w="{int(w*12700)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
            f'<a:srgbClr val="{color}"/></a:solidFill><a:prstDash val="solid"/></a:{tag}>')


def style_cell(cell, runs, size=9, align=PP_ALIGN.LEFT, fill=None, bottom=None, top=None,
               bottom_w=0.5, top_w=0.5, indent=0.0):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        tcPr.remove(child)
    tcPr.set("marL", str(int(Inches(0.05 + indent))))
    tcPr.set("marR", str(int(Inches(0.05))))
    tcPr.set("marT", str(int(Inches(0.01))))
    tcPr.set("marB", str(int(Inches(0.01))))
    tcPr.set("anchor", "ctr")
    xml = ln_xml("lnL") + ln_xml("lnR") + ln_xml("lnT", top, top_w) + ln_xml("lnB", bottom, bottom_w)
    for el in etree.fromstring(f'<x xmlns:a="{A_NS}">{xml}</x>'):
        tcPr.append(el)
    if fill:
        f = etree.SubElement(tcPr, qn("a:solidFill"))
        etree.SubElement(f, qn("a:srgbClr")).set("val", fill)
    else:
        etree.SubElement(tcPr, qn("a:noFill"))
    tfc = cell.text_frame
    tfc.word_wrap = True
    p = tfc.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = align
    for text, o in runs:
        if not text:
            continue
        r = p.add_run()
        r.text = text
        style_run(r, o.get("size", size), o.get("b", False), o.get("i", False),
                  o.get("c", INK), o.get("sup", False))
    for para in tfc.paragraphs:
        end = para._p.find(qn("a:endParaRPr"))
        if end is None:
            end = etree.SubElement(para._p, qn("a:endParaRPr"))
        end.set("lang", "en-US")
        end.set("sz", str(int(size * 100)))
        end.set("dirty", "0")


HDR = [("Tranche", None), ("Issuer / Guarantor", None), ("Coupon", None), ("Maturity", None),
       ("Face", "1"), ("Price", "2"), ("Mkt. Val.", None), ("YTW", "2"), ("Cum. Lev.", "3")]
ALIGN = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.RIGHT,
         PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
R = PP_ALIGN.RIGHT

cum_face = tot_face = tot_mv = 0.0
row_y, y_cur = {}, T_Y
for ri, (kind, payload) in enumerate(rows):
    row_y[ri] = y_cur
    y_cur += HGT[kind]
    cells = [tbl.cell(ri, c) for c in range(len(COLS))]
    if kind == "hdr":
        for c, (label, sup) in enumerate(HDR):
            runs = [(label, {"b": True, "c": WHITE})]
            if sup:
                runs.append((f"({sup})", {"b": True, "c": WHITE, "sup": True}))
            style_cell(cells[c], runs, align=ALIGN[c], fill=INK)
    elif kind == "grp":
        trs = [t for g, ts in groups if g == payload for t in ts]
        g_face = sum(t[4] or 0 for t in trs)
        g_mv = sum(mv(t[4], t[5]) or 0 for t in trs)
        cum_face += g_face
        tot_face += g_face
        tot_mv += g_mv
        cells[0].merge(cells[3])
        kw = dict(fill=LIGHT, bottom=GREY_LT)
        style_cell(cells[0], [(payload, {"b": True})], **kw)
        for c in (1, 2, 3, 5, 7):
            style_cell(cells[c], [], **kw)
        style_cell(cells[4], [(money(g_face), {"b": True})], align=R, **kw)
        style_cell(cells[6], [(money(g_mv), {"b": True})], align=R, **kw)
        style_cell(cells[8], [(f"{cum_face / EBITDA_EXT:.1f}x", {"b": True})], align=R, **kw)
    elif kind == "tr":
        name, issuer, cpn, mat, face, price, ytw = payload
        m = mv(face, price)
        vals = [
            [(name, {})], [(issuer, {})], [(cpn, {})], [(mat, {})],
            [(money(face) if face is not None else "–", {})],
            [(f"{price:.2f}" if price is not None else "n/a",
              {"c": RED, "b": True} if (price is not None and price < 70) else {})],
            [(money(m) if m is not None else "–", {})],
            [(ytw, {})], [],
        ]
        for c in range(len(COLS)):
            style_cell(cells[c], vals[c], align=ALIGN[c], bottom=HAIR, indent=0.1 if c == 0 else 0.0)
    elif kind == "tot":
        tot_face += OTHER_DEBT
        tot_mv += OTHER_DEBT
        cells[0].merge(cells[3])
        kw = dict(fill=INK)
        style_cell(cells[0], [("Total Debt", {"b": True, "c": WHITE})], **kw)
        for c in (1, 2, 3, 7):
            style_cell(cells[c], [], **kw)
        style_cell(cells[4], [(money(tot_face), {"b": True, "c": WHITE})], align=R, **kw)
        style_cell(cells[5], [(f"{100 * tot_mv / tot_face:.1f}", {"b": True, "c": WHITE})], align=R, **kw)
        style_cell(cells[6], [(money(tot_mv), {"b": True, "c": WHITE})], align=R, **kw)
        style_cell(cells[8], [(f"{tot_face / EBITDA_EXT:.1f}x", {"b": True, "c": WHITE})], align=R, **kw)
    else:  # market-implied TEV
        tev = tot_mv + PREF + MCAP - CASH
        cells[0].merge(cells[3])
        kw = dict(fill=LIGHT, bottom=INK, bottom_w=1.0)
        style_cell(cells[0], [("Market-Implied TEV", {"b": True}), ("(4)", {"b": True, "sup": True})], **kw)
        for c in (1, 2, 3, 4, 5, 7):
            style_cell(cells[c], [], **kw)
        style_cell(cells[6], [(money(tev), {"b": True})], align=R, **kw)
        style_cell(cells[8], [(f"{tev / EBITDA_EXT:.1f}x", {"b": True})], align=R, **kw)
T_BOTTOM = y_cur

# payment priority arrow
pa = line(slide, 0.36, T_Y + HGT["hdr"] + 0.02, 0.36, T_BOTTOM - HGT["tev"] - 0.02, GREY_MID, 1.25,
          name="Payment priority arrow")
tail = etree.SubElement(pa.line._get_or_add_ln(), qn("a:tailEnd"))
tail.set("type", "triangle")
tail.set("w", "med")
tail.set("len", "med")
pl = textbox(slide, 0, 0, 1.25, 0.18, [[("Payment priority", {"i": True, "c": GREY_TXT})]], size=9,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name="Payment priority label")
pl.rotation = 270
pl.fill.solid()
pl.fill.fore_color.rgb = rgb(WHITE)
pl.left = Inches(0.36 - 0.625)
pl.top = Inches((T_Y + T_BOTTOM - HGT["tev"]) / 2 - 0.09)

# dotted red box around the unsecured layers, with a "likely fulcrum" pill
first_u = [i for i, (k, p) in enumerate(rows) if k == "grp" and p.startswith("Holdings")][0]
last_u = [i for i, (k, p) in enumerate(rows) if k == "tot"][0] - 1
fx0, fx1 = T_X - 0.04, T_X + sum(COLS) + 0.04
fy0, fy1 = row_y[first_u] - 0.015, row_y[last_u] + HGT["tr"] + 0.015
fb = box(slide, fx0, fy0, fx1 - fx0, fy1 - fy0, fill=None, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT,
         line_w=1.5, radius=0.03, name="Fulcrum box")
fb.text_frame.text = ""
pill_x = T_X + sum(COLS[:2]) + 0.04
fp = box(slide, pill_x, row_y[first_u] + 0.015, sum(COLS[2:4]) - 0.08, HGT["grp"] - 0.03, fill=WHITE,
         line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0, radius=0.5, name="Fulcrum pill")
fp.text_frame.margin_top = fp.text_frame.margin_bottom = 0
fp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(fp.text_frame, [[("Likely fulcrum", {"b": True, "c": RED})]], 8.5, align=PP_ALIGN.CENTER)

# stakeholders strip
SH_Y = T_BOTTOM + 0.10
sh = box(slide, L0, SH_Y, L1 - L0, 0.30, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.25, name="Advisors strip")
sh.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(sh.text_frame, [[
    ("Company: ", {"b": True}), ("Lazard / Kirkland & Ellis", {}), ("     ", {}),
    ("Crossover group: ", {"b": True}), ("Gibson Dunn / Moelis", {}), ("     ", {}),
    ("2L / unsecured group: ", {"b": True}), ("Paul Hastings", {}),
]], 10, align=PP_ALIGN.CENTER)

# ================================================================= RIGHT TOP: structure
section_header(slide, R0, ROW1, R1 - R0, "Corporate and Guarantee Structure", "S header")


def arrow(x1, y1, x2, y2, color, w=1.0, name=None):
    c = line(slide, x1, y1, x2, y2, color, w, name=name)
    t = etree.SubElement(c.line._get_or_add_ln(), qn("a:tailEnd"))
    t.set("type", "triangle")
    t.set("w", "sm")
    t.set("len", "sm")
    return c


def card(x, y, w, title, body_paras, body_h, hdr_fill=INK, hdr_color=WHITE, hdr_line=None,
         body_line=GREY_MID, body_dash=None, name="card", hdr_h=0.22, body_size=9, hdr_size=10.5):
    hb = box(slide, x, y, w, hdr_h, fill=hdr_fill, line=hdr_line, line_w=1.0, radius=0.15,
             name=f"{name} header")
    hb.text_frame.margin_top = hb.text_frame.margin_bottom = 0
    hb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(hb.text_frame, [[(title, {"b": True, "c": hdr_color})]], hdr_size, align=PP_ALIGN.CENTER)
    bb = box(slide, x, y + hdr_h + 0.03, w, body_h, fill=WHITE, line=body_line, dash=body_dash,
             line_w=0.75 if body_dash is None else 1.0, radius=0.06, name=f"{name} body")
    bb.text_frame.margin_top = bb.text_frame.margin_bottom = Inches(0.02)
    bb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(bb.text_frame, body_paras, body_size, align=PP_ALIGN.CENTER)
    return y + hdr_h + 0.03 + body_h


SX0, SX1 = R0, 11.05          # Holdings / Corp / sub cards column
IX0, IX1 = 11.68, R1          # IPCo column
HOLD_Y = 1.68
hold_bot = card(SX0, HOLD_Y, SX1 - SX0, "Xerox Holdings Corp.",
                [[("Issuer of ", {}), ("$1,706M", {"b": True, "c": RED}), (" unsecured notes", {})]],
                0.24, name="Entity Holdings")
CORP_Y = hold_bot + 0.38
corp_bot = card(SX0, CORP_Y, SX1 - SX0, "Xerox Corporation",
                [[("Borrower: ", {"b": True}), ("ABL, TLB, 1L, 2L ($1,599M)", {})],
                 [("Issuer: ", {"b": True}), ("legacy 2035s / 2039s ($600M)", {})]],
                0.38, name="Entity Corp")

# double-dip: two claims into Corp
for k, xx in enumerate((SX0 + 0.14, SX1 - 0.14)):
    arrow(xx, hold_bot, xx, CORP_Y, RED, 1.25, name=f"Double-dip arrow {k+1}")
dd = box(slide, SX0 + 0.30, hold_bot + 0.04, (SX1 - SX0) - 0.60, 0.31, fill=WHITE, line=RED,
         dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0, radius=0.5, name="Double-dip pill")
dd.text_frame.margin_top = dd.text_frame.margin_bottom = 0
dd.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(dd.text_frame, [[("Double-dip", {"b": True, "c": RED})],
                         [("intercompany loan + guarantee", {"i": True})]], 8.5, align=PP_ALIGN.CENTER)

# IPCo JV alongside Corp, linked only by royalty / license
ipco_bot = card(IX0, CORP_Y, IX1 - IX0, "IPCo JV (TPG)",
                [[("$405M TL + $45M pref.", {})], [("Xerox holds Class B", {})]],
                0.38, hdr_fill=WHITE, hdr_color=RED, hdr_line=RED, body_line=RED,
                body_dash=MSO_LINE_DASH_STYLE.ROUND_DOT, name="Entity IPCo")
ay1, ay2 = CORP_Y + 0.30, CORP_Y + 0.50
arrow(SX1, ay1, IX0, ay1, INK, 1.0, name="Royalty arrow")
arrow(IX0, ay2, SX1, ay2, GREY_MID, 1.0, name="License arrow")
textbox(slide, SX1, CORP_Y - 0.12, IX0 - SX1, 0.40, [[("2% royalty", {"b": True})], [("~$150M/yr", {"i": True})]],
        size=9, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, name="Royalty label")
textbox(slide, SX1, ay2 + 0.03, IX0 - SX1, 0.17, [[("IP license", {"i": True, "c": GREY_TXT})]], size=9,
        align=PP_ALIGN.CENTER, name="License label")
textbox(slide, SX1 + 0.12, ipco_bot + 0.10, IX1 - SX1 - 0.12, 0.36,
        [[("Not a “Subsidiary” under debt docs; consolidated as a VIE", {"i": True, "c": GREY_TXT})]],
        size=9, align=PP_ALIGN.CENTER, name="IPCo note")

# guarantor / non-guarantor subsidiaries under Corp
BR_Y = corp_bot + 0.08
SUB_Y = BR_Y + 0.08
cw = (SX1 - SX0 - 0.12) / 2
cx = [SX0, SX0 + cw + 0.12]
mid = (SX0 + SX1) / 2
line(slide, mid, corp_bot, mid, BR_Y, GREY_MID, 1.0, name="Branch stem")
line(slide, cx[0] + cw / 2, BR_Y, cx[1] + cw / 2, BR_Y, GREY_MID, 1.0, name="Branch bar")
for k in range(2):
    line(slide, cx[k] + cw / 2, BR_Y, cx[k] + cw / 2, SUB_Y, GREY_MID, 1.0, name=f"Branch drop {k+1}")
sub_bot = card(cx[0], SUB_Y, cw, "Guarantor subs", [[("Guarantee secured debt", {})], [("and Holdings notes", {})]],
               0.40, name="Entity guarantor subs", hdr_size=9.5)
card(cx[1], SUB_Y, cw, "Non-guarantor subs",
     [[("~26% of 2024 PF sales", {}), ("(7)", {"sup": True})]], 0.40, hdr_fill=GREY_MID,
     name="Entity non-guarantor subs", hdr_size=9.5)

# S&P value split bar, label inline
SP_Y, SP_H = sub_bot + 0.12, 0.22
LBL_W = 1.05
textbox(slide, R0, SP_Y, LBL_W, SP_H, [[("S&P value split", {"i": True, "c": GREY_TXT})]], size=9,
        anchor=MSO_ANCHOR.MIDDLE, name="S&P split label")
segs = [("Obligors 55%", 0.55, INK, WHITE), ("JV 17%", 0.17, RED_BAR, WHITE),
        ("Non-obligors 28%", 0.28, GREY_LT, INK)]
x = R0 + LBL_W
for lab, share, fillc, tc in segs:
    w = (R1 - R0 - LBL_W) * share
    s_ = box(slide, x, SP_Y, w, SP_H, fill=fillc, line=WHITE, line_w=0.75, shape=MSO_SHAPE.RECTANGLE,
             name=f"S&P split {lab}")
    s_.text_frame.margin_top = s_.text_frame.margin_bottom = 0
    s_.text_frame.margin_left = s_.text_frame.margin_right = Inches(0.02)
    s_.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(s_.text_frame, [[(lab, {"b": True, "c": tc})]], 9, align=PP_ALIGN.CENTER)
    x += w

# ================================================================= RIGHT BOTTOM: maturities by priority
ROW2 = SP_Y + SP_H + 0.10
section_header(slide, R0, ROW2, R1 - R0, "Maturity Profile by Priority ($M)", "M header")
mh = [s for s in slide.shapes if s.name == "M header"][0]
r = mh.text_frame.paragraphs[0].add_run()
r.text = "(5)"
style_run(r, 14, True, sup=True)

cats = ["2028", "2029", "2030", "2031", "2035", "2039"]
series = [
    ("IPCo JV", [None, None, None, 450, None, None], GREY_LT, INK),
    ("1L", [None, 705, 400, None, None, None], INK, WHITE),
    ("2L", [None, None, None, 494, None, None], DARK_GREY, WHITE),
    ("Holdings unsecured", [556, 500, 650, None, None, None], RED_BAR, WHITE),
    ("Legacy Corp unsecured", [None, None, None, None, 250, 350], RED_TINT, INK),
]
totals = [556, 1205, 1050, 944, 250, 350]

# legend
LGY = ROW2 + 0.36
lx = R0 + 0.02
for lab, _, colr, _ in series:
    box(slide, lx, LGY + 0.035, 0.11, 0.11, fill=colr, shape=MSO_SHAPE.RECTANGLE, name=f"Legend key {lab}")
    textbox(slide, lx + 0.15, LGY, 1.6, 0.18, [[(lab, {"c": GREY_TXT})]], size=9, name=f"Legend {lab}")
    lx += 0.15 + 0.18 + len(lab) * 0.060

CM_Y = LGY + 0.20
CM = dict(x=R0, y=CM_Y, w=R1 - R0, h=6.36 - CM_Y, px=0.02, pw=0.96, py=0.30, ph=0.58, vmax=1450, gap=55)


def mat_xml(rid):
    body = '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
    for i, (lab, vals, colr, tc) in enumerate(series):
        body += bar_ser(i, f"{lab} ($M)", cats, vals, colr, dlbls('#,##0', "ctr", 9, tc, True))
    body += (f'<c:gapWidth val="{CM["gap"]}"/><c:overlap val="100"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>'
             + cat_ax(5001, 5002, line_color=GREY_MID) + val_ax(5002, 5001, 0, CM["vmax"]))
    return chart_space((CM["px"], CM["py"], CM["pw"], CM["ph"]), body, rid)


add_chart(slide, CM["x"], CM["y"], CM["w"], CM["h"], cats, [(f"{l} ($M)", v) for l, v, _, _ in series],
          mat_xml, "Maturity profile by priority chart")

pm_x0 = CM["x"] + CM["px"] * CM["w"]
pm_w = CM["pw"] * CM["w"]
pm_y0 = CM["y"] + CM["py"] * CM["h"]
pm_h = CM["ph"] * CM["h"]
pm_bot = pm_y0 + pm_h
slot = pm_w / len(cats)
bar_w = slot / (1 + CM["gap"] / 100)
cent = [pm_x0 + slot * (k + 0.5) for k in range(len(cats))]


def ym(v):
    return pm_bot - pm_h * v / CM["vmax"]


for k, v in enumerate(totals):
    textbox(slide, cent[k] - 0.4, ym(v) - 0.20, 0.8, 0.18, [[(f"${v:,}M", {"b": True})]], size=10,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, name=f"Maturity total {k+1}")

# "three hurdles" bracket over 2028 to 2030
bx0, bx1 = cent[0] - bar_w / 2, cent[2] + bar_w / 2
by = ym(1205) - 0.26
line(slide, bx0, by, bx1, by, RED, 1.0, name="Hurdles bracket")
line(slide, bx0, by, bx0, by + 0.06, RED, 1.0, name="Hurdles bracket L")
line(slide, bx1, by, bx1, by + 0.06, RED, 1.0, name="Hurdles bracket R")
hp = box(slide, (bx0 + bx1) / 2 - 0.62, by - 0.23, 1.24, 0.20, fill=WHITE, line=RED,
         dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0, radius=0.5, name="Hurdles pill")
hp.text_frame.margin_top = hp.text_frame.margin_bottom = 0
hp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(hp.text_frame, [[("Three hurdles", {"b": True, "c": RED})]], 9, align=PP_ALIGN.CENTER)

# ================================================================= bottom callout
co = box(slide, L0, 6.43, R1 - L0, 0.36, fill=CALLOUT_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT,
         line_w=1.0, radius=0.2, name="IPCo change-of-control callout")
co.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(co.text_frame, [[
    ("IPCo change of control", {"b": True}), ("(6)", {"b": True, "sup": True}),
    (": any holder or creditor group taking >50% of Holdings equity, in or out of court, lets IPCo "
     "terminate the Xerox brand license, so an equitization ", {}),
    ("needs TPG consent or a take-out", {"b": True, "c": RED}),
]], 10.5, align=PP_ALIGN.CENTER)

prs.save(OUT)
print("saved", OUT, "| MV", round(tot_mv, 1), "| face", tot_face, "| wtd px", round(100 * tot_mv / tot_face, 1),
      "| TEV", round(tot_mv + PREF + MCAP - CASH, 1), "| table bottom", round(T_BOTTOM, 2),
      "| strip bottom", round(SH_Y + 0.30, 2), "| chart y", round(CM_Y, 2))
