"""Add the Precedent Transactions slide (slide 9, ahead of the Recommendation) to Project Print deck v11.

Usage: python3 build_precedents.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v15 (Precedent Transactions tab: deal multiples and restructuring precedents).
"""
import sys
from statistics import median

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
slide.placeholders[21].text_frame.text = "Precedent Transactions"
slide.placeholders[16].text_frame.text = (
    "Print assets now sell for 3x to 6x EBITDA, and even the precedent median leaves Xerox $0.4B short of its debt")

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
         "Notes: (1) EV / LTM EBITDA per merger proxies and releases; HP / Samsung printing (2016) excluded, no EBITDA "
         "disclosed; (2) Growth-stage inkjet asset, bar truncated; (3) Applied to $614M LTM Adj. EBITDA excl. $105mm "
         "IEEPA benefit; (4) Filing to plan effective date")
ftr_para(tf.add_paragraph(),
         "Sources: ARC (2024), EFI (2019) and Lexmark (2016) merger proxies; company filings and releases; Project "
         "Print model")

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


# ================================================================= LEFT: deal multiples over time
textbox(slide, L0, ROW1, LP1 - L0, 0.24, [[("Print Deal Multiples (EV / LTM EBITDA), 2005 to 2024", {"b": True}),
                                           ("1", {"sup": True, "b": True})]], size=14, name="A header")
line(slide, L0, ROW1 + 0.29, LP1, ROW1 + 0.29, INK, 1.0, name="A header rule")
# (target, year, acquirer, EV / LTM EBITDA)
deals = [
    ("Imagistics", "2005", "Océ", 7.1),
    ("IKON", "2008", "Ricoh", 8.8),
    ("Océ", "2009", "Canon", 7.7),
    ("Lexmark", "2016", "Apex / PAG", 7.0),
    ("EFI", "2019", "Siris", 15.2),
    ("InnerWorkings", "2020", "HH Global", 5.5),
    ("LSC", "2020", "Atlas (363)", 3.1),
    ("RRD", "2021", "Chatham", 5.8),
    ("ARC", "2024", "Management", 4.26),
    ("Lexmark", "2024", "Xerox", 4.72),
]
N = len(deals)
SPLIT = 5               # first decline-era deal
CAP = 9.9               # EFI bar is truncated here
Q1, MED = 4.91, 6.40
old_med = median(d[3] for d in deals[:SPLIT])
new_med = median(d[3] for d in deals[SPLIT:])
cats = [f"{d[0]} {d[1]}" for d in deals]
vals = [min(d[3], CAP) for d in deals]

C = dict(x=0.40, y=2.20, w=7.85, h=2.72, vmin=0.0, vmax=10.2)
slot = C["w"] / N
xc = [C["x"] + slot * (k + 0.5) for k in range(N)]
yv = lambda v: C["y"] + C["h"] * (C["vmax"] - v) / (C["vmax"] - C["vmin"])
BW = slot / 1.6         # bar width at gapWidth 60

# legend
box(slide, L0, 1.70, 0.30, 0.14, fill=BAND, line=None, shape=MSO_SHAPE.RECTANGLE, name="A legend band")
textbox(slide, L0 + 0.38, 1.66, 4.4, 0.2, [[("Range applied to Xerox: ", {"c": GREY_TXT}),
                                            (f"{Q1:.1f}x to {MED:.1f}x", {"b": True}),
                                            (" (Q1 to median, all 10 deals)", {"c": GREY_TXT})]], size=10,
        name="A legend band text")
line(slide, 5.05, 1.77, 5.35, 1.77, INK, 1.0, dash=MSO_LINE_DASH_STYLE.SQUARE_DOT, name="A legend median")
textbox(slide, 5.43, 1.66, 1.6, 0.2, [[("Era median", {"c": GREY_TXT})]], size=10, name="A legend median text")

# range applied to Xerox (Q1 to median of all 10 deals), behind the bars
box(slide, C["x"], yv(MED), C["w"], yv(Q1) - yv(MED), fill=BAND, line=None, shape=MSO_SHAPE.RECTANGLE,
    name="A applied band")

dpt = {k: (GREY_LT if k < SPLIT else INK) for k in range(N)}
dpt[N - 1] = RED_BAR


def deal_xml(rid):
    body = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            + bar_ser(0, "EV / LTM EBITDA", cats, vals, GREY_LT, dlbls("", None, 9, INK, show=False), dpt_fills=dpt)
            + '<c:gapWidth val="60"/><c:axId val="2401"/><c:axId val="2402"/></c:barChart>'
            + cat_ax(2401, 2402, delete=True) + val_ax(2402, 2401, C["vmin"], C["vmax"]))
    return chart_space((0, 0, 1, 1), body, rid)


# era divider, era labels and era medians
xd = C["x"] + slot * SPLIT
line(slide, xd, C["y"] - 0.30, xd, yv(0), GREY_MID, 0.75, dash=MSO_LINE_DASH_STYLE.DASH, name="A era divider")
for lo, hi, med, lbl, c, nm in ((0, SPLIT, old_med, "2005 to 2019: growth era", GREY_TXT, "old"),
                                (SPLIT, N, new_med, "2020 to 2024: decline era", INK, "new")):
    x0, x1 = C["x"] + slot * lo, C["x"] + slot * hi
    textbox(slide, x0, C["y"] - 0.30, x1 - x0, 0.2,
            [[(lbl, {"b": True, "c": c}), (f", median {med:.1f}x", {"c": c})]], size=10.5,
            align=PP_ALIGN.CENTER, name=f"A era label {nm}")
    line(slide, x0 + 0.08, yv(med), x1 - 0.08, yv(med), c, 1.0, dash=MSO_LINE_DASH_STYLE.SQUARE_DOT,
         name=f"A era median {nm}")

add_chart(slide, C["x"], C["y"], C["w"], C["h"], cats, [("EV / LTM EBITDA", vals)], deal_xml, "A deals chart")
line(slide, C["x"], yv(0), C["x"] + C["w"], yv(0), GREY_MID, 0.75, name="A baseline")

# EFI truncation mark
k_efi = 4
brk = box(slide, xc[k_efi] - BW / 2 - 0.03, yv(CAP - 1.0), BW + 0.06, 0.06, fill=WHITE, line=None,
          shape=MSO_SHAPE.PARALLELOGRAM, name="A efi break")
brk.rotation = -8

# value and category labels
for k, (tgt, yr, acq, m) in enumerate(deals):
    xerox = k == N - 1
    c = RED if xerox else INK
    runs = [(f"{m:.1f}x", {"b": True, "c": INK if k < SPLIT else WHITE})]
    if k == k_efi:
        runs.append(("2", {"sup": True}))
    textbox(slide, xc[k] - 0.4, yv(vals[k]) + 0.05, 0.8, 0.18, [runs], size=10.5, align=PP_ALIGN.CENTER,
            name=f"A value {k + 1}")
    textbox(slide, xc[k] - slot / 2 - 0.1, yv(0) + 0.05, slot + 0.2, 0.5,
            [[(tgt, {"b": True, "c": c, "size": 9.5})], [(yr, {"c": GREY_TXT, "size": 9})],
             [(acq, {"c": RED if xerox else GREY_TXT, "size": 9, "b": xerox})]],
            size=9.5, align=PP_ALIGN.CENTER, name=f"A category {k + 1}", wrap=False)

callout(L0, 5.62, LP1 - L0, 0.46,
        [("Xerox itself paid ", {}), ("4.7x", {"b": True, "c": RED}), (" for Lexmark in Dec-24: buyers now price "
                                                                       "print as a ", {}),
         ("declining business", {"b": True}), (", and Xerox is no exception", {})],
        "A callout", size=10.5)

# ================================================================= RIGHT TOP: range applied to Xerox
section_header(slide, RP0, ROW1, R1 - RP0, "Precedent Range Applied to Xerox", "B header")
CW = [1.75, 0.88, 0.88, 0.98]
TX = RP0
EB = 614
TOT = 4355
cols = [Q1, MED, new_med]
ev = [3016, 3930, round(new_med * EB)]     # all-deal Q1 / median EV per the model
# group labels over the value columns
textbox(slide, TX + CW[0], 1.70, CW[1] + CW[2], 0.18, [[("All 10 deals", {"i": True, "c": GREY_TXT})]],
        size=10, align=PP_ALIGN.RIGHT, name="B group all")
textbox(slide, TX + CW[0] + CW[1] + CW[2], 1.70, CW[3], 0.18, [[("2020 to 2024", {"i": True, "c": GREY_TXT})]],
        size=10, align=PP_ALIGN.RIGHT, name="B group recent")


def table(y0, hdr, body, cw, name, bold_rows=(), red_rows=(), rule_before=(), row_h=0.26, size=10.5, left=(0,)):
    x = TX
    for j, h in enumerate(hdr):
        runs = h if isinstance(h, list) else [(h, {})]
        textbox(slide, x, y0, cw[j], 0.2, [[(t, {"b": True, **o}) for t, o in runs]], size=size,
                align=PP_ALIGN.LEFT if j in left else PP_ALIGN.RIGHT, name=f"{name} hdr {j + 1}")
        x += cw[j]
    line(slide, TX, y0 + 0.25, TX + sum(cw), y0 + 0.25, INK, 0.75, name=f"{name} hdr rule")
    y = y0 + 0.31
    for k, row in enumerate(body):
        if k in rule_before:
            line(slide, TX, y - 0.03, TX + sum(cw), y - 0.03, INK, 0.75, name=f"{name} total rule")
        x = TX
        for j, cell in enumerate(row):
            runs = cell if isinstance(cell, list) else [(cell, {})]
            o = {"b": k in bold_rows, "c": RED if (k in red_rows and j > 0) else INK}
            textbox(slide, x, y, cw[j], row_h - 0.04, [[(t, {**o, **oo}) for t, oo in runs]], size=size,
                    align=PP_ALIGN.LEFT if j in left else PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
                    name=f"{name} r{k + 1} c{j + 1}")
            x += cw[j]
        y += row_h
    return y


money = lambda v: f"(${-v:,})" if v < 0 else f"${v:,}"
yt = table(1.92, [[("($M)", {}), ("3", {"sup": True})], "Q1", "Median", "Median"],
           [["EV / LTM Adj. EBITDA"] + [f"{m:.1f}x" for m in cols],
            ["Implied enterprise value"] + [money(v) for v in ev],
            ["(-) Total debt"] + [f"({TOT:,})"] * 3,
            ["Shortfall to total debt"] + [money(v - TOT) for v in ev]],
           CW, "B table", bold_rows=(3,), red_rows=(3,), rule_before=(3,))

# ================================================================= RIGHT BOTTOM: how comparable credits were fixed
section_header(slide, RP0, 3.42, R1 - RP0, "Prearranged Deals Beat a Free Fall", "C header")
CW2 = [1.25, 1.25, 0.70, 1.29]
yr = table(3.84, ["Company", "Path", [("Time", {}), ("4", {"sup": True})], "Outcome"],
           [[[("Diebold Nixdorf", {"b": True})], "Prepack, 2023", "~10 wks", "~$2.1B of debt cut; lenders own the equity"],
            [[("Exela", {"b": True})], "Prearranged, 2025", "~5 mos", "Secured notes converted to equity"],
            [[("Kodak", {"b": True})], "Free fall, 2012", "~20 mos",
             [("Unsecured creditors recovered ", {}), ("~4c to 5c", {"b": True, "c": RED})]]],
           CW2, "C table", row_h=0.42, size=10, left=(0, 1, 2, 3))
callout(RP0, yr + 0.08, R1 - RP0, 0.62,
        [("The ABL matures in ", {}), ("May 2028", {"b": True}), (", about ", {}), ("19 months", {"b": True, "c": RED}),
         (" away: only a prearranged process fits that runway", {})],
        "C callout", size=10.5)

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Precedents conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Deal multiples have fallen from ", {}), (f"{old_med:.1f}x to {new_med:.1f}x", {"b": True}),
    (" as print declined: precedents confirm the ", {}), ("value gap", {"b": True, "c": RED}),
    (", and comparable credits closed it fastest with ", {}), ("prearranged deals", {"i": True})]],
    11.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- move to slide 9 (before the Recommendation)
lst = prs.slides._sldIdLst
ids = list(lst)
new = ids[-1]
lst.remove(new)
lst.insert(8, new)
prs.save(OUT)
print("saved", OUT, f"old median {old_med}, new median {new_med}, EV {ev}")
