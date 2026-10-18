"""Replace the Preliminary Recommendation slide (slide 13) with the RSA-backed prepack recommendation.

Usage: python3 build_rec2.py <input.pptx> <output.pptx>
Numbers: Xerox_RX_Model_v16 ($XRX Transaction Analysis, status quo vs. Case B prepack; Case A exchange).
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
# drop the old recommendation slide (last slide)
lst = prs.slides._sldIdLst
old = list(lst)[-1]
assert "Recommendation" in prs.slides[len(prs.slides) - 1].placeholders[21].text_frame.text
prs.part.drop_rel(old.get(qn("r:id")))
lst.remove(old)

ref = prs.slides[2]
slide = prs.slides.add_slide(ref.slide_layout)
for ph in list(slide.placeholders):
    if ph.placeholder_format.idx == 20:
        ph._element.getparent().remove(ph._element)
slide.placeholders[21].text_frame.text = "Preliminary Recommendation"
slide.placeholders[16].text_frame.text = (
    "File an RSA-backed prepack that hands Xerox to the fulcrum notes, halves the debt and clears every maturity to 2031")

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
         "Notes: (1) Status quo at 12/31/26E vs. pro forma at close; 2027E Adj. EBITDA $728M; (2) 90% participation; "
         "holdouts and legacy notes stay ahead of the new equity; (3) Mid plan EV $3,357M; 37c to 67c across the range; "
         "(4) Plan equity $842M after $881M after-tax pension; (5) 11 U.S.C. 1129(b)")
ftr_para(tf.add_paragraph(),
         "Sources: Project Print model ($XRX Transaction Analysis, $XRX Recovery Waterfall); FactSet as of 10/14/26")

L0, R1 = 0.23, 13.11
LP1 = 8.33              # left panel right edge
RP0 = 8.62              # right panel left edge
ROW1 = 1.30
NOLBL = dlbls("", None, 9, INK, show=False)


def callout(x, y, w, h, runs, name, size=10.5):
    c = box(slide, x, y, w, h, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
            radius=0.08, name=name)
    c.text_frame.margin_left = c.text_frame.margin_right = Inches(0.08)
    c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(c.text_frame, [runs], size, align=PP_ALIGN.CENTER)
    return c


def col_chart(x, y, w, h, cats, vals, fills, vmax, name, gap=70):
    def xml(rid):
        body = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
                + bar_ser(0, name, cats, vals, fills[0], NOLBL, dpt_fills=dict(enumerate(fills)))
                + f'<c:gapWidth val="{gap}"/><c:axId val="3201"/><c:axId val="3202"/></c:barChart>'
                + cat_ax(3201, 3202, delete=True) + val_ax(3202, 3201, 0, vmax))
        return chart_space((0, 0, 1, 1), body, rid)
    add_chart(slide, x, y, w, h, cats, [(name, vals)], xml, name)
    slot = w / len(cats)
    xs = [x + slot * (k + 0.5) for k in range(len(cats))]
    yv = lambda v: y + h * (1 - v / vmax)
    line(slide, x, y + h, x + w, y + h, GREY_MID, 0.75, name=f"{name} baseline")
    return xs, yv, slot


# ================================================================= LEFT TOP: before and after
textbox(slide, L0, ROW1, LP1 - L0, 0.24, [[("Before and After the Prepack", {"b": True}),
                                           ("1", {"sup": True, "b": True})]], size=14, name="A header")
line(slide, L0, ROW1 + 0.29, LP1, ROW1 + 0.29, INK, 1.0, name="A header rule")
minis = [("Total debt", 4324, 2234, lambda v: f"${v / 1000:.1f}B", 5600),
         ("Net debt / 2027E EBITDA", 4.95, 2.25, lambda v: f"{v:.1f}x", 6.4),
         ("2027E cash interest", 363, 189, lambda v: f"${v:,.0f}M", 470),
         ("Maturities 2027 to 2030", 2644, 0, lambda v: f"${v / 1000:.1f}B" if v else "$0", 3420)]
MW, MG = 1.80, 0.27
MY, MH = 2.02, 1.30
for k, (title, a, b, fmt, vmax) in enumerate(minis):
    x = L0 + k * (MW + MG)
    textbox(slide, x, 1.66, MW, 0.2, [[(title, {"b": True})]], size=10.5, align=PP_ALIGN.CENTER,
            name=f"A title {k + 1}")
    xs, yv, slot = col_chart(x, MY, MW, MH, ["Status quo", "Prepack"], [a, b], [RED_BAR, INK], vmax,
                             f"A mini {k + 1}")
    for j, (v, c) in enumerate(((a, RED), (b, INK))):
        textbox(slide, xs[j] - 0.45, yv(v) - 0.22, 0.9, 0.2, [[(fmt(v), {"b": True, "c": c})]], size=10.5,
                align=PP_ALIGN.CENTER, name=f"A value {k + 1}-{j + 1}")
        textbox(slide, xs[j] - 0.45, MY + MH + 0.04, 0.9, 0.2, [[(("Status quo", "Prepack")[j], {"c": GREY_TXT})]],
                size=9.5, align=PP_ALIGN.CENTER, name=f"A cat {k + 1}-{j + 1}")
textbox(slide, L0, 3.66, LP1 - L0, 0.22,
        [[("Liquidity at 2030E goes from a ", {}), ("$2.1B shortfall", {"b": True, "c": RED}), (" to ", {}),
          ("$1.4B", {"b": True}), (", with net leverage below 2.0x by 2028", {})]], size=10.5, align=PP_ALIGN.CENTER,
        name="A liquidity line")

# ================================================================= LEFT BOTTOM: path to emergence
section_header(slide, L0, 4.05, LP1 - L0, "Path to Emergence", "B header")
steps = [("Dec '26", "Sign RSA with the", "Holdings crossover group", INK),
         ("Jan '27", "Solicit votes; commit", "the exit term loan", INK),
         ("Feb '27", "File after the 2/17 IPCo", "premium step-down", INK),
         ("Q2 '27", "Confirm and emerge", "(Diebold: ~10 weeks)", INK),
         ("May '28", "ABL springer: 12+", "months of cushion", RED)]
TL_Y = 4.86
tx = [0.95 + k * 1.62 for k in range(len(steps))]
line(slide, tx[0], TL_Y, tx[3], TL_Y, INK, 1.5, name="B timeline")
line(slide, tx[3], TL_Y, tx[4], TL_Y, RED, 1.5, dash=MSO_LINE_DASH_STYLE.DASH, name="B timeline gap")
for k, (d, l1, l2, c) in enumerate(steps):
    box(slide, tx[k] - 0.08, TL_Y - 0.08, 0.16, 0.16, fill=c, line=WHITE, line_w=1.0, shape=MSO_SHAPE.OVAL,
        name=f"B dot {k + 1}")
    textbox(slide, tx[k] - 0.75, TL_Y - 0.36, 1.5, 0.2, [[(d, {"b": True, "c": c})]], size=10.5,
            align=PP_ALIGN.CENTER, name=f"B date {k + 1}")
    textbox(slide, tx[k] - 0.8, TL_Y + 0.14, 1.6, 0.38, [[(l1, {"c": GREY_TXT})], [(l2, {"c": GREY_TXT})]],
            size=9.5, align=PP_ALIGN.CENTER, name=f"B step {k + 1}")
callout(L0, 5.72, LP1 - L0, 0.56,
        [("Legacy notes may vote no at 26c vs. a 32c market price. ", {}),
         ("Cramdown", {"b": True}), ("5", {"sup": True}),
         (" still works: the Holdings class accepts, equity gets nothing, and 26c beats ", {}),
         ("7.9c in liquidation", {"b": True, "c": RED})], "B callout", size=10.5)

# ================================================================= RIGHT TOP: why the holdings notes sign
section_header(slide, RP0, ROW1, R1 - RP0, "Why the Holdings Notes Sign", "C header")
textbox(slide, RP0, 1.66, R1 - RP0, 0.2, [[("Holdings notes recovery, cents per $1 of face", {"i": True, "c": GREY_TXT})]],
        size=10, name="C caption")
paths = [(["Out-of-court", "exchange"], "2", 20.4, GREY_LT),
         (["Strict-priority", "waterfall"], "", 36.0, GREY_LT),
         (["Market price", "10/14/26"], "", 46.3, GREY_MID),
         (["Prepack", "plan"], "3", 51.9, INK)]
xs, yv, slot = col_chart(RP0 + 0.05, 2.12, R1 - RP0 - 0.1, 1.00, [" ".join(p[0]) for p in paths], [p[2] for p in paths],
                         [p[3] for p in paths], 66, "C recovery chart", gap=60)
for k, (two, sup, v, c) in enumerate(paths):
    hot = k == 3
    textbox(slide, xs[k] - 0.5, yv(v) - 0.22, 1.0, 0.2,
            [[(f"{v:.0f}c", {"b": True, "c": INK}), (sup, {"sup": True})]], size=10.5, align=PP_ALIGN.CENTER,
            name=f"C value {k + 1}")
    textbox(slide, xs[k] - slot / 2, 3.16, slot, 0.36, [[(two[0], {"b": hot})], [(two[1], {"b": hot, "c": GREY_TXT})]],
            size=9.5, align=PP_ALIGN.CENTER, name=f"C cat {k + 1}")
textbox(slide, RP0, 3.55, R1 - RP0, 0.2,
        [[("Only the prepack pays the fulcrum ", {}), ("above where it trades today", {"b": True})]], size=10.5,
        align=PP_ALIGN.CENTER, name="C note")

# ================================================================= RIGHT BOTTOM: plan terms by class
section_header(slide, RP0, 3.86, R1 - RP0, "Plan Terms by Class", "D header")
CW = [1.32, 2.45, 0.72]
TY = 4.24
for j, h in enumerate(["Class", "Treatment", "Recovery"]):
    textbox(slide, RP0 + sum(CW[:j]), TY, CW[j], 0.2, [[(h, {"b": True})]], size=10.5,
            align=PP_ALIGN.RIGHT if j == 2 else PP_ALIGN.LEFT, name=f"D hdr {j + 1}")
line(slide, RP0, TY + 0.25, R1, TY + 0.25, INK, 0.75, name="D hdr rule")
terms = [("Secured", "Extended to 2031-32 for a 1% fee", "100%"),
         ("IPCo JV", "Repaid by a $461M exit term loan", "100%"),
         ("Holdings notes", "85% of new equity + 10c takeback notes", "52c"),
         ("Legacy notes", "15% of new equity + 5c takeback notes", "26c"),
         ("Pension, trade", "Unimpaired", "100%"),
         ("Pref., common", "Cancelled", "0c")]
RH = 0.27
y = TY + 0.30
box(slide, RP0, y + 2 * RH - 0.01, R1 - RP0, 2 * RH, fill=BAND, line=None, shape=MSO_SHAPE.RECTANGLE, name="D highlight")
for k, (c0, c1, c2) in enumerate(terms):
    hot = k in (2, 3)
    cells = [[(c0, {"b": hot}), ("4", {"sup": True})] if k == 2 else [(c0, {"b": hot})],
             [(c1, {"c": INK if hot else GREY_TXT})],
             [(c2, {"b": True, "c": RED if hot else INK})]]
    for j, runs in enumerate(cells):
        textbox(slide, RP0 + sum(CW[:j]) + (0.04 if j == 0 else 0), y, CW[j] - (0.04 if j == 0 else 0.04), RH - 0.02,
                [runs], size=10 if j == 1 else 10.5, align=PP_ALIGN.RIGHT if j == 2 else PP_ALIGN.LEFT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"D r{k + 1} c{j + 1}")
    y += RH
    if k < len(terms) - 1 and k not in (1, 2, 3):
        line(slide, RP0, y - 0.01, R1, y - 0.01, HAIR, 0.5, name=f"D rule {k + 1}")

# ================================================================= closing line
cl = box(slide, L0, 6.43, R1 - L0, 0.36, fill=WHITE, line=INK, dash=MSO_LINE_DASH_STYLE.DASH, line_w=0.75,
         radius=0.2, name="Recommendation conclusion")
cl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(cl.text_frame, [[
    ("Value breaks in the unsecured notes, so the unsecured notes should own Xerox: the prepack ", {}),
    ("pays them above market", {"b": True, "c": RED}), (" and leaves a business that ", {}),
    ("can carry its debt", {"i": True})]], 11.5, align=PP_ALIGN.CENTER)

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
