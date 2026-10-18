"""Append the 8 appendix pages (Excel model outputs + commentary) to Project Print deck v16.

Usage: python3 build_appendix.py <input.pptx> <output.pptx>
Images: xl_img/*.png rendered from the recalculated model by render_xl.py.
"""
import sys

_argv = list(sys.argv)
sys.argv = sys.argv[:1]
from PIL import Image  # noqa: E402
from pp_helpers import *  # noqa: E402,F401,F403
from pp_helpers import textbox, section_header  # noqa: E402

SRC, OUT = _argv[1], _argv[2]
GREY_FILL = "D6D6D6"
L0, R1, ROW1 = 0.23, 13.11, 1.30
IMG_Y = 1.72

prs = Presentation(SRC)
ref = prs.slides[2]


def new_slide(title, takeaway, source):
    s = prs.slides.add_slide(ref.slide_layout)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 20:
            ph._element.getparent().remove(ph._element)
    s.placeholders[21].text_frame.text = title
    s.placeholders[16].text_frame.text = takeaway
    tree = s.shapes._spTree
    for sp in ref.shapes._spTree.iterchildren(qn("p:sp")):
        ph = sp.find(".//" + qn("p:ph"))
        if ph is not None and ph.get("type") in ("sldNum", "ftr"):
            tree.append(copy.deepcopy(sp))
    f = [x for x in s.placeholders if x.placeholder_format.idx == 18][0]
    f.top = Inches(6.93)
    tf = f.text_frame
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in list(tf.paragraphs[0].runs):
        r._r.getparent().remove(r._r)
    r = tf.paragraphs[0].add_run()
    r.text = source
    r.font.size = Pt(10)
    return s


def img(s, name, x, y, w=None, h=None, header=None, hw=None):
    path = f"xl_img/{name}.png"
    W, H = Image.open(path).size
    if w is None:
        w = h * W / H
    if h is None:
        h = w * H / W
    if header:
        section_header(s, x, y - 0.42, hw or w, header, f"hdr {name}")
    pic = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    pic.name = f"Excel output {name}"
    return x + w, y + h


def commentary(s, x, points, call, extra=None):
    w = R1 - x
    section_header(s, x, ROW1, w, "Commentary", "commentary header")
    paras = []
    for k, runs in enumerate(points):
        paras.append([(f"({k + 1}) ", {"b": True})] + runs)
    tb = textbox(s, x, 1.74, w, 3.0, paras, size=10.5, space_after=7, name="commentary text")
    if extra:
        extra(s, x, w)
    c = box(s, x, 5.62, w, 0.66, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
            radius=0.08, name="commentary callout")
    c.text_frame.margin_left = c.text_frame.margin_right = Inches(0.08)
    c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(c.text_frame, [call], 10.5, align=PP_ALIGN.CENTER)
    return tb


B = {"b": True}
R = {"b": True, "c": RED}
SRC_M = "Source: Project Print model"

# ---------------------------------------------------------------- A1 revenue build
s = new_slide("Appendix: Revenue Build",
              "Base case revenue falls ~3.7% a year to $6.5B by 2030E; IT Solutions is held near flat",
              f"{SRC_M} ($XRX Revenue Build); Xerox 10-Ks and Q2'26 10-Q; Lexmark pre-acquisition revenue per 8-K")
xe, _ = img(s, "rev", L0, IMG_Y, w=9.22, header="Revenue Build, Base Case ($M)")
commentary(s, xe + 0.27, [
    [("2026E growth of 8.2% is the full-year Lexmark consolidation; ", {}), ("pro forma revenue falls 4.5%", B),
     (" from $7,961M in 2025", {})],
    [("Equipment (-5% to -6% a year) and post-sale (-4% to -4.5%) decline with the print market; post-sale stays "
      "~69% of revenue", {})],
    [("IT Solutions is held at $726M to $770M: ", {}), ("no credit", B), (" for a diversification turnaround", {})]],
    [("Revenue falls from ", {}), ("$7.6B", B), (" in 2026E to ", {}), ("$6.5B", R), (" in 2030E", {})])

# ---------------------------------------------------------------- A2 operating assumptions
s = new_slide("Appendix: Operating Assumptions",
              "Base case margins reach 6.5% by 2028E on Lexmark synergies, between a repeat of 2025 and management's plan",
              f"{SRC_M} ($XRX DCF Operating Assumptions); Xerox 2025 guidance and Q2'26 earnings materials")
xe, _ = img(s, "opa", L0, IMG_Y, w=9.22, header="Income Statement Drivers by Case ($M)")
commentary(s, xe + 0.27, [
    [("Adj. operating margin: base ", {}), ("5.6% to 6.5%", B), ("; bear 5.0% to 5.25% (2025 misses repeat); bull "
                                                                   "6.25% to 7.5% (management plan)", {})],
    [("Adj. EBITDA excl. the IEEPA tariff benefit ", {}), ("peaks at $728M", B), (" in 2027E and erodes to $670M by "
                                                                                  "2030E", {})],
    [("Restructuring cash falls from $85M to $98M in 2026E to 2027E to ~$36M by 2030E; tax at 25% of EBIT", {})]],
    [("Base 2027E EBITDA is ", {}), ("$63M below", R), (" management's $791M plan", {})])

# ---------------------------------------------------------------- A3 DCF
s = new_slide("Appendix: Discounted Cash Flow Detail",
              "A 12.0% reorganized WACC yields $2.9B to $3.2B; no WACC and exit multiple combination reaches the debt",
              f"{SRC_M} ($XRX DCF); FactSet as of 10/14/26 (risk-free rate, betas); Kroll cost of capital data")
_, yt = img(s, "dcf_top", L0, IMG_Y, w=9.05, header="Free Cash Flow Build, Base Case ($M)", hw=9.05)
img(s, "dcf_ev", L0, yt + 0.55, w=8.2, header="Terminal Value and Enterprise Value ($M)", hw=9.05)
section_header(s, 9.62, ROW1, R1 - 9.62, "WACC Build (Reorganized Company)", "hdr wacc")
s.shapes.add_picture("xl_img/dcf_wacc.png", Inches(9.62), Inches(IMG_Y), Inches(R1 - 9.62)).name = "Excel output dcf_wacc"
c = box(s, 9.62, 5.62, R1 - 9.62, 0.66, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
        radius=0.08, name="dcf callout")
c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(c.text_frame, [[("Terminal value is ", {}), ("65%", B), (" of EV; even 10.0% and 6.0x leaves ", {}),
                        ("$528M short", R)]], 10.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- A4 liquidation
s = new_slide("Appendix: Liquidation Analysis Detail",
              "Net proceeds of $2.1B to $2.9B cover secured claims in full but leave unsecured creditors 2c to 13c",
              f"{SRC_M} ($XRX Liquidation Analysis); Xerox Q2'26 10-Q balance sheet (6/30/26)")
xe, ya = img(s, "liq_a", L0, IMG_Y, w=5.0, header="Hypothetical Chapter 7 Liquidation ($M)")
img(s, "liq_b", L0, ya + 0.10, w=5.0)
commentary(s, xe + 0.40, [
    [("Receivables recover ", {}), ("50% to 90%", B), ("; goodwill, intangibles and deferred tax assets recover "
                                                       "~0% in a forced sale", {})],
    [("Wind-down, trustee and advisor costs absorb 11% to 14% of gross proceeds", {})],
    [("$7.2B of unsecured claims share $156M to $955M; ", {}), ("trade and other liabilities ($3.7B)", B),
     (" are the largest claim", {})],
    [("Even with trade claims halved, unsecured recovery is only ", {}), ("~11c", B),
     (" at mid, far below 26c to 52c under the plan", {})]],
    [("Best-interests test: every impaired class recovers ", {}), ("more under the plan", R)])

# ---------------------------------------------------------------- A5 recovery waterfall
s = new_slide("Appendix: Recovery Waterfall Detail",
              "Secured claims are covered at every plan value, and the unsecured notes are the fulcrum in all cases",
              f"{SRC_M} ($XRX Recovery Waterfall); sensitivity cases recalculated in the model with each claim switched off")
_, yb = img(s, "wf_dv", L0, IMG_Y, w=8.0, header="Distributable Value ($M)", hw=8.0)
img(s, "wf", L0, yb + 0.12, w=8.0)


def sens(s, x, w):
    y0 = 3.62
    textbox(s, x, y0, w, 0.2, [[("Mid recovery: Holdings / legacy", {"i": True, "c": GREY_TXT})]], size=10,
            name="sens caption")
    rows = [("Base (both claims allowed)", "36% / 18%"), ("Pension claim excluded", "46% / 23%"),
            ("No Holdings double-dip", "27% / 27%"), ("Both excluded", "40% / 40%")]
    y = y0 + 0.26
    line(s, x, y - 0.03, x + w, y - 0.03, INK, 0.75, name="sens rule top")
    for k, (a, b) in enumerate(rows):
        textbox(s, x, y, w * 0.66, 0.22, [[(a, {})]], size=10, anchor=MSO_ANCHOR.MIDDLE, name=f"sens a{k}")
        textbox(s, x + w * 0.66, y, w * 0.34, 0.22, [[(b, {"b": True})]], size=10, align=PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE, name=f"sens b{k}")
        y += 0.25
        line(s, x, y - 0.02, x + w, y - 0.02, "D1D1D1", 0.5, name=f"sens rule {k}")


commentary(s, 8.60, [
    [("IPCo JV (17% of EV, brand IP) is repaid first incl. its 3.875% premium; $19M to $124M of residual flows to "
      "the estate", {})],
    [("Holdings notes recover ", {}), ("~2x the legacy notes", B), (" through the double-dip (guarantee plus "
                                                                     "intercompany loan claims at Xerox Corp.)", {})]],
    [("The prepack's ", {}), ("52c", R), (" beats the Holdings mid recovery in every case", {})], extra=sens)

# ---------------------------------------------------------------- A6 status quo liquidity
s = new_slide("Appendix: Status Quo Liquidity Build",
              "Without a transaction, liquidity turns negative in 2028E and the funding need reaches $2.4B by 2030E",
              f"{SRC_M} ($XRX Liquidity Analysis); ABL credit agreement (Amendment No. 2); indentures")
xa, yb = img(s, "lq_fcf", L0, IMG_Y, w=6.30, header="Free Cash Flow Build, Base Case ($M)")
img(s, "lq_cash", L0, yb + 0.12, w=6.30)
commentary(s, xa + 0.35, [
    [("FCF after interest and dividends is only ", {}), ("$14M to $78M", B), (" a year from 2027E: cash interest "
                                                                             "of $340M to $363M absorbs it", {})],
    [("The ABL terminates in 2028 (springer), requiring $115M of LC cash collateral and removing ~$250M of "
      "availability", {})],
    [("Maturities of $556M (2028), $1,038M (2029) and $1,050M (2030) cannot be met from cash", {})]],
    [("Funding need of ", {}), ("$387M in 2028E", R), (", rising to $2.4B by 2030E", {})])

# ---------------------------------------------------------------- A7 transaction overview
s = new_slide("Appendix: Transaction Overview",
              "The prepack exchanges $2.3B of unsecured notes for 100% of the equity plus takeback notes, cutting debt "
              "by $2.1B",
              f"{SRC_M} ($XRX Transaction Analysis); IPCo credit agreement and SSLA; terms are preliminary")
xa, ya = img(s, "tx_pf", L0, IMG_Y, w=5.25, header="Pro Forma Capital Structure ($M)")
xm = xa + 0.30
xb, yb = img(s, "tx_eq", xm, IMG_Y, w=6.3, header="Plan Equity Value and Recoveries", hw=R1 - xm)
pts = [[("(1) ", B), ("Debt exchange: ", B), ("$2,306M of notes for $201M of takeback notes cuts debt by ", {}),
        ("$2,105M", B), ("; $121M of fees and consent payments at close", {})],
       [("(2) ", B), ("A $461M exit term loan takes out the IPCo JV; secured debt is extended to 2031 to 2032 for a "
                      "1% fee", {})]]
textbox(s, xm, yb + 0.12, R1 - xm, 0.75, pts, size=10.5, space_after=4, name="tx commentary")
c = box(s, xm, 5.62, R1 - xm, 0.66, fill=GREY_FILL, line=RED, dash=MSO_LINE_DASH_STYLE.ROUND_DOT, line_w=1.0,
        radius=0.08, name="tx callout")
c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(c.text_frame, [[("Case A (out-of-court) leaves ", {}), ("$771M", R),
                        (" of holdouts and legacy notes ahead of the new equity: Holdings recover only 20c", {})]],
        10.5, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- A8 prepack projections
s = new_slide("Appendix: Pro Forma Projections and Credit Metrics",
              "After emergence, Xerox holds $1.1B to $1.4B of liquidity with leverage under 2.0x and no maturities "
              "before 2031",
              f"{SRC_M} ($XRX Transaction Analysis, Case B); base case operating projections")
xe, _ = img(s, "tx_b", L0, IMG_Y, w=9.22, header="Case B: Prepackaged Chapter 11 ($M)")
commentary(s, xe + 0.27, [
    [("Cash interest falls to ", {}), ("~$190M", B), (" a year; EBITDA covers interest 3.5x to 3.8x", {})],
    [("No maturities until 2031; debt grows only through 2L PIK accretion ($494M to $648M)", {})],
    [("FCF after debt service of ", {}), ("$87M to $217M", B), (" a year; net leverage falls from 2.0x to 1.7x", {})]],
    [("Supports plan feasibility: liquidity ", {}), ("never falls below $1.0B", R), (" through 2030E", {})])

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
