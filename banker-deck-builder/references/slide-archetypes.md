# Slide archetypes

Each archetype below was built and approved in the Project Print deck. For each: the story job, the layout, how
to construct it, and the fixes that the render/QA loop found. The full working builder for each is in
`examples/project_print/builders/` (Xerox numbers, but the construction is general); the rendered result is in
`examples/project_print/renders/`. Adapt the content; keep the mechanics.

## Contents
1. Company snapshot
2. Situation overview (2x2 modules)
3. Capital structure (debt table + entity chart + maturities)
4. Business plan (bridge + cases + track record + cash vs. debt service)
5. Valuation summary (football field vs. claims)
6. Trading comps (logo bars + leverage + shortfall table)
7. Precedent transactions (multiples over time + applied range + restructuring precedents)
8. DCF (cash flows + value vs. claims + sensitivity grid + assumptions)
9. Liquidation (book-to-net waterfall + who gets the proceeds + best-interests test)
10. Recovery waterfall (priority bars + recovery vs. market + fulcrum split)
11. Recommendation (before/after + path + why they sign + terms)
12. Appendix model-output page
13. Agenda and cover

---

## 1. Company snapshot  (user-built benchmark; `renders/s-03.png`)
Story job: the business is real and worth preserving. Layout: headline strip of 5 key figures in a dark rounded
bar with white chevrons; business-model flow (4 cards with dark headers and white line icons: what it sells, how
it earns, with FY revenue and %); segment combo charts (bars = revenue, accent line = margin, aligned growth row);
competitive positioning as horizontal bars with **logos as row labels**, peers grey, subject accent in a grey
translucent highlight row with accent dotted border, a margin column on the right with the subject's value in an
accent pill. Lessons: label every figure's period exactly (FY vs. LTM); a bracket over the boxes must span exactly
the boxes its percentage covers; footnote markers must point to the right note; one superscript style.

## 2. Situation overview  (`build_situation.py`, `renders/s-04.png`)
Story job: why the balance sheet broke and when the walls arrive. 2x2 modules, no headline strip:
- M1 "Path to Distress": ink debt columns by year with "$3,277M" labels; accent leverage line on a hidden second
  axis with white-filled markers and labels; an extra hollow accent marker for an adjusted point with a dotted pill;
  deal tags (small dotted rounded boxes: bold title line + detail line) above each bar joined by thin connectors.
  Raise the value-axis max so tags and connectors have room above the bars.
- M2 "Quality of Cash Flow": paired columns (reported grey, core accent) with a brace between each pair labeled
  with the gap ("Receivable runoff $663M"); dotted callout on the right.
- M3 "Liquidity vs. Maturities": a stacked liquidity column (cash ink + availability grey) beside accent maturity
  columns; a dotted pill joined to the bar that triggers a springing maturity; optional dashed liquidity reference
  line across the maturity bars; a one-line ratio note ("maturities total $2,979M, 3.7x current liquidity").
- M4 "Key Issues": three numbered items with bold lead phrases; sub-points with ">" markers.
- Bottom: a thin dashed callout with the CEO framing (exact quote + source, or footnoted as paraphrase).

## 3. Capital structure  (`build_capstruct_v3.py`, `renders/s-05.png`)
Story job: who is owed what, in what order, and where the market says value breaks. Left ~60%: native table
restyled cell by cell (`style_cell`), grouped by payment priority with shaded group rows showing face, market value
and **cumulative leverage at each layer**; columns issuer/guarantor, coupon, maturity, face, price, market value,
YTW; prices below a threshold in accent; a dark total row with weighted-average price; a market-implied TEV row; an
accent dotted box around the fulcrum layers with a "Likely fulcrum" pill; a grey "Payment priority" arrow down the
left. Right top: entity chart from `card`s (parent issuer, borrower/issuer, guarantor / non-guarantor subs) with
the double claim shown as two accent arrows and a pill; a JV beside (not under) the borrower, linked only by
labeled royalty/license arrows; a value-split bar (obligors / JV / non-obligors). Right bottom: maturities by
priority at final maturity with a bracket over the wall years. Bottom callout: the single most original structural
risk (e.g. change-of-control lets a JV terminate the brand license). Lessons: takeaway on observable prices beats
one on a calculated claim; footnote why maturity figures differ between slides (amortization vs. face at maturity);
never carry an unsourced figure from someone else's version.

## 4. Business plan  (`build_bizplan.py`, `renders/s-06.png`)
Story job: the plan supports normalized EBITDA, but not the debt. Four modules:
- M1 EBITDA bridge LTM to normalized: stacked waterfall (invisible base + grey increases + accent decreases + ink
  totals); category labels as two-line textboxes; a dotted growth pill over the end bar; side callout netting the
  ups against the downs.
- M2 cases vs. management/consensus: lines for base / management / downside; single reference points (consensus,
  rating agency) drawn as shapes, not one-point series (they do not render reliably) and explained in the legend.
- M3 assumptions vs. track record: table with the base column outlined; a callout on the guidance history.
- M4 cash generation vs. debt service by year: the coverage gap in accent.
- Closing line: "The business works; the balance sheet does not."

## 5. Valuation summary  (`build_valsum.py`, `renders/s-07.png`)
Story job: value vs. claims, and where it breaks. Left: horizontal football field (stacked bar: invisible base +
range; core methods ink, reference methods grey via dPt; categories reversed so the list reads top-down); row labels
as textboxes (method name bold + multiple range grey); a `band` rectangle behind the chart for the selected range
with a label; dashed claim lines for secured debt (grey) and total debt (accent) labeled at the top; an accent
`hbracket` under the chart between the two claim lines with a pill "Unsecured notes ($X face): the fulcrum".
Right: "Selected Range and Implied Multiples" table (EV low/mid/high, EV/LTM, EV/forward, EV/total debt in accent)
and "Equity Is Out of the Money" bridge table (EV - debt + cash - preferred = negative equity in accent), plus a
callout reconciling market cap with negative value. Fixes found: claim labels collided with the header rule (move
the chart down); low labels of reference bars collided with a claim line (put them inside the bars).

## 6. Trading comps  (`build_comps.py`, `renders/s-08.png`)
Story job: peer multiples are a ceiling, not a target. Left: the snapshot's peer-bar language reused: logo rows
(copied with `copy_picture` from the slide where the user placed them), grey bars sorted descending, a `band` for
the peer Q1 to median range labeled above, the subject in the highlight row with an accent bar at its market-
implied multiple plus a dashed outline segment extending to its face-value multiple ("4.4x at market / 7.1x at
face"), a margin column with the subject's margin in an accent pill and the peer median pill below. Right top:
leverage column chart (peers grey, subject accent) with a dashed peer-median line and an "~4x the peer median"
note. Right bottom: "Peer Multiples Still Leave a Shortfall" table (multiple, implied EV, less debt, shortfall in
accent) and a callout explaining the discount applied. Fixes: exact source values (6.2514 rounds to 6.3x, not
6.2x); white-fill a value label that sits on the median line; shorten the callout to kill an orphan word.

## 7. Precedent transactions  (`build_precedents.py`, `renders/s-09.png`)
Story job: buyers already de-rated the sector. Left: chronological columns split into eras by a dashed divider,
each era with its median as a dotted line drawn BEHIND the bars and labeled in an era header ("2005 to 2019:
growth era, median 7.7x"); early era grey, recent era ink, the subject's own deal accent; an outlier truncated with
a white parallelogram break mark and a footnote; the applied range as a `band` with a legend swatch; category
labels as three lines (target bold / year / acquirer). Value labels inside the bar tops so median lines never cross
them. Right top: the range applied to the subject (Q1, median, recent-era median) with implied EV and shortfall.
Right bottom: comparable restructurings (path, time in court, outcome) feeding a callout tied to the subject's
deadline. Fixes: legend text box with wrap off shifted in render (keep wrap on, widen instead).

## 8. DCF  (`build_dcf.py`, `renders/s-10.png`)
Story job: even crediting the plan, value falls short. Left top: unlevered FCF columns (stub year grey, footnoted)
with the EBITDA line above and a conversion row (FCF / EBITDA %) under the axis; a short side note. Left bottom:
"DCF Value vs. Claims": horizontal stacked bars (PV of FCF ink + PV of terminal value grey, with TV % in the label)
for exit and perpetuity methods, and a claims bar (secured grey_mid + unsecured as an accent dashed outline via a
dPt with noFill); a total-debt claim line; accent shortfall brackets with "($1,180)" labels. Right top: the EV
sensitivity grid from the model drawn as shape cells (base case dark), all cells shaded because all are below the
debt, with a callout giving the break-even exit multiple. Right bottom: key assumptions (value + one-line basis).
The WACC basis must be defensible (cost of equity above pre-tax cost of debt; in restructuring, the reorganized
company's target structure).

## 9. Liquidation  (`build_liquidation.py`, `renders/s-11.png`)
Story job: liquidation is the floor and destroys value, so the plan passes the best-interests test. Left: waterfall
from book assets to net proceeds: book (grey_lt), value lost by asset group (grey_mid) with recovery ranges as a
third label line, gross proceeds (ink), costs, net proceeds (accent), going-concern value (ink) with a dashed link
and a "+$857M preserved" pill. Callout on where the value went. Right top: "Who gets the $X in Chapter 7": a
percentStacked horizontal chart with two rows (proceeds split: admin / secured / unsecured pool; claims sharing the
pool: notes / pension / trade), a funnel of two thin lines from the pool segment to the second bar, and a result
line "every unsecured class recovers ~8c". Right bottom: best-interests table (class, Chapter 7, plan) with the plan
column outlined. Fixes: escape "&" in chart category names ("PP&E" broke the XML); keep funnel labels short so the
diagonal line does not cross them; footnote why secured claims differ from other slides.

## 10. Recovery waterfall  (`build_waterfall.py`, `renders/s-12.png`)
Story job: exactly where value runs out and who holds the fulcrum value. Left: one row per class in priority order;
claim as an outline bar (noFill + grey_mid line) overlapping the recovered amount as a fill (ink = full, accent =
partial) via a clustered bar with overlap 100; columns for claim, recovery $, mid % and low-to-high; an accent
dashed "Value breaks here: $947M left for $3,521M of unsecured claims" line with a white-filled label in a blank
category row. Right top: recovery vs. market price per tranche: mid-recovery bars, low-high whiskers with end caps,
market price diamonds (white on dark bars, outlined on light) with "77c" labels placed beside the diamond. Right
bottom: percentStacked single bar splitting the fulcrum value by holder, plus a callout on the double-dip.

## 11. Recommendation  (`build_rec2.py`, `renders/s-13.png`)
Story job: land the whole argument in one look. Takeaway is the conclusion ("File an RSA-backed prepack that hands
X to the fulcrum notes, halves the debt and clears every maturity to 2031"). Left top: four small before/after
column pairs (status quo accent, proposal ink) for the metrics that defined the problem (total debt, net leverage,
cash interest, maturities) and a one-line liquidity result. Left bottom: path-to-emergence timeline (dots, dates
above, steps below, dashed accent segment to the hard deadline) and a callout pre-answering the obvious objection
(e.g. a rejecting class and cramdown). Right top: "Why [the fulcrum] signs": recovery under each path as columns
(alternatives grey, market grey_mid, proposal ink). Right bottom: terms by class table with the fulcrum rows
highlighted. No banner restating the takeaway; key considerations, alternatives and tax go to the appendix/video.
Closing line wraps the whole deck.

## 12. Appendix model-output page  (`build_appendix.py`, `renders/s-14..21.png`)
Real Excel outputs rendered as images (see appendix-model-outputs.md) on the left with a panel header each, and a
"Commentary" column on the right: 2 to 4 numbered points ("(1) ...") with bold key numbers, and one callout with the
key number at the bottom. Variants: two images stacked; image + image + commentary; image + side image (e.g. WACC
build) instead of commentary. Each page keeps the main-slide title/takeaway/footer system with "Appendix: " titles.

## 13. Agenda and cover
Agenda: two columns of rounded near-black bars with soft shadows, white number circles with accent numbers, italic
grey descriptions, light grey chevrons; no Notes/Sources labels in its footer. Cover: full-bleed image with a dark
panel, title, subtitle, rule, project name, event, author, date; no footer elements at all.
