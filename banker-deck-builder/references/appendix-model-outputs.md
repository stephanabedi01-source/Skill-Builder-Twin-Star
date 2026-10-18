# Appendix pages from real model outputs

The user's rule: "where you can, just put the Excel outputs in, because if you don't it looks very Claude." The
winning deck from the prior year did the same (pure model-output pages). The appendix is backup a judge flips to in
order to check a number, so it should look like the model, in the model's own formatting.

## Choosing the pages (8-page budget in the case competition; adapt)
Pick pages that back the main slides' numbers and pre-answer the hard questions:
1. Revenue build (by segment, with pro forma bridge for acquisitions)
2. Operating assumptions by case (margins, EBITDA, restructuring costs, tax)
3. DCF detail: FCF build, terminal value and EV, WACC build
4. Liquidation detail: proceeds by asset, costs, distribution to claims
5. Recovery waterfall detail by tranche (low/mid/high) + sensitivity to contested claims
6. Status quo liquidity build (why the deadline is real)
7. Transaction: pro forma capital structure, plan equity value and recoveries, exchange summary
8. Pro forma projections and credit metrics (feasibility)
Leave out pages whose model ranges contradict the deck (e.g. a comps tab still on a different EBITDA basis) until
the model is fixed; never paste a number that disagrees with a main slide.

## Rendering the ranges
Use `scripts/render_xl.py jobs.json`. Requirements and lessons:
- Render from a recalculated copy (cached values present). If FactSet/Bloomberg formulas cannot evaluate in the
  sandbox, use a verification copy with those cells replaced by the user's cached values.
- Choose ranges that start at the section band (not the sheet title), exclude marker columns, and exclude rows that
  are template leftovers or contradict the deck. Split tall tabs into two ranges (e.g. proceeds + distribution) so
  each image stays readable.
- Excel data tables (What-If grids): LibreOffice does not evaluate them; compute the anchor cells and pass them as
  overrides, then check the corner value visually.
- Typical sizes: one wide image at ~9.2 in wide; two stacked images at ~5 to 6.3 in wide; side-by-side images at
  ~4.5 to 5.3 in. Text in the image should land at roughly 7 to 9 pt; if it is smaller, split the range or drop the
  commentary column for that page.

## Page layout
- Same title/takeaway/footer system as the main slides; title "Appendix: <Topic>", takeaway is the claim the
  backup supports.
- Panel header over each image ("Revenue Build, Base Case ($M)").
- Commentary column on the right (x ~9.7 to 13.1): header "Commentary", 2 to 4 numbered points "(1) ..." with the
  key numbers in bold, and one callout at the bottom with the single most important number.
- A small native table is fine in the commentary column when it adds information the model does not show in one
  place (e.g. a four-case sensitivity recalculated with model toggles).
- Footer: "Source: <Project> model (<tab name>); <filings>".

## Sensitivities recalculated from the model
When a page shows cases the model does not tabulate (e.g. claims switched on/off), recalculate on copies of the
model with the toggles changed, read the results, and never modify the user's model file for it. Say in the footer
that the cases were recalculated with the toggles off.
