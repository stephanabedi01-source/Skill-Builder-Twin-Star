# Tab anatomy

Row-by-row layouts. Column letters assume the standard grid: A and B markers, C labels, data from D. Shift the data
columns, never the top rows.

## Contents
1. Standard top of tab
2. Time series (years across)
3. Matrix (Low / Mid / High or cases across)
4. Assumption block (one input column)
5. Schedule by year (the user's debt schedule layout)
6. Capital structure table
7. Sensitivity grid
8. Cover Page
9. Divider tabs and tab order

---

## 1. Standard top of tab

```
row 1   A1 =C3 (hidden ;;;, grey)   B1 ='Cover Page'!A1  (link, shows "Return to Cover")
row 3   C3 Tab Title (bold) ............................................ Entity (bold, right) [thin rule under C3:last]
row 4   C4 All values in $ millions (USD)        (italic)
row 6   C6 Active Case (select on Cover Page)    D6 ='Cover Page'!$F$21   (green bold, input fill, boxed)
row 7/8 first band
```
Title wording: "$XRX, Liquidation Analysis (Chapter 7, Illustrative)" style: what the tab is, then the basis.
The A1 echo is what the Cover Page contents links display, so each link shows the tab's real title.

## 2. Time series (years across)

```
row 8   x x | Income Statement (navy, white bold) | Fiscal Year Ended December 31, (italic, centered) | Projected (215F9A)
row 9       |                                     | 2023A  2024A  2025A | LTM (4E95D9) | 2026E  2027E ...   (grey header)
row 10      (3pt spacer)
row 11      Revenue (bold)                          $x,xxx ...                         (blue history, black/green projection)
row 12        % Growth (italic, indent 2)             x.x% ...                           (projection drivers: input fill)
row 13      Cost of Revenues (indent 1)             (x,xxx) ...
...
row 23      Adjusted Operating Income (bold, top rule)   $xxx ...
```
- Years are numbers: type the first, then `=prev+1`, formatted `General"A"` / `0"E"`.
- Projection drivers (growth, margin) sit on the % line under the line they drive, on the input fill, so the
  reader sees the driver beside its result.
- If one tab feeds another (Operating Assumptions -> DCF), the receiving tab links every line (green); it never
  re-derives.

## 3. Matrix (Low / Mid / High or cases across)

From the liquidation analysis the user liked:
```
row 8   x x | Gross Liquidation Proceeds | Book Value | Low (215F9A, spans 2) | Mid (spans 2) | High (spans 2)
row 9       |                            | 6/30/26    | % Rec.  Proceeds       | % Rec.  Proceeds | % Rec.  Proceeds
row 11      Cash and cash equivalents      x,xxx (green link)  100.0% (input)  x,xxx (calc) ...
...
row 23      Total Gross Liquidation Proceeds (bold, top rule)  $x,xxx  xx.x%  $x,xxx ...
row 24        Check: Book Value vs. Balance Sheet (italic check)   –
```
- Assumptions (% recovery) sit next to the amount they drive, on the input fill, one per cell.
- Blocks chain down the tab: proceeds, costs, net proceeds, distribution by priority. Each block starts with a
  band and ends on a total; the next block starts by linking that total.
- Derived lines that only inform ("Implied multiple of LTM EBITDA", "Value remaining for unsecured") are italic,
  indent 2, no total rule.

## 4. Assumption block (one input column)

```
row 38  x x | Liquidation Claim Assumption | Input (italic)
row 39      |                              | Value (grey header)
row 41      ABL drawn at liquidation (% of commitment)   60% (blue, input fill)
row 42      Accrued & unpaid interest (months)            3
```
Each assumption: a plain-English label that says the unit, one input cell, a comment with the source or the
reasoning. Group assumptions by the block they drive, on the tab they drive, unless several tabs share them (then
one Assumptions tab).

## 5. Schedule by year (the user's debt schedule layout)

From the user's Maturity Schedule tab:
```
row 7   x x | Debt Maturity Profile (band across C:J)
row 8       |            | Liquidity  2028  2029  2030  2031  2035  2039   (grey header, thin bottom)
row 9       Cash & Equivalents          $xxx        (green link, $ on first row)
row 10      ABL Availability             xxx
row 11      Secured Term Loan                    –    xxx   ...   (SUMIFS on the capital structure tab, green)
...
row 16      Total (bold, top rule)      $xxx  $xxx  $x,xxx ...
row 35  x x
row 36  End
```
- One row per instrument, one column per period; amounts pulled from the capital structure tab with SUMIFS on
  maturity year, never retyped.
- Amortizing loans (real estate, term loans): beginning balance, (-) amortization, (+) draws, ending balance (total
  rule); then interest and total debt service (key total). Closing balance in the first column links the loan
  amount. Credit metrics (DSCR, debt yield, LTV) in their own band below.
- The user rejected a bespoke liquidity build that mixed sources, uses and timing on one grid: keep schedules
  as simple roll-forwards.

## 6. Capital structure table

The user's Capital Structure tab: two header rows (`Face / Value`, `Leverage / (Face)`, `Interest / Rate`, ...),
3pt spacer, one row per tranche (label indented, face in `dol`, cumulative leverage `mult`, coupon `pct3`,
maturity `mmm-yy`, price `num1`, YTW `pct`), key-total rows for Total Secured Debt, Total Debt, Net Debt and Net
Debt & Preferred (bold, `DCEAF7`, thin top and bottom), a grey memo row for the EBITDA that leverage is measured
on, and a liquidity walk block (commitments, less borrowing-base limit, less drawn, less LCs = availability; plus
cash = total liquidity).

## 7. Sensitivity grid

- Header row on the `ASIDE` fill (`A6CAEC`) or a navy band, the axis values blue (they are inputs), the base case
  cell highlighted (bold, or the darker grey used on the deck).
- Prefer formula-built grids that recompute only what the axis changes (see the Returns tab in
  `scripts/example_re_model.py`): they recalculate in Excel, LibreOffice and Google Sheets. An Excel data table
  (What-If) is fine in Excel, but LibreOffice will not evaluate it: compute the anchor values separately when
  rendering.
- Add a check that the base-case cell equals the headline output.

## 8. Cover Page

```
row 1   A1 Return to Cover (grey; the text every B1 link displays)
row 2   C2 Entity name (bold, centered across C:J, thin top rule)
row 3   C3 All $ in Millions (USD), except per-share values (italic)
row 8   C8 Model subtitle (12 italic grey), e.g. "Credit & Restructuring Analysis"
row 10  Last Updated  =TODAY()               |  Data As Of  2Q'26 (June 30, 2026)  (blue, input fill)
row 12  Color Coding                          |  Links (Ctrl + [ to Jump to Sheet)
row 13  Formulas / Hard-Code / Link / Check / |  Section heading (bold underlined)
        Heading 1 / Heading 2 / Input-Driver  |    ='Tab'!A1 (indent 1, hyperlinked), one per tab
row 21  Active Case  Base (blue, input fill, list validation)
row 24  Sources: one line per source (filings, data providers, rating agencies, broker / lender documents)
```

## 9. Divider tabs and tab order

- Order: Cover Page, then sections each opened by a divider tab named `Section Name >`: inputs, outputs, credit /
  analysis, valuation. The Xerox order: Cover Page | Financial Statement Inputs > | IS, BS, CFS | Financial
  Statement Outputs > | outputs | Credit Analysis > | NWC, Capital Structure, Maturity Schedule | Analysis > |
  precedents, comps, revenue build, operating assumptions, DCF, liquidity, liquidation | Stock Performance > |
  Valuation Football Field > | valuation, waterfall, transaction | Charts.
- Divider tab: A1 the tab name in grey bold, B1 the link home, zoom 60, nothing else.
- Analysis tabs specific to the company carry a ticker prefix (`$XRX DCF`); generic tabs do not.
