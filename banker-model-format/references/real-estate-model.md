# Real estate models in the house style

The format carries over unchanged: Garamond, navy bands, grey headers, blue inputs on the input fill, green links,
key totals, comments for sources, checks, End markers. What changes is the tab set, the units and the outputs.
`scripts/example_re_model.py` is a working small version of the acquisition layout below.

## Contents
1. Precedence
2. Tab blueprint (acquisition / value-add; development additions)
3. Units and number formats
4. Tab-by-tab layout notes
5. Cases: sponsor vs. ours, breakevens
6. Checks specific to real estate
7. Questions to settle with the user before building

---

## 1. Precedence

- A fund, sponsor or class template the user must use wins on layout and tab names. Apply this style to it only
  if asked, and never break the template's links.
- If the model feeds the `real-estate-deck` skill (the SMU Cox fund frame), build the outputs that deck expects:
  sponsor vs. our assumptions side by side, one underwritten case plus breakevens instead of bear / bull cases, a
  rent-by-exit-cap grid in 2.5% steps, and an IRR bridge that starts with "model mechanics" (our model run on the
  sponsor's inputs).

## 2. Tab blueprint

Acquisition or value-add (annual, 5 to 10 year hold):

| Section | Tab | Owns |
|---|---|---|
| | Cover Page | entity, data-as-of, sources, case selector, contents |
| Inputs > | Assumptions | purchase, closing costs, growth, vacancy, expenses, capex, financing, exit, hold |
| | Rent Roll (or Unit Mix) | units by type: count, SF, in-place and market rent; loss to lease |
| | T-12 / Historicals | trailing operating statement as reported (blue), normalized lines (black) |
| Analysis > | Operating Pro Forma | GPR to NOI to cash flow before debt service, years across |
| | Renovation / Capex | value-add units per year, cost per unit, premium achieved, lag |
| | Debt | loan sizing (LTV, DSCR, debt yield constraints), schedule, debt service, metrics |
| | Returns | levered and unlevered cash flows, exit, IRR, multiple, profit, sensitivities |
| | Waterfall | LP / GP distribution by hurdle (if there is a promote) |
| Outputs > | Sources and Uses, Summary | one page that only links (the page the deck uses) |

Development adds: Development Budget (land, hard, soft, financing costs, contingency, developer fee; $ total,
per unit, per SF, % of total), Schedule (months to permits, construction start, completion, lease-up,
stabilization), Construction Draw (monthly: cost spend by S-curve or straight line, equity first then loan,
interest reserve, capitalized interest), Lease-Up (units leased per month, concessions), then a permanent loan
take-out on the Debt tab.

## 3. Units and number formats

- Amounts in $ thousands for a single asset (say so on every units line); $ for per-unit and per-SF values. Keep
  one convention per tab and put conversions (`*1000/units`) in the formula, not in a hidden scale.
- Rents: monthly $ per unit (`dol`) or annual $ per SF (`dol2`); say which in the label.
- Rates: cap rates, interest rates, LTV inputs in `pct2` (5.50%); growth, vacancy, margins in `pct` (3.0%).
- DSCR and equity multiple: `mult2` (1.25x); debt yield `pct`.
- Period headers: `"Year "0` for hold years (numbers, so `=prev+1` works), `mmm-yy` for monthly models, with a
  "Hold Period Year" or "Month Ending" band above.
- $ on the first line of each block and on totals, plain numbers in between, negatives in parentheses.

## 4. Tab-by-tab layout notes

**Assumptions.** One-column input blocks (property and purchase, operating, financing, exit), each with a band and
a "Value" header. Derived figures that help the reader (price per unit, going-in cap rate, loan amount, equity
required) sit inside the block as italic or total lines, in black or green. Case-dependent inputs use the side-by-
side layout with an Active column (section 5).

**Rent Roll / Unit Mix.** One row per unit type (or per unit if the rent roll is pasted): units, SF / unit,
in-place rent, market rent, loss to lease, total SF; a weighted-average total row; a check that units and SF tie
to the Assumptions tab. A pasted rent roll stays blue and untouched on its own tab; summaries use SUMIFS on it.

**Operating Pro Forma.** Years across, a "Per Unit, Yr 1" column in D for the main lines.
```
Gross potential rent (bold, $)
  % Growth (italic)
(-) Vacancy and credit loss
(-) Concessions
(+) Other income
Effective Gross Income (total)
(-) Operating expenses by line (payroll, R&M, utilities, insurance, admin, marketing)
(-) Management fee (% of EGI)
(-) Real estate taxes
Total Operating Expenses (total)
  % of EGI (italic)
Net Operating Income (key total, DCEAF7)
  % Growth (italic)
(-) Replacement reserves / capex
Cash Flow Before Debt Service (total)
```
Each growth or ratio driver links to the Assumptions tab (green); Year 1 lines come from the rent roll and the
T-12, with the normalization explained in comments.

**Debt.** Sizing block first (max loan by LTV, by DSCR at the underwriting rate, by debt yield; loan = MIN), then
the schedule: beginning balance, (-) amortization, ending balance (closing balance in the first column links the
loan amount), interest, principal, total debt service (key total). Credit metrics band: DSCR, debt yield, LTV on
a value at the exit cap. For a construction loan, the monthly draw tab owns the balance and this tab links it.

**Returns.** Year 0 to exit across. Equity contribution (Year 0), cash flow before debt service, (-) debt
service, gross sale price (forward NOI / exit cap), (-) cost of sale, (-) loan repayment, Levered Cash Flow (key
total). Then a Returns block (levered IRR bold, equity multiple, profit) and the unlevered equivalents if the deck
needs them. Use IRR on annual flows; XIRR only with real dates (then the header row holds dates, `m/d/yy`).

**Sensitivities.** Build formula grids that recompute only what the axis changes (exit cap varies only the exit
year; rent growth needs its own pro forma rows or a data table). The example's exit-cap grid shows the pattern.
For a rent-by-exit-cap grid, an Excel data table is acceptable in Excel; note that LibreOffice and Google Sheets
will not evaluate it.

**Waterfall.** Rows: distributable cash, then each tier (return of capital, preferred return at x%, catch-up,
split above each IRR hurdle), each with LP and GP columns and a running balance; hurdles computed with a cumulative
IRR or accrued-pref balance per period. Check: LP + GP distributions equal distributable cash every year.

**Sources and Uses / Summary.** Links only: purchase price, closing costs, financing fees, reserves = loan +
equity. Check: sources = uses.

## 5. Cases: sponsor vs. ours, breakevens

- The case selector can hold "Sponsor / Ours" instead of "Downside / Base / Upside". Each case-dependent input is
  a row with one column per case and an Active column (INDEX/MATCH on the Cover Page selection); every formula
  reads only Active. The side-by-side inputs tab doubles as the "key assumptions: sponsor vs. ours" table.
- Breakevens: solve for the exit cap, rent growth or occupancy that returns 1.0x and the fund's target IRR. Show
  the solved values as outputs (computed with Goal Seek by the user, or with a formula where a closed form
  exists), with the method in a comment.
- "Model mechanics" check: run our model on the sponsor's inputs; if the IRR differs from the sponsor's by more
  than a point or two, reconcile the arithmetic before arguing about assumptions.

## 6. Checks specific to real estate

- Units and SF tie between the rent roll and the Assumptions tab.
- Sources = uses; the loan amount equals the binding sizing constraint.
- Ending loan balance at exit equals the payoff on the Returns tab.
- Waterfall distributes 100% of distributable cash; LP + GP = total every period.
- Sensitivity base cell equals the headline IRR.
- Going-in cap rate and price per unit recompute to the purchase price (catches a units or $000s slip).

## 7. Questions to settle with the user before building

1. Asset type and strategy (stabilized acquisition, value-add, development), hold period, annual or monthly.
2. Whose template, if any, and whether the model feeds a deck (which skill, which outputs).
3. Units ($ or $000s) and the rent convention (monthly per unit, annual per SF).
4. Cases: sponsor vs. ours, a single case with breakevens, or downside / base / upside.
5. Financing structure (senior only, mezz / pref equity, construction plus permanent) and whether there is a
   promote waterfall.
6. Source documents available (rent roll date, T-12, OM, term sheet, appraisal) for the data-as-of date and the
   comments.
