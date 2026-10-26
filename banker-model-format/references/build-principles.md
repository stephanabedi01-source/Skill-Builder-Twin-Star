# Build principles

How to keep a model followable, and what the user corrected during the Xerox build.

## Contents
1. Followable by construction
2. Formula conventions
3. Inputs, cases and toggles
4. Notes and sources
5. Checks
6. Consistency across tabs
7. Lessons log (Xerox build)

---

## 1. Followable by construction

- **Flow runs one way.** Inputs, then calculations, then outputs, left to right through the tabs and top to
  bottom within a tab. A block only references blocks above it or tabs before it, except for clearly labelled
  links back to an assumptions tab.
- **One tab owns each number.** Revenue lives on the revenue build; the DCF links it. If two tabs show the same
  number, one of them is a green link.
- **Every block is a band, a header, lines and a total.** A reader scanning bands with Ctrl+arrow on the x markers
  should be able to name every block on the tab.
- **Labels say what and in what unit.** "ABL drawn at liquidation (% of commitment)", not "ABL %". Signs are in
  the label: "(-) Capex", "(+) Depreciation", "Less: Letters of Credit".
- **No orphan or decorative rows.** No summary that repeats the tab, no "key outputs" box that re-links numbers
  already on the page, no commentary cells. If a number is not used by another tab, the deck or a check, ask
  whether it needs to be there.
- **Simple beats clever.** The user's test: someone opening the tab should know what is going on and not feel
  anything was overcomplicated. One clear roll-forward beats a dense grid that mixes timing, sources and uses.

## 2. Formula conventions

- No typed constants inside formulas. `=D11*1.05` hides an assumption: put 5.0% in a blue input cell and
  reference it. Structural constants (12 months, 100 for prices, 365 days, 1,000 for unit conversion) are fine.
- Typed historical arithmetic (`=1053+55` summing reported segments) is a hard-code: blue, with the source in a
  comment.
- Same formula across a row. If one column differs (a stub period, the first projection year), comment why.
- Prefer SUM, SUMIFS, INDEX/MATCH, MIN/MAX. Avoid OFFSET and INDIRECT (volatile, hard to audit), nested IFs
  deeper than three, and array formulas unless the user's model already uses them.
- IFERROR only where an error is expected (an empty peer row), never to hide a broken link.
- Signs: costs and outflows negative in the model, shown in parentheses; totals are plain SUMs.

## 3. Inputs, cases and toggles

- Inputs: blue on the input fill, one per cell, grouped by what they drive, each with a comment.
- Cases: a single selector on the Cover Page (list validation: Bear / Base / Bull or Downside / Base / Upside).
  Each tab shows it in a case row (row 6). Case-specific inputs sit side by side in columns with an "Active"
  column that picks one with INDEX/MATCH; every formula reads only the Active column.
- On/off switches (include a contested claim, terminate the pension) are 1/0 inputs formatted
  `[=1]"On";[=0]"Off";General` or as Yes/No with a comment explaining both states. Sensitivities that flip a
  switch are recalculated on a copy, never by editing the user's file.

## 4. Notes and sources

- Every hard-code gets a comment: source document and date (e.g. "Xerox Q2'26 10-Q, p. 14; 6/30/26"), or for an
  assumption the reasoning and the range considered.
- Comments are per cell and specific. Check them after copying rows: the Xerox precedents tab carried a comment
  copied from another deal until it was caught.
- Units on the tab's units line must be true for every number on it. The comps tab said "$USD" while Japanese
  peers were in yen; the fix was a correct units line plus a converted USD output block with the FX rate
  commented.
- The Cover Page lists the sources and the data-as-of date.

## 5. Checks

- Add a check wherever two numbers must agree: balance sheet balances, sources equal uses, book values tie to the
  balance sheet, units tie to the rent roll, the waterfall distributes exactly the distributable value, a
  sensitivity's base cell equals the headline output.
- Check lines are italic, indented, labelled "Check: ...", and show an en dash when they pass; conditional
  formatting turns them red otherwise (`Tab.check`). Never leave a failing check in a delivered model.

## 6. Consistency across tabs

Before delivery, trace the headline numbers through every tab that uses them: the same EBITDA basis, cash, debt
and pension figures everywhere, the same valuation date, the same claims date. The Xerox v17 consistency pass
found cash and pension treated differently between the waterfall and the transaction tab; the fix was one basis
and links, not two calculations.

## 7. Lessons log (Xerox build)

| What happened | Rule it became |
|---|---|
| DCF free cash flow build had its own summary block at the bottom | Link the FCF build to the operating assumptions; delete summary blocks |
| Liquidity analysis tab was hard to follow and did not look like the user's debt schedules | Rebuild as a simple roll-forward in the user's schedule layout |
| Liquidation, normalized EBITDA and football field tabs had a notes column (K) | Notes and sources go in cell comments |
| Stock performance compared Xerox with IGV, a software ETF | Every comparator must be relevant; cut what is not |
| Football field and recovery waterfall were dense | Simplify until a reader can follow it without help |
| Formatting drifted from the user's tabs in places | Copy styles from the user's best tab; audit before delivery |
| Cash and pension differed between the waterfall and the transaction case | One basis, linked everywhere; consistency pass before delivery |
| Comps tab was labelled USD but held yen | Units lines must be true; convert explicitly with a commented FX rate |
| A deal's comment was copied from another deal | Re-read comments after copying rows |
| Model outputs pasted into the deck were not in Garamond | The model's font carries into every rendered output |
