# Classification heuristics

Everything here serves **impose mode**. If the workbook already has an
established format, stop — clone it instead (`clone_format.py`; see Mode 0 in
SKILL.md). Re-deriving an existing format from these heuristics is the failure
mode, not the goal.

**Generalization note:** the keyword lists, tier examples, tab-name hints, and
any concrete value below are *evidence from specific models*, not universal
constants — see the "Generalization first" section in SKILL.md. They describe the
*kind* of signal to look for; the actual values (palette, fonts, freeze panes,
formats, names) come from whatever model is in front of you. Classify against the
model's own structure and reproduce the model's own formatting — never substitute
Twin Star's values for what the target actually uses.

How to recognize what you're looking at in an unfamiliar model, so the right
convention from `design-language.md` / `tab-patterns.md` gets applied. `audit_workbook.py`
makes a first pass with these heuristics and attaches a confidence to every guess;
**your job is to review and correct that map** — especially anything it flagged as
low-confidence — before running `apply_format.py`. Sample real cells; don't trust a
label keyword blindly.

The split is deliberate: **classification needs judgment (yours); application is
deterministic code.** Get the classification right and the formatting follows
exactly, because the script never paraphrases a hex code or a format string.

---

## Cell provenance (mechanical — the script derives this live)

**Master switch first — is this a live model tab or a presentation/output page?**
A sheet that mostly *pulls* other sheets (>~55% cross-sheet links — "IS Out", "O1",
"LFCF Build", valuation summaries) is an **output page → format everything gray
`525766`, no provenance colors** (`tab_type: output`). The provenance table below
applies only on **live model tabs**. (Across 5 keys: output-page cells are ~75-80%
gray; model-tab hardcodes ~44% blue.) This is the rule earlier versions missed.

| Cell content (on a model tab) | Provenance | Font color |
|---|---|---|
| Constant (number / date / text typed in) | hardcode | blue `0000FF` (assumptions/inputs); loaded historical data may be blue or left gray |
| Formula with no other-sheet reference | calculation | gray `525766` |
| Formula referencing another sheet (`=Sheet!A1`, `='Other Tab'!$B$2`) | cross-sheet link | green `008000`/`00B050` (a convention applied ~half the time; gray is acceptable) |

You don't tag colors per cell — the script reads each cell and colors it. What you
*do* decide is the master switch (model vs output tab) and the semantic layer below.
See `house-conventions.md` for the full purpose→format rule set and what's
invariant vs up-for-interpretation.

## Row class

Work from the row's label text **and** its cells. Keyword lists are starting points,
not proof — confirm against the cells.

| Signal | Class | Notes |
|---|---|---|
| Label has **Total / Net / Gross / Subtotal / EBITDA / "net change" / "availability"** *and* the cells are sums | total → assign a **tier** (see below) | the keyword alone isn't enough; "Net Sales" as a plain loaded line isn't a total |
| Label has **Check / Tie / Tie-out** | check row | italic red, 3-decimal format, to expose tiny breaks |
| Label has **% / Margin / Growth / "% of"**, or the cells are percent-formatted, or values sit under a parent as a ratio | percent row | italic gray, directly under its parent |
| Label row with no numeric data, **"X" marker** in col B, heads a section | banner | navy fill spanning label→last data col |
| Label row, no data, ends with **":"** ("...Activities:") | sub-section label | bold `0067A5`, no fill |
| Label row, no data, e.g. BS **"Assets" / "Liabilities & Equity"** | section underline | bold + underlined gray text, **not** a banner fill |
| Mostly constants across the period | loaded-data row | blue font; first such row in a block gets the `$` format |
| Mostly formulas | calculation row | gray font; subsequent-row number format |
| "#'s in thousands" / "$ in millions", once near the top | unit note | gray-purple italic |
| Bare 4-digit years / chained dates / =YEAR() / "Actual"–"Forecast" near the top | header band | see chassis in `design-language.md` |

The "first `$` row of a block" rule: within a section, only the first data row gets
the `$` number format; later rows drop the `$`. A block resets after any banner,
sub-section header, or total.

**Percent FORMAT is not a percent (margin) ROW.** The italic-gray `percent` class is
for *derived* margin / growth / % rows (formulas under a parent). A **hardcoded**
percentage — an assumption like a hedge %, discount rate, tax rate — is a blue
**input** (often `0.00%`), frequently in a yellow input block; not italic, not gray.
Classify by provenance first: a constant `%` cell is an input, not a margin.

**Column / table sub-headers are not section headers.** A row of column captions
("Description", "Face", "Rate", "Maturity", "Year Ended") is **gray bold**, not a
navy `section_header`. Reserve section styling for the outline labels heading a group.

**Number format is per data-type, often per COLUMN.** On cap tables and assumption
blocks the columns hold different units side by side ($ amount, rate %, multiple x,
maturity date). Don't apply one row format across them — supply `column_numfmt` so
each column gets the format its data calls for. Also: the unit note "($ millions)"
is plain **gray** italic (purple is Twin-Star-specific), and "DRAFT"/"CONFIDENTIAL"
header stamps are **red** on any tab.

## Tier assignment (the judgment call the script is weakest at — review it)

Three fills, three ranks. The principle: **a total of totals outranks its
components, and the statement's bottom line / balancing figure is the headline.**

| Tier | Fill | Gets it |
|---|---|---|
| 1 — subtotal | `F2F2F2` | the first level of summation (Total Current Assets, Total Product Cost, a per-section net-cash subtotal) |
| 2 — major total | `DCE8F4` | a total that sums subtotals (Total Cost of Sales, EBITDA, Total Liabilities, net change in cash) |
| 3 — headline total | `D3EFEF` | the bottom line / balancing figure (Gross Profit, Total Assets, Total Liabilities & Equity, Cash at end of period, Total Availability) |

Reason about the arithmetic, not just the words. If "Total X" is computed by adding
two rows that are themselves "Total …" rows, it's at least tier 2. The figure the
whole statement drives toward (or that must balance) is tier 3. Two totals with the
same word "Total" can land in different tiers depending on what they sum.

## Tab type (name keywords + structure)

Confirm with structure, not just the name — a tab named "Summary" could be an output
page or a working scratchpad.

| Tab type | Name hints | Structural tells |
|---|---|---|
| divider | name ends with **`>`** | empty tab |
| cover | "Cover", "Title", "Disclaimer" | one text block, confidentiality language |
| assumptions | "Assumptions", "Drivers", "Inputs" | stacked input blocks of hardcodes |
| income statement | "Income Statement", "P&L", "Statement of Operations" | revenue → margins → EBITDA |
| balance sheet | "Balance Sheet", "Financial Position" | Assets / Liabilities & Equity, a check row |
| cash flow | "Cash Flow" | operating / investing / financing sections |
| liquidity | "Liquidity", "Borrowing Base", "Availability" | advance rates, greyed closed columns |
| forecast | "Forecast", "COGS", "Sales", "Build" | loaded detail rows + group subtotals |
| output | "Out", "Output", "CIM", "UFCF" | **Segoe UI font**, annual-only, all gray |
| working | "Bridge", "Scratch", "Raw", "Tie-out", "Staging" | gridlines on, raw constants, often already a green tab |
| source | data dumps | huge raw grids, often already a light-blue tab |

**Strong structural signals override the name:** a Segoe-UI tab is an output page; a
tab the modeler already colored working-green / source-blue is working / source; a
tab with gridlines left on and no banners is probably working.

### The fork worth asking about
Whether a borderline tab (a "Bridge", a "Summary", a half-built schedule) is a
**presentation tab to format** or a **working tab to leave alone** is the one call
genuinely worth a question to the user — formatting a working tab, or skipping a
real schedule, are both clearly wrong, and the structure alone may not settle it.

## Adapting to a different layout

- **Labels already in col A** → adapt to the existing skeleton; don't insert columns
  unless the user approves restructuring.
- **Annual-only model** → one annual block; no monthly axis, no "x" separator, no
  SUMIFS annuals.
- **Quarterly model** → same chassis, quarterly EOMONTH stepping.
- **No Assumptions tab** → skip toggles and green Assumptions links entirely.
- **Target lacks a feature** → skip that convention. Apply everything proportionally
  to the structure the model actually has.
