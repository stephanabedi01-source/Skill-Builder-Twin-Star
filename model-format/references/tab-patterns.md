# Tab-type playbooks

How each kind of tab is formatted. Recognize the tab type (see
`classification.md`), then apply its playbook. All hex codes, fonts, and number
formats come from `design-language.md` / `style_constants.py` — this file is about
*which conventions go where*.

The golden rule that overrides everything: **never invent structure the target
doesn't have.** No Assumptions tab → skip toggles and green links. Annual-only
model → single annual block, no monthly axis, no "x" separator. When the target
lacks a feature, skip that convention.

---

## Working tabs — leave them UNFORMATTED (the most important rule here)

Bridges, scratch dumps, raw source-paste tabs, "by-X" breakouts, reconciliation
tabs, anything an outsider will never see: **do not format the body.** Default
font, gridlines stay ON, no fills, no number-format changes. The script skips any
tab whose map says `format_body: false`. A section **tab color** is still applied
(working green `00B050`, source blue `00B0F0`), because the tab strip is part of
navigation — but the cells are left exactly as they are.

Why: formatting effort is a signal. Reserve it for presentation surfaces.
Over-formatting a scratch tab makes it look load-bearing when it isn't, and risks
disturbing fragile paste ranges. When in doubt whether a tab is working or
presentation (a classic "Bridge" tab could be either), **ask the user** rather
than guess.

## Cover

Logo top-center if present; large text box with confidentiality / disclaimer
language; right-aligned **bold red** warnings ("DO NOT DISTRIBUTE", "CONFIDENTIAL
/ PRIVILEGED", "SUBJECT TO MATERIAL REVISION"). Portrait, ~50% print scale.
Gridlines off. Body font slate gray. The script red-bolds any cell whose text
matches confidentiality/disclaimer keywords.

## Assumptions

A vertical stack of small input blocks. Per block:
- Bold `0067A5` **sub-section label row**.
- One row per division/category: label left in the label column, inputs in a
  **column-per-forecast-year rectangle** — blue font, yellow `FFFFCC` fill, `%`
  format, a single blue box border around the whole block (one box).
- Navy `002855` banners split major groups ("IS Items", "BS Items", "CFS Items");
  steel-blue `9FC3DA` banners split subgroups.
- A first forecast-year column may be `F2F2F2`-filled where it is an LTM-derived
  anchor rather than a pure input.
- Red bold notes flag open items.

This is the one tab where hardcoded numbers **are** inputs (→ triple-marked). On
statements, hardcoded numbers are *loaded data* (blue font only, no yellow).

## Income Statement

Repeating divisional blocks — Consolidated first, then one identical block per
division, each headed by a navy banner + red "X" in col B. Loaded monthly data
rows blue; derived rows (margins, GP, EBITDA) gray formulas. **Margin % rows
italic gray directly under their parent.** Tiers: subtotals `F2F2F2`; EBITDA /
Total Cost of Sales `DBE8F4` (tier 2); Gross Profit `CDF5F5` (tier 3). Below
EBITDA, other income/expense items are tagged "Cash" / "Non-Cash" in col B (these
feed output-tab SUMIFS — leave the tags alone). Check rows italic red, 3 decimals.

## Balance Sheet

Same chassis. "Assets" / "Liabilities & Equity" headers are **bold + underlined
gray text — NOT banner fills** (use the `section_underline` class). Tiers: Total
Current Assets / Liabilities `F2F2F2` (tier 1); Total Liabilities `DBE8F4`
(tier 2); Total Assets and Total L&E `CDF5F5` (tier 3). A "Check" row (Assets minus
L&E) italic red. Subsequent data rows use the `-_);` number-format variant.

## Cash Flow

GAAP indirect layout. Section headers bold `0067A5` ("Cash Flows from Operating
Activities:"). Net-cash-per-section subtotals `F2F2F2` (tier 1); net change in cash
`DBE8F4` (tier 2); ending cash `CDF5F5` (tier 3). "Memo:" rows plain gray.
Subsequent data rows use the `-_);` variant.

## Liquidity / borrowing base

Teal header family (`24B1B1` / `9CEBEB`). Advance-rate and rate-spread inputs blue
font in the live columns; closed historical month columns greyed `F2F2F2` down the
body. Availability subtotals follow the 3-tier hierarchy. Cell notes/comments are
used for documentation — leave them.

## Sales / COGS / detail forecasts

Gray header family (`BCBFC6` / `E3E4E8`). Loaded detail rows greyed `F2F2F2`; bold
`0067A5` group headers; a bold `F2F2F2` subtotal per group. Col B may repeat a
normalized lookup key next to the display name in C — leave it.

## Output tabs (presentation pages)

A different aesthetic:
- **Segoe UI 10**, annual columns only.
- **Everything gray** — no blue / green / yellow. These are presentation pages, not
  working models, so provenance coloring is dropped; all text is slate gray.
- Expenses shown as negatives via patterns like
  `=-SUMIFS('Income Statement'!38:38,'Income Statement'!$6:$6,E$7)/1000` (converted
  to $mm). Leave the formulas; just style.
- Header: "Year Ended" over year numbers, with medium `BFBFBF` / `BCBFC6` double-
  rule lines above and below; thick white borders as column separators.
- Same 3-tier total fills and thin-gray total borders as everywhere else.
- Italic red check row at the bottom tying back to the source statement's annual
  column.

> Because output tabs drop provenance coloring, treat their data rows as plain gray
> regardless of whether a cell is a hardcode or a link.

## Carveout / special analyses

Teal tab color (`24B1B1`). Otherwise follows the statement playbooks (a carveout
income statement is formatted like the Income Statement above).
