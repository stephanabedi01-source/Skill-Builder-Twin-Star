# Mode B — Format from scratch

When the source is unformatted, stripped, or a chaotic mess, there is no key to
copy. **Your judgment is the standard.** The deliverable is a clean,
professional, polished model — the kind a careful analyst would be proud to send
out. This file is the bar to hit and the workflow to hit it with.

The iron rule still holds in this mode: **formatting only — never change values,
formulas, structure (rows/columns/tabs), or defined names.**

## When Mode B applies

Read the workbook first. Signals that the format is absent or not deliberate:
default Calibri 11 everywhere, gridlines on across the board, no fills or only
random ones, General number formats on money, no tab colors, inconsistent fonts
from copy-paste, mixed number formats on the same kind of row. A consistent,
intentional format — even an ugly one — means Mode A (replicate) unless the user
asks for a restyle. Sloppy *and* inconsistent means Mode B. If genuinely unsure
whether an odd format is intentional, ask.

## What "excellent" means (the judgment standard)

Use the house design language (`design-language.md`) as your default vocabulary —
it is a complete professional system. **But first read
`references/house-conventions.md`**: the values in design-language are Twin Star's,
and real models in this style are often more restrained or use a different font /
tab-color scheme. Set a `palette` and per-sheet `font_family`/`font_size` in the
map to match the model (see SKILL.md "Mode B palette is overridable"). Where the
model's content calls for something the language doesn't cover, decide like an
analyst would; **consistency and restraint are the bar** — when unsure, prefer the
quieter choice (light-gray total fills + bold + a hairline rule over loud fills).
Concretely:

- **Coherent font system.** One body font/size everywhere on model tabs (house
  default Arial 10, body color slate gray `525766`, never pure black); at most a
  second font family for presentation/output pages. No font roulette.
- **Color coding = provenance.** Blue hardcodes/inputs, gray formulas, green
  cross-sheet links; red strictly for checks/flags/warnings. Inputs triple-marked
  (yellow fill + blue font + one blue box per input block).
- **Number formats.** Negatives in parentheses; zeros as dashes; `$` on the first
  row of a block and on totals only; percents at one decimal; month headers as
  mmm-yy. State units once per tab ("$ in thousands") and scale consistently —
  pick the unit that keeps body numbers 1–6 digits; never mix scales in a block
  without saying so. Exact strings from `design-language.md` §5 — never
  paraphrase.
- **Clear section headers.** Dark banner rows spanning label → last data column,
  white bold label; bold colored sub-section labels; bold+underlined gray text
  for statement section heads (Assets / Liabilities & Equity).
- **Distinct subtotal vs total treatments.** The three-tier fill hierarchy with
  thin gray rules above and below; bold on all tiers. A reader must be able to
  see the rollup at a glance.
- **Alignment & indent hierarchy.** Labels left-aligned with real indent levels
  (`indent` 1/2/3 on line items under their headers — set the per-row `indent`
  field in the map); period headers centered; numbers aligned by their formats.
- **Sensible geometry.** Chassis widths via `auto_widths` (narrow margin/marker ≈
  1.8, label ≈ 26, data ≈ 10 — widen if content needs it); default row heights;
  small spacer rows where the model already has them. Never let labels truncate
  into data columns on a tab you're calling finished.
- **Useful freeze panes — sparingly.** Freeze the main multi-screen statement at
  its header-rows × label-columns intersection. Do NOT freeze short output/exhibit
  tabs or every schedule; real models freeze only where scrolling loses context.
- **Tab colors that organize.** One color family per workbook section, as theme
  tints of the firm palette; dividers/cover uncolored. The hue→section mapping is a
  per-model choice (see house-conventions.md) — pick a coherent scheme; the tab
  strip should read as a table of contents.
- **Gridlines OFF on every client-facing tab — including data/exhibit tabs.** Set
  `gridlines_off: true` even on body-skipped data tabs (apply_format honors it
  without restyling the body). Reserve gridlines-on only for genuinely internal
  scratch the client never sees.
- **Clean print setup.** Portrait Letter, ~0.7"/0.75" margins, fit one page wide,
  gridlines off.

## Workflow

1. Read `design-language.md` (vocabulary) and `classification.md` (how to read
   the model). Run `audit_workbook.py`; review and correct the map exactly as in
   the impose workflow — Mode B *is* that workflow with the judgment dial turned
   up.
2. While reviewing, set the Mode-B fields the scripts honor: per-row `indent`
   for the line-item hierarchy; per-sheet `auto_widths: true` (plus any explicit
   `column_widths` / `row_heights` overrides); confirm `freeze`, `tab_color`,
   `numfmt_family`, and input blocks.
3. `apply_format.py` → then **look at the result** (re-audit, sample cells,
   render a tab to image if available). Fix the map and re-run until it reads
   clean. One pass is rarely excellent.
4. **If the source is rich** (charts, data tables, cached values, defined names),
   treat the `apply_format.py` output as a *style donor* and transplant it onto
   the original package with `merge_format.py` — openpyxl's round-trip would
   otherwise strip cached values and degrade charts. A plain workbook can skip
   this. Verify zero value/formula change (compare against the source, comparing
   `ArrayFormula`/`DataTableFormula` by attributes and floats by tolerance) and
   that charts / defined names / cached values survived.
5. Finish with the delivery standard: `finalize_delivery.py` (first tab active,
   A1 everywhere, 85% zoom). Every workbook handed back gets this, no exceptions.

## What Mode B must never do

Copy nothing blindly from any previous model's *content-specific* choices (its
sheet names, its row numbers, its division names). And never let "no key exists"
license value edits: if a number looks wrong, flag it in your report — formatting
only.
