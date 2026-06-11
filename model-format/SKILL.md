---
name: model-format
description: Format Excel financial and operating models. Two modes — losslessly clone or restore a workbook's existing formatting (exact style records, theme, number formats, views, charts, drawings, cached values) when an established format exists, or impose the house style on unformatted models (banker color-coding with blue inputs, slate-gray formulas, green cross-sheet links; three-tier total fill hierarchy; navy section banners; triple-marked yellow input blocks; exact number formats; model chassis). Use whenever the user uploads or points to an .xlsx and asks to format it, apply the house style or "our model format", make it look like our models or like the original, restore or preserve formatting, transplant formatting between versions, color-code inputs vs formulas, fix totals or number formatting, or make a model client-ready. Triggers on "format this model", "apply house style", "match the original formatting", "style this workbook". Works on any model — monthly or annual, one division or ten.
---

# Model Format

Apply a specific investment-banking house style to any Excel model. The style is a
**design language, not a literal template**: the conventions below were distilled
from one model, but this skill formats *any* model — different tab names, different
row/column counts, annual instead of monthly, one division or ten. Your job is to
(a) **classify** what you're looking at — what kind of tab, row, and cell — then
(b) apply the matching convention **exactly**.

This works in two layers, and the split is the whole point:

- **Semantic classification — your judgment.** Recognize subtotal vs. major total
  vs. headline total; input vs. formula vs. cross-sheet link; section banner vs.
  sub-section label; margin/% rows; check rows; working tab vs. presentation tab vs.
  output tab; monthly vs. annual; actual vs. forecast.
- **Deterministic application — Python.** Once classified, formatting is exact and
  non-negotiable: exact hex codes, exact number-format strings, exact border
  weights, all pulled from `scripts/style_constants.py`. **Never** retype a hex code
  or a number format from memory — that is what the script is for.

## Mode 0 — preserve vs. impose (decide this FIRST)

**When a workbook already has an established format, your job is to replicate
that workbook's exact formatting — clone its style records verbatim — not to
re-derive formatting from the house conventions.** The design language below is
for *authoring*: unformatted targets, new builds, or an explicit instruction to
restyle. Re-deriving an existing format always loses its long tail (per-block
tint bands, hidden white plumbing text, scaling number formats, year-separator
borders, drawings, charts, cached values) and is therefore wrong by default.

- Restoring / round-tripping / "make it match the original" / applying one
  version's look to an updated copy → **clone mode**: `scripts/clone_format.py
  SOURCE.xlsx TARGET.xlsx -o OUT.xlsx`. It works at the package level — styles,
  theme, numFmts, sheetViews, row/col geometry, tab colors, drawings, charts,
  media, cached values are carried byte-for-byte from the source; only content
  differences are patched in.
- Formatting a workbook with no established style (or the user explicitly wants
  the house style imposed) → the audit → review → apply workflow below.

Know your writer's losses: any openpyxl round-trip **strips cached formula
values** (recalculate before delivery — e.g. headless LibreOffice — or use
clone mode), **drops embedded images**, and **re-serializes charts lossily**.
Never push a workbook containing drawings/charts through apply_format without
restoring those parts from the source package afterwards (clone_format's
machinery is the reference for how).

## The hard rules (never break these)

- **Never change cell values or formulas.** Only styling changes.
- **Never re-derive formatting a workbook already has.** Established format →
  clone it exactly; impose the house language only on unformatted targets or
  explicit request.
- **Never delete, insert, or reorder** rows, columns, or tabs.
- **Never invent structure the target doesn't have.** No toggles in the model → don't
  fabricate toggles. Annual-only model → no monthly-axis machinery.
- **When the target lacks a feature, skip that convention.**
- **Apply everything proportionally to the model's actual structure** — never
  reproduce the example model's cell addresses, row counts, section names, division
  names, or "8 header rows" assumption.

The scripts enforce the first two by construction (`apply_format.py` only ever sets
`.font` / `.fill` / `.number_format` / `.alignment` / `.border`, plus sheet-level
gridlines / freeze / tab-color / page-setup). The rest is on you during review.

## Setup

The scripts need `openpyxl` (`pip install openpyxl` if it isn't present). Run them
with their own path so their imports resolve, e.g.
`python3 scripts/audit_workbook.py ...`. Work on a copy or use `-o` to write a new
file rather than overwriting the user's original in place.

## Workflow

### 1. Read the design language
Read `references/design-language.md` first — it is the exact lookup tables (colors,
fonts, number formats, fills, borders, chassis, page setup). Skim
`references/tab-patterns.md` (per-tab playbooks) and `references/classification.md`
(how to recognize things) so the vocabulary is loaded before you look at the target.

### 2. Audit the target
```
python3 scripts/audit_workbook.py TARGET.xlsx -o audit_map.json
```
This scans the workbook and writes a JSON map: per tab a tab-type guess + confidence
and geometry (label columns, data columns, header band); per row a row-class guess +
confidence; plus auto-detected input blocks on Assumptions tabs. The console summary
flags every low-confidence call. **These are guesses.**

### 3. Review and correct the map — this is where your judgment goes
Open `audit_map.json` and fix it, using `references/classification.md`:
- **Sample real cells** to confirm guesses — open the workbook and look. Don't trust
  a label keyword blindly ("Net Sales" may be a loaded line, not a total).
- **Check every tier.** Tier assignment is the script's weakest guess. Reason about
  the arithmetic: a total of totals outranks its components; the statement's bottom
  line / balancing figure is the headline (tier 3). Fix `tier` and the matching
  `class` (`subtotal` / `major_total` / `headline_total`).
- **Fix low-confidence rows**, banners vs. underlined section headers vs.
  sub-section labels, and any tab type the audit was unsure about.
- **Add input blocks** the audit couldn't see — e.g. live-column advance rates on a
  liquidity tab. (On statements, hardcodes are *loaded data*, not inputs — leave them
  blue-font, not yellow. The audit only auto-proposes input blocks on Assumptions.)
- **Set `format_body: false`** on anything that should stay unformatted, and confirm
  working/source/divider tabs are marked that way.
- **Ask the user only about a genuine fork** — most commonly: "is the `Bridge` tab a
  working tab I should leave unformatted, or a presentation tab?" Don't ask about
  things you can settle by looking.

Map fields you'll most often edit: a row's `class` and `tier`; a sheet's `tab_type`,
`format_body`, `marker_col`, `label_cols`, `data_col_start`/`data_col_end`,
`header_rows`, `freeze`, `tab_color`, `numfmt_family` (`balance` for BS/CF), and
`input_blocks`.

Row classes the application understands: `banner`, `subbanner`, `subsection_label`,
`section_underline`, `subtotal`, `major_total`, `headline_total`, `percent`, `check`,
`unit_note`, `data_first`, `data`, `header_year`, `header_date`, `header_fy`,
`actual_forecast`, `toggle`, `label`.

### 4. Apply
```
python3 scripts/apply_format.py TARGET.xlsx --map audit_map.json -o TARGET.formatted.xlsx
```
Deterministic, and **idempotent** — safe to run twice; the second pass changes
nothing. It skips every tab marked `format_body: false` (their bodies are left
exactly as-is; a tab color is still applied). It never touches values or formulas.

### 5. Re-audit and report
Re-run the audit on the output (or spot-check in Excel) and tell the user: what was
formatted, what was deliberately skipped (and why — working/source/divider tabs),
and anything you couldn't classify and left alone. Be honest about residual
low-confidence areas so they can eyeball them.

## Adapting to different layouts

- **Labels already in col A** → adapt to the existing skeleton; don't insert columns
  unless the user approves restructuring.
- **Annual-only model** → a single annual block; no monthly axis, no "x" separator,
  no SUMIFS annuals.
- **Quarterly model** → same chassis, quarterly EOMONTH stepping.
- **No Assumptions tab** → skip toggles and green Assumptions links.
- **Anything the model doesn't have** → skip that convention. Format proportionally to
  what's actually there.

## A note on real-world workbooks

openpyxl re-serializes the whole file on save, which can nudge the *text
representation* of raw floating-point constants by a last digit (e.g.
`...0005` vs `...001`) — Excel stores and shows the identical number. This is a
property of any openpyxl-based formatter, not a value change; formulas round-trip
exactly. The skill never assigns a cell value.

## What's in here

- `references/design-language.md` — exact lookup tables for every value.
- `references/tab-patterns.md` — per-tab-type playbooks, incl. "working tabs stay
  unformatted".
- `references/classification.md` — how to classify tabs/rows/cells, and the tier
  logic.
- `scripts/style_constants.py` — single source of truth: every hex, number format,
  font, border + style factories (incl. `tint()` for band/date-row derivation).
- `scripts/audit_workbook.py` — scans a workbook → reviewable JSON map.
- `scripts/apply_format.py` — applies the reviewed map; idempotent; never touches
  values/formulas; skips working tabs; never moves an existing freeze or print setup.
- `scripts/clone_format.py` — lossless package-level formatting transplant for
  already-formatted workbooks (clone mode).
