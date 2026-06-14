---
name: model-format
description: Format Excel financial and operating models. Mode A losslessly clones a workbook's established formatting (exact style records, theme, number formats, views, charts, drawings, cached values); Mode B formats an unformatted, stripped, or messy model from scratch to a clean professional standard (banker color-coding with blue inputs, gray formulas, green links; tiered total fills; section banners; exact number formats; widths, freezes, tab colors, print setup) — every deliverable finished to the delivery standard (opens on first tab, all tabs at A1, 85% zoom). Use whenever the user uploads or points to an .xlsx and asks to format it, apply the house style or "our model format", make it look like our models or like the original, restore or preserve formatting, format or clean up a messy or stripped model, color-code inputs vs formulas, fix number formats, or make a model client-ready. Triggers on "format this model", "apply house style", "match the original formatting", "make this look professional".
---

# Model Format

## ⚠️ Generalization first — specific values are EVIDENCE, not constants

**This skill is applied to many different models, not just the Twin Star model
its examples are drawn from.** Every model is genuinely different: different
palettes, themes, fonts, layouts, freeze panes, number formats, tab colors,
sheet names, and sheet structures.

**The instructions in this skill are not always to be taken literally. They are
general practices, illustrated with evidence from specific models. The general
practices apply to every model; the specific values do not.** Treat every
concrete value anywhere in this skill — every hex color (`#002855`, `#DCE8F4`,
`#D3EFEF`, `#F2F2F2`, the row-6 tint bands, …), theme-slot assignment,
freeze-pane location (`E9`, `BY9`, `B5`, …), tab color, font and size (Arial
9/10, Aptos Narrow 14, …), number-format string, sheet name, cell reference, row
height, and column width — as **evidence from one workbook, not a universal
constant.** Never paint a different model with Twin Star's palette or geometry
just because this skill happens to mention those values.

**The one rule that IS universal:** for any workbook with an established format,
read *that workbook's own* theme, style records, number-format strings, sheet
views and freeze panes, tab colors, row/column geometry, drawings, and charts —
and reproduce them exactly. **Copy the source's exact formatting; do not
reinterpret it through your own conventions.** The specifics change with every
model; the discipline does not. (`scripts/clone_format.py` enforces this
mechanically — see Mode 0.)

### How to read the literal rules
Every rule written against a specific value must be read in its **general** form.
The translation pattern is always *[specific fix from one model] → [match
whatever the source actually has]*:

| Written as (evidence from one model) | Read as (the actual, general rule) |
|---|---|
| "restore the freeze pane at `BY9`" | preserve the source's exact freeze pane, wherever it is — or its absence |
| "row-6 bands are `#B9D1E9` / `#E4E5E8` / …" | preserve each cell's exact fill, whatever color the source uses |
| "use the trailing-comma scaling number format" | copy the source's exact number-format string, whatever it is |
| "output pages are Arial 9" | match the source's exact font and size there, whatever they are |
| "tier 2 = `#DCE8F4`" | use whatever fill the source actually assigns that row (often a theme tint) |

### When a model does something no example covers
When a new model does something none of these examples mention — a different
palette, a different header construction, gradient fills, unusual fonts, a layout
the chassis doesn't describe, anything — the answer is the same general practice:
**replicate what that model actually does.** The absence of an example is *not*
permission to fall back on your own default formatting, or on Twin Star's. Read
the source; reproduce the source.

---

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

## The two modes, named

- **Mode A — Replicate.** The source has a deliberate, established format →
  clone it exactly. This is everything Mode 0 above describes
  (`clone_format.py`, package-level fidelity, no re-derivation).
- **Mode B — Format from scratch.** The source is unformatted, stripped, or a
  chaotic mess → apply your own professional formatting system and make it look
  excellent. **In this mode your judgment IS the standard — there is no key to
  copy.** Read `references/from-scratch.md` for the bar to hit (coherent font
  system, provenance color coding, exact number formats with negatives in
  parentheses / dashes for zeros / sensible $ % and unit scaling, clear section
  headers, distinct subtotal vs total treatments, a real alignment and indent
  hierarchy, sensible column widths and row heights, useful freeze panes,
  organizing tab colors, clean print setup). Mechanically it is the audit →
  review → apply workflow below, with the Mode-B map fields (`indent`,
  `auto_widths`, `column_widths`, `row_heights`) in play.

**Mode B on a rich workbook — apply the STYLING, not an openpyxl round-trip.**
`apply_format.py` writes through openpyxl, which on a feature-rich file strips
cached formula values, drops chart style/rels parts and drawings, expands shared
formulas, and loses defined names — the same losses Mode A avoids. So when the
target has charts, data tables, or cached values, do not ship the openpyxl output
directly. Run `apply_format.py` on a copy to produce a STYLE DONOR, then transplant
only its styling onto the original package with `scripts/merge_format.py`:

```
python3 scripts/apply_format.py SOURCE.xlsx --map MAP.json -o STYLE_DONOR.xlsx
python3 scripts/merge_format.py SOURCE.xlsx STYLE_DONOR.xlsx -o OUT.xlsx
python3 scripts/finalize_delivery.py OUT.xlsx
```

`merge_format` carries the donor's `styles.xml`, per-cell style indices, styled
empty cells (so banner/tier fills span), column widths, sheet views, and tab
colors onto the source — while keeping the source's cell values, formulas, cached
results, charts, drawings, media, and defined names byte-for-byte. For a plain
workbook with none of those features, `apply_format.py` output can be delivered
directly (still finish with `finalize_delivery.py`).

**How to choose: read the workbook.** A consistent, intentional existing format
means Mode A. Absent, stripped, or sloppy/inconsistent formatting means Mode B.
If an odd-but-consistent format leaves you genuinely unsure, ask. And in both
modes the iron rule holds: **formatting only — never change values, formulas,
structure, or defined names.**

## Delivery standard — the final save on EVERY workbook handed back

Apply this to every workbook you return, in either mode, as the last step before
delivery (`scripts/finalize_delivery.py` does all of it surgically, without
disturbing styles, charts, drawings, or cached values):

- The workbook **opens on the first tab** (workbook active tab = first visible
  sheet; only that sheet carries `tabSelected`).
- **Every tab's selection is on A1** and scrolled to the top-left corner
  (`activeCell`/`sqref` = A1, saved `topLeftCell` cleared). Frozen panes are
  kept exactly as they should be — just select A1 and scroll each pane back to
  its origin.
- **Every tab at 85% zoom** (`zoomScale="85"` on the sheet's active view).

Precedence, explicitly: this delivery standard is a **deliberate exception** to
the "preserve the source's view settings exactly" rule. Active tab, selected
cell, scroll position, and zoom always follow this standard at delivery time —
on every model — even when the source file had something different. Everything
else about views (gridlines on/off, view mode, freeze panes) still follows the
existing rules.

```
python3 scripts/finalize_delivery.py DELIVERABLE.xlsx
```

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
- `scripts/clone_format.py` — Mode A: lossless package-level formatting transplant
  for already-formatted workbooks (clone mode).
- `scripts/merge_format.py` — Mode B package-safe applier: transplants a style
  donor's styling onto a rich source package without openpyxl's content losses.
- `scripts/finalize_delivery.py` — the universal final-save step: first tab
  active, every tab at A1 and 85% zoom; surgical (views only).
- `references/from-scratch.md` — the Mode B playbook: what "excellent" means
  when there is no key and your judgment is the standard.
