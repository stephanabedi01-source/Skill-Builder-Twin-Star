# HL formatting — themes, invariants & options

The growing pattern library for Houlihan Lokey model/material formatting. Built by
diffing my output against ground-truth keys (Twin Star, JELD-WEN, Traeger, Bingo;
more coming). Read with the "Generalization first" rule in SKILL.md: when a model
already has a format, replicate **its own** values (Mode A); when it doesn't, these
themes are how you choose well (Mode B). The aim of this file is to capture **why
and when** each formatting choice is made, what is **invariant**, and what is **up
for interpretation (your options)** — so replication and from-scratch creation both
improve.

The organizing principle behind everything below: **formatting encodes meaning.**
Every visual choice answers a reader's question — "is this an input I can change?
a subtotal? a different unit? a draft?" Format by what a cell *means*, not by where
it sits.

---

## 0. Invariants (true in every HL model seen)

- One shared **theme palette**: `dk1 #525766`, `lt1 #FFFFFF`, `dk2 #002855`,
  `lt2 #0067A5`, accents `508BC9 / BCBFC6 / 7E8597 / 24B1B1 / 6C1E36 / 9FC3DA`.
  Colors are stored as **theme index + tint**, not raw RGB.
- **Body text is slate gray `#525766`, never black.**
- **Color = provenance/role** (the banker code): blue hardcoded inputs, gray
  formulas/labels, green cross-sheet links, red checks/flags/warnings.
- **Negatives in parentheses; zero shown as a dash.**
- **Gridlines OFF** on every client-facing tab.
- Output/presentation pages drop provenance color → **all gray**.

---

## 1. Color = role  (theme: a reader must see what's editable vs derived)

| Role | Color | When |
|---|---|---|
| Hardcoded input | blue `#0000FF` | a typed constant — **including assumption %/rate inputs** (a blue % input is still an input, NOT a margin row) |
| Formula / label | gray `#525766` | any calc or text |
| Cross-sheet link | green `#00B050` | formula pulling another sheet (suppressed to gray on output pages) |
| Check / flag / nav "x" | red | check rows, error flags, the col-B "x" markers |
| **"DRAFT" / "CONFIDENTIAL" / disclaimer** | red bold | **on ANY tab**, usually the top-right header corner — not just a cover |

- **Invariant:** the blue/gray/green/red meaning; gray (not black) body; outputs all gray.
- **Interpretation / options:** exact red — pure `FF0000` (Twin Star) vs dark `C00000`
  (Traeger, Bingo); whether links are even used (output pages don't).
- **Mistake to avoid:** italic-graying a hardcoded `%` input because it "looks like a
  margin." Inputs stay blue; only *derived* margin/growth rows are italic gray.

## 2. Number format = data type  (theme: the format states the unit and aligns the column)

Choose the format by **what the cell holds**, and it commonly **varies by column**
within one row — a cap-table or assumption row is `$ amount | rate % | x multiple |
maturity date` across adjacent columns. A single row-level format cannot express
that; use per-column formats (`column_numfmt` in the map).

| Data type | Format (representative) |
|---|---|
| Currency, first row of a block / totals | `"$"#,##0_);("$"#,##0)` or with a `$`/dash zero clause |
| Currency, subsequent rows | `#,##0_);(#,##0); -` (zero = dash, often space-padded to align) |
| Percent (margin/derived) | `0.0%` / `#,##0.0%_);(#,##0.0%)` — one decimal |
| Percent (hardcoded input) | often `0.00%` (two decimals) — still a blue input |
| Multiple | `#,##0.0x ;"NMF"` |
| Date | `mmm-yy` (periods) or `m/d/yy` |
| Count / ratio | `0`, `0.0` |

- **Invariant:** parentheses for negatives; zero as a dash; `$` on first-of-block and
  totals; percent and date formats present where those units appear.
- **Interpretation / options:** exact strings — the zero clause (`-;` vs `;-` vs
  space-padded `\-\ \ \ ` for alignment), `$` with vs without a zero clause,
  decimals (0 vs 1), `mmm-yy` vs `m/d/yy`. Match the model; when authoring, be
  consistent within a block and align zeros under the digits.

## 3. Fills = hierarchy + inputs  (theme: fills mark rollups and editable cells)

- **Subtotals** light gray `#F2F2F2`; **headline totals** a slightly distinct tone
  (Twin Star `D3EFEF`; Traeger/Bingo accent2@0.8 `F2F2F4` — barely different, the
  **bold + a thin rule above/below carries the hierarchy**). Total rule color is a
  gray (`BFBFBF` or accent2 `BCBFC6`).
- **Inputs** get yellow `#FFFFCC` fill (the triple-mark: yellow + blue font + a thin
  blue box) — **this includes assumption/driver blocks**, which is easy to miss.
- Output pages sometimes add a light header band (accent1@0.8 `DCE8F4`) and explicit
  white fills.
- **Invariant:** inputs yellow; totals filled + bold + ruled; subtotal lighter than
  it sits above.
- **Interpretation / options:** how loud the tier fills are (restrained gray vs the
  bolder cyan/blue), and how many tiers (2 is common; 3 if the statement needs it).
  Lean restrained when unsure.

## 4. Fonts  (theme: one coherent family per surface)

- Family: **Segoe UI** in modern HL models (Arial in older ones). Data tabs may keep
  their native font (e.g. Aptos Narrow) — don't force-convert raw data grids.
- Size is **interpretive and tab-dependent**: primary statements/outputs often **12**
  (Bingo), schedules **10–11**, dense/footnote **9**, titles **12–14**. Match the
  model; when authoring, default ~Segoe UI 10–11 for body, larger for titles.
- Title block: **title bold (larger), subtitle italic, "as of" + "($ millions)"
  plain gray** (the unit note is **gray**, not purple — purple was Twin-Star-only).
- **Invariant:** one body family per surface, gray body, titles bold.
- **Interpretation / options:** exact size; Segoe UI vs Arial; whether "Segoe UI
  Bold" is a named font or just the bold flag (equivalent).

## 5. Section vs sub-headers  (theme: signal the outline, not every label)

- **Section headers** ("Revenues", "Operating Expenses", "Memo:", "Current
  Liquidity"): three house treatments — **navy `dk2` bold TEXT, no fill**
  (Traeger/Bingo); navy fill banner + white text; or bold `lt2` sub-label. Pick one
  per model and stay consistent.
- **Column / table sub-headers** ("Description", "Face", "Rate", "Maturity") are
  **NOT** section headers — they are **gray bold** (sometimes underlined), not navy.
- Line items below a section are **left-aligned, indent 1** (indent 2/3 for
  sub-items). Indentation is real (`indent`), not leading spaces.

## 6. Chrome: tab colors, gridlines, freeze, zoom

- **Tab colors**: one theme-tint family per workbook section; the *hue→section*
  mapping is per-model (Traeger: core light-blue, outputs navy, scenarios maroon,
  data teal; Bingo: cap-struct accent2@.8, statements/outputs navy, data teal).
- **Gridlines off** on every client-facing tab (data/exhibit tabs included).
- **Freeze sparingly** — the main multi-screen statement, below its period header;
  not short output/exhibit tabs.
- **Delivery standard** (the one place we override the model): open on the first
  tab, every tab selection A1 + scrolled to pane origin, **85% zoom**. Real keys
  use varied zoom and saved scroll; we normalize at delivery per instruction.

## 7. Preserve, then build  (theme: don't re-guess what the model got right)

Models often arrive **partly formatted** — deliberate tab colors and freeze splits
already set, body unformatted. Format only what's missing: keep existing tab colors
(`tab_color: null`) and freeze splits (omit `freeze`), and build the body. This is a
per-tab Mode A / Mode B blend.

---

## Recurring tab archetypes

- **Spread / Historicals:** col-B/C "x" markers; section headers (navy) + subtotals
  in the primary label col; indented line items in the secondary; period header band
  (years centered bold, quarter labels, a `mmm-yy`/`m/d/yy` date row); margin rows
  italic; check rows red italic; `$` on first-of-block + totals.
- **Output pages (O1/O2/O3):** all gray, smaller font, "Year Ended" header
  (centerContinuous), same total treatment, sometimes a light header band; no freeze.
- **Capital Structure:** assumption inputs up top (blue, **yellow-filled**, % to 2dp);
  a debt table with **per-column** formats ($ face | rate % | x | maturity date);
  Total Debt / Net Debt / Total Liquidity as totals; "Memo:" / "Current Liquidity"
  section headers.
- **Scenario tabs:** maroon family; inputs + subtotals; layout varies — classify each.
- **Market/data tabs:** teal family; gridlines off; body largely left as native raw
  data (may keep Aptos Narrow), gray text.

---

## Lessons log (evidence)

**Traeger key** — corrected my first Mode B: use Segoe UI not Arial; tab-color
families are the model's (not yellow outputs); navy section text not `0067A5`;
restrained gray totals not cyan; gridlines off on data tabs; sparing freeze; dark-red
markers.

**Bingo key** — diffing my output against the real key (same model) surfaced misses I
*hadn't* caught (my earlier "clean" claim was premature, made before I had the key):
- **Number formats were mostly wrong on the spread/cap-structure** because I applied
  one format per row. The key formats **per column by data type** ($/%/x/date). →
  added `column_numfmt`; this is the #1 fix.
- **Assumption % inputs**: I italic-grayed them as "margin rows"; the key keeps them
  **blue inputs with a `0.00%` format and a yellow fill**. % format ≠ margin row.
- **Inputs need yellow fill** — I set no input blocks on Bingo and missed all of them.
- **Unit note "($ millions)"** is **gray**, not purple.
- **"DRAFT" / "CONFIDENTIAL"** (top-right of working tabs) are **red** — now applied
  on any tab, not just covers.
- **Column sub-headers** ("Description/Face/Rate") are **gray bold**, not navy
  section headers.
- **Font size**: statements/outputs were **12** (I used 11/8); cap-structure 10.
  Size is genuinely interpretive — match the model.
- Confirmed: preserve existing tab colors/freezes; even "small" models are rich
  (route through `merge_format`); zero value/formula changes held.

## Open questions to resolve with the next 5–10 models

- Is the headline-total fill consistently accent2@0.8, or does it vary (cyan, blue)?
- Default body size — is 12 the norm for primary statements across models, or Bingo-
  specific?
- Exact dollar zero-clause and dash-padding conventions — is there one house string?
- How much do data/exhibit tabs get touched (font/color) vs left native?
- Section-header treatment — does a given era/team consistently pick navy-text vs
  banner vs sub-label?
