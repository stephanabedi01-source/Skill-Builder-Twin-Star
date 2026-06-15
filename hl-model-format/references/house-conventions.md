# HL formatting — purpose → format rules (the analyst's reasoning)

How a strong HL analyst formats: **every cell's format states its role.** Format by
what a cell *is* (input, calc, link, total, label, memo, header, check), not by where
it sits. This guide gives the **purpose → format** rules — specific enough for polished,
consistent output, general enough for a model you've never seen.

Evidence base: cell-by-cell study of 5 ground-truth keys (Twin Star, JELD-WEN,
Traeger, MITER, Nine Energy) — ~200 sheets, ~150k formatted cells. All five share one
**house theme**: `dk1 #525766 · lt1 #FFFFFF · dk2 #002855 · lt2 #0067A5 · accents
508BC9 / BCBFC6 / 7E8597 / 24B1B1 / 6C1E36 / 9FC3DA` (tab/fill colors are theme
index+tint, not raw RGB). Percentages below are how often a rule held across the keys.

---

## THE MASTER SWITCH: is this a live model tab or a presentation page?

This one decision gates everything else, and it's the rule earlier attempts missed.

- **Live model / working tab** (you build it, it has inputs + calcs): apply the full
  **provenance color code** below — blue inputs, gray calcs, green links.
- **Output / presentation page** (it mostly *pulls* other sheets via links, e.g. "IS
  Out", "O1", "LFCF Build", valuation summaries): **drop provenance — everything is
  gray `#525766`.** (On output pages: constants 75% gray, formulas 80% gray, links
  80% gray. On model tabs: constants 44% blue, formulas 88% gray, links ~41% green.)
  How to tell: a sheet that is >~55% cross-sheet links is an output page.

Everything below assumes a model tab unless it says otherwise.

## 1. Color = provenance/role  (why: a reader must see what's editable vs derived)

| The cell is… | so its font is… | evidence / notes |
|---|---|---|
| a hardcoded **input** (typed number, incl. a %/rate assumption) | **blue** `#0000FF` (some models `#0000E1`) | 44% of model-tab constants; the rest are loaded data / headers left gray |
| a **formula / calculation** (and most labels) | **gray** `#525766` (never black) | 88% of formulas |
| a **link to another sheet** (`=Sheet!…`) | **green** `#008000` or `#00B050` | a house convention applied ~half the time on model tabs; always gray on output pages |
| a **check / flag / nav "x"** | **red** | checks are italic, color red *or* gray (≈50/50) |
| a **"DRAFT / CONFIDENTIAL / Subject to…"** stamp | **`#C00000`** bold (some `#FF0000`) | top-right of any tab, not just covers (`C00000` 77%) |

- **Invariant:** blue=input, gray=calc, gray body never black, output pages all-gray.
- **Interpretation:** exact input blue (`0000FF` vs `0000E1`); whether links are
  greened at all (some analysts leave them gray); exact red (`C00000` vs `FF0000`).
- **Trap:** a hardcoded `%` is a **blue input**, not an italic-gray margin row. Decide
  by provenance first (constant → input → blue), unit second.

## 2. Number format = data type  (why: the format states the unit and aligns the column)

Pick by what the number **means**; it commonly varies **by column** within a row
(cap tables/assumptions: `$ amount | rate % | x multiple | maturity date`). Use
`column_numfmt` for mixed columns — one row format can't express them.

| The number is… | house format (dominant) | also seen |
|---|---|---|
| **currency** | `"$"#,##0_);("$"#,##0)` (no zero clause) | `…;"$" -` / `…;-` zero variants; accounting `_($* #,##0…)` |
| **plain (thousands)** | accounting `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | `#,##0_);(#,##0);–?` (en-dash zero, an HL favorite); `#,##0_);(#,##0)` |
| **percent** | `#,##0.0%_);(#,##0.0%);#,##0.0%_)` (1 dp, parens) | `0.0%`; `0%`; hardcoded % input often `0.00%` (2 dp) |
| **multiple** | `0.0"x"` | `#,##0.0\x ;"NMF"` |
| **date / period** | `[$-409]mmm-yy;@` | `m/d/yy`, `mm-dd-yy` |
| **count / ratio** | `0` / `0.0` | |

- **Invariant:** negatives in **parentheses**; zero shown as a **dash** (often padded
  to align under the digits); `$` on the first row of a block and on totals; the unit
  is always legible from the format. State scale once per tab ("$ in millions").
- **Interpretation:** the exact string (accounting `_(*…` vs `#,##0…–?` vs `;-;`),
  decimals (0 vs 1), `$`-on-every-row vs first+totals. Match the model; when authoring
  be consistent within a block and pick from `NF_HL_*` in style_constants.

## 3. Fills + weight + borders = the rollup hierarchy  (why: show structure at a glance)

- **Subtotal / total rows:** **bold**, a light fill, and a thin gray rule above & below
  (`#BFBFBF` or accent2 `#BCBFC6`). Fills seen: `#F2F2F2` (most common), `#DCE8F4`,
  `#F2F2F4`, `#D3EFEF` — all light tints. Many totals carry **no fill** and rely on
  bold + the rule alone. So: hierarchy = bold + (light fill and/or rule), restrained.
- **Inputs:** yellow `#FFFFCC` fill — the triple-mark (yellow + blue font + a thin blue
  box) — **including assumption/driver blocks** (easy to miss).
- **Output pages** sometimes add a light header band (accent1@0.8 `#DCE8F4`).
- **Invariant:** totals are bold and set off; inputs are yellow-boxed.
- **Interpretation:** which light tint (or none); 2 vs 3 tiers; rule vs fill emphasis.
  When unsure, lean restrained (gray fill + bold + hairline rule).

## 4. Italics = "this is a ratio, memo, or note"  (why: de-emphasize non-primary rows)

Italic marks **margin/%/growth rows, check rows, memo/footnote rows, and per-unit
ratios** — derived or secondary figures sitting under primary numbers. It does **not**
mark inputs (a blue % input is upright). Margin rows are italic **gray**.

## 5. Labels, sections, headers

- **Section headers** (a label row heading a group): **bold**, no fill; color is gray
  or `lt2 #0067A5`, or **navy `dk2 #002855` text** (a strong house look). Banner-fill
  headers (navy fill, white text) also occur. Pick one treatment per model; ~62% bold.
- **Column / table sub-headers** ("Description", "Face", "Rate", "Maturity", "Year
  Ended") are **gray bold** (sometimes underlined) — **not** section headers.
- **Line items** under a section: left-aligned with a real **indent** (1; 2/3 for
  sub-items) — indentation is `alignment.indent`, not leading spaces.
- **Title block:** title bold (larger, 12–14); subtitle italic; "as of" + unit note
  plain **gray** (the unit note is gray, *not* purple — purple was Twin-Star-only).
- **Period header band:** year row centered bold; period/quarter labels; a date row
  formatted `mmm-yy`. Centered.

## 6. Fonts

- **Segoe UI** is the house font (Arial is legacy; raw-data tabs may keep Aptos Narrow
  — don't force-convert them). Sizes run **9–11** (10–11 body, 9 dense/footnote; titles
  12–14). 12 is not the default (Bingo's 12 was atypical). Match the model; default
  Segoe UI 10–11 for body.

## 7. Chrome: tab colors, gridlines, freeze, zoom, print

- **Tab colors:** one theme-tint family per workbook section; the hue→section mapping
  is per-model (e.g. core/light-blue, outputs/navy, scenarios/maroon, data/teal). Tints
  vary (0, 0.6, 0.8, −0.25). Cover/divider tabs uncolored.
- **Gridlines OFF** on every client-facing tab (off 128 : on 19; "on" only on genuine
  scratch).
- **Freeze sparingly** — the main multi-screen statement below its period header; not
  short output/exhibit tabs. Preserve a model's existing deliberate freeze split.
- **Print/page** off-gridlines, portrait/letter, fit-to-width.
- **Delivery standard** (the one override): open on first tab, A1 + pane-origin scroll,
  85% zoom — always, regardless of the model's saved views.

## 8. Preserve, then build

Models arrive partly formatted (deliberate tab colors / freeze splits, unformatted
body). Format only what's missing: keep existing tab colors (`tab_color: null`) and
freezes (omit `freeze`); build the body. A per-tab Mode A / Mode B blend.

---

## Calibration: invariant vs up-for-interpretation

**Invariant (always do):** format-by-role; gray body never black; blue inputs; output
pages all-gray; parentheses negatives + dash zero; totals bold & set off; inputs
yellow-boxed; gridlines off; delivery standard. These produce "HL-correct" reliably.

**Up for interpretation (judgment / match the model):** exact input-blue and red
shades; whether links are greened; which light tint for totals (or none); the exact
number-format string per unit; font size 9/10/11; section-header treatment (gray vs
lt2 vs navy text vs banner); tab-color hue→section mapping; `$`-on-every-row vs
first+totals. Choose one coherent answer per model and apply it consistently.

## Open questions resolved by the 5-key study (vs the prior log)
- Headline-total fill is **not** one color — `F2F2F2`/`DCE8F4`/`F2F2F4`/`D3EFEF` all
  occur (or none + bold + rule). Restraint, not a fixed hex.
- Body size norm is **9–11**, not 12 (Bingo was the outlier).
- `$` dominant string is `"$"#,##0_);("$"#,##0)`; plain-number dominant is the
  accounting `_(*…` and the en-dash `…–?` forms. Defaults updated in style_constants.
- Data/exhibit tabs are largely left native (font/Aptos Narrow) but gridlines-off and
  gray-bodied; not heavily restructured.
- Section-header treatment genuinely varies by model/team — it is interpretive.

## Per-model evidence log
- **Traeger key:** Segoe UI not Arial; model's own tab-color families; navy section
  text; restrained gray totals; gridlines off on data tabs; dark-red markers.
- **Bingo key:** number format is per-column by data type (→ `column_numfmt`); %-input
  ≠ margin row; assumptions get yellow fill; unit note gray; column sub-headers gray
  bold not navy; DRAFT/CONFIDENTIAL red on any tab.
- **5-key cross study:** the master switch (model-tab provenance vs output-page gray);
  the house number-format set; Segoe UI + 9–11 sizing; gridlines-off norm; C00000
  warnings; checks italic (red or gray); totals = bold + restrained fill/rule.
