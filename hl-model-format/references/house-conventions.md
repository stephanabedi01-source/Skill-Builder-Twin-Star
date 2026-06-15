# House conventions & lessons log

A running record of what real, correctly-formatted models in this firm's style
actually do — distilled across the example models (Twin Star, JELD-WEN, Traeger).
Read it the way the "Generalization first" section in SKILL.md says: these are
**evidence and house tendencies, not constants**. Every model still gets its own
values read from its own theme/styles (Mode A) or chosen to fit it (Mode B). The
point of this log is to make Mode B's *defaults* land closer to the house bar so
there's less to correct.

How to use it in Mode B: set a top-level `palette` in the map (apply_format reads
it) and per-sheet `font_family` / `font_size`, rather than relying on the Twin
Star defaults baked into style_constants.

## The recurring firm palette (theme, shared across models)

Every example shares one theme: `dk1 #525766`, `lt1 #FFFFFF`, `dk2 #002855`,
`lt2 #0067A5`, accents `508BC9 / BCBFC6 / 7E8597 / 24B1B1 / 6C1E36 / 9FC3DA`.
Tab colors are stored as **theme index + tint**, not raw RGB. So when you author
tab colors in Mode B, use tints of these (e.g. accent1@0.8 ≈ `DCE8F4`).

## What varies model-to-model (decide per model — do NOT hardcode)

- **Body font.** Twin Star = Arial 10. JELD-WEN & Traeger = **Segoe UI** (11 on
  statements, ~8 on dense output pages). The firm's newer models trend Segoe UI;
  default Mode B to Segoe UI unless the model says otherwise, sized to the tab's
  density (statements 10–11; tight output grids 8).
- **Tab-color family → section mapping.** The *hues* are house, but which section
  gets which hue differs. Traeger: core financials = accent1@0.8 (light blue),
  output pages = `dk2` navy, transaction scenarios = accent5 maroon, market/data
  = accent4@−0.25 teal. Twin Star: statements blue, outputs yellow, working green,
  source light-blue. Pick a coherent per-section scheme; don't assume a fixed one.
- **Total-fill palette.** Can be bold (Twin Star: `DCE8F4` / `D3EFEF`) or very
  restrained (Traeger: subtotals `F2F2F2`, headline ≈ accent2@0.8 `F2F2F4` — nearly
  the same, distinguished by **bold + a thin `BCBFC6` rule** above & below). When
  unsure, lean restrained; let bold and a hairline rule carry the hierarchy.
- **Section headers.** Three house treatments seen: navy fill banner + white text;
  bold `0067A5` sub-label; and **bold `dk2` navy TEXT, no fill** (Traeger's
  "Revenue" / "Operating Expense"). Use the `section_header` class for the last.
- **Total-rule color.** Twin Star `BFBFBF`; Traeger `BCBFC6` (accent2). Both gray.
- **Markers.** Twin Star pure red `FF0000`; Traeger dark red `C00000`.
- **$ number format.** Twin Star `$#,##0_);($#,##0);-;` (dash for zero); Traeger
  uses `"$"#,##0_);("$"#,##0)` (no zero clause) widely. Both fine; match the model.

## What's consistent (safe Mode B defaults)

- Body text slate gray `#525766`, never black. Hardcoded inputs blue `#0000FF`.
- Output/presentation pages: **all gray**, no provenance blue/green.
- Negatives in parentheses; margins/percent rows italic, one decimal.
- **Gridlines OFF on every client-facing tab — including data/exhibit tabs.** The
  Twin Star "working tabs keep gridlines on" rule is the exception, not the norm;
  reserve gridlines-on for genuinely internal scratch the client won't see.
- **Freeze sparingly** — typically only the main multi-screen statement, not every
  schedule. Don't freeze short output/exhibit tabs.
- Delivery standard (first tab, A1, 85% zoom) always, regardless of the model's
  own saved views.

## Per-tab notes from the Traeger test (read as patterns, not addresses)

- **Historical Financials (spread):** col B `x` markers (dark red); col C section
  headers (navy bold text) + subtotals; col D indented line items; data from the
  first period column. Period header = year band (centered bold) + quarter labels +
  a date-serial row formatted `mmm-yy`. `($ millions)` is an italic unit note.
  Subtotals (Total Revenue, Operating Income, Income Before Tax) = `F2F2F2`;
  headline (Gross Profit, Net Income, Adjusted EBITDA) = the headline tint. Freeze
  below the period header. Segoe UI 11.
- **O1 / O2 / O3 (output pages):** Segoe UI **8**, all gray (links included), navy
  tab color, "Year Ended" header centerContinuous, same total treatment, no freeze.
- **Scenario tabs (Amend & Extend, Uptier…, Drop-Down):** maroon tab; section
  headers + input cells (blue) + subtotals; varied per-tab layout — classify each.
- **Market/data tabs (Steel Pricing, Pricing, Industry Data, Tariff Slide,
  Holders):** teal tab; gridlines off; bodies largely left as raw data (not
  heavily restructured) but still cleaned.

## Mistakes this log exists to prevent (from the Traeger Mode B attempt)

1. Defaulting to Arial when the model's style is Segoe UI.
2. Guessing tab-color families (yellow outputs) instead of the model's scheme.
3. Coloring section headers `0067A5` when the model uses navy text.
4. Over-emphatic cyan/blue total fills when the model is restrained gray.
5. Leaving gridlines on data tabs; over-freezing short tabs.
6. Pure-red markers when the model uses dark red.

## Bingo Industries test (second HL model, same template — confirmations + new lessons)

Bingo (Discussion Materials: Capital Structure, Historicals spread, O1–O3
outputs, Pricing data, a small Sheet1) confirmed the corrected Traeger
conventions transfer cleanly to another HL model. New, generalizable lessons:

- **Preserve the model's existing deliberate chrome; format only what's
  unformatted.** Bingo's body was unformatted (gridlines on, mixed
  Arial/Calibri/Aptos, ~no fills) but it already had **deliberate, house-scheme
  tab colors and freeze splits**. The right move is hybrid: Mode B the body, but
  *keep* the existing tab colors (`tab_color: null` so apply doesn't overwrite)
  and the existing freeze splits (omit `freeze` so the guardrail leaves them).
  Don't re-guess chrome the model already got right.
- **`freeze_panes` (openpyxl) conflates the freeze split with the saved scroll.**
  Bingo's Historicals read `B79` but the real split was 16 rows (`ySplit=16`)
  scrolled to row 79. `finalize_delivery` keeps the split and resets the scroll
  to its origin (`B17`) — that IS the delivery standard, not a lost freeze.
  Verify the split (`xSplit`/`ySplit`), not the reported `freeze_panes` string.
- **Rich even when "small".** A 7-tab "small" model still had 23 charts, 6
  drawings, 110k defined names, 3070 cached values — Mode B still routes through
  `merge_format` (package-safe), never a bare openpyxl save.
- Same restrained palette (Segoe UI 11/8, navy section headers, `F2F2F2`
  subtotals, accent2-tint headlines, `BCBFC6` rules, `C00000` markers) produced a
  clean result with no corrections needed beyond per-tab geometry.
