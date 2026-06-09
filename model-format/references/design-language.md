# Design Language — exact lookup tables

The complete house style, as exact values. `scripts/style_constants.py` is the
machine-readable copy of this file; the two must always agree. When formatting,
the script pulls every value from `style_constants.py` — never retype a hex code
or a number-format string from memory.

All colors are 6-digit RGB hex. The script adds the `FF` alpha prefix and writes
**resolved RGB** (never theme indices or tints), so the style reproduces faithfully
on any target workbook regardless of its theme.

- [1. Fonts](#1-fonts)
- [2. Font color = data provenance](#2-font-color--data-provenance)
- [3. Fills = 3-tier total hierarchy + special fills](#3-fills--3-tier-total-hierarchy--special-fills)
- [4. Borders](#4-borders)
- [5. Number formats](#5-number-formats)
- [6. The universal sheet chassis](#6-the-universal-sheet-chassis)
- [7. Workbook architecture](#7-workbook-architecture)
- [8. Page setup](#8-page-setup)

---

## 1. Fonts

| Where | Font | Size |
|---|---|---|
| Every model / working tab | **Arial** | 10 |
| Output / presentation-page tabs ("IS Out", "UFCF Out", "CIM Outputs") | **Segoe UI** | 10 |

Default body text color is slate gray `#525766` — **never pure black**. Every
label, formula result, and header that would normally be black is this gray. This
is the single most distinctive trait of the style.

## 2. Font color = data provenance

Font color encodes where a value came from. Applied at the **row** level when a
whole row shares one provenance (e.g. an entire loaded historical row is blue),
otherwise per cell.

| Color | Hex | Meaning |
|---|---|---|
| Slate gray | `525766` | All formulas / calculations and labels (the default) |
| Blue | `0000FF` | Hardcoded inputs **and** pasted/loaded data — entire historical data rows on statements are blue |
| Green | `00B050` | Links pulling from another worksheet (e.g. `=Assumptions!$D$2`) |
| Red | `FF0000` | Check rows, error flags, bold "X" nav markers in col B, to-do notes, confidentiality text |
| Gray-purple, *italic* | `666699` | Unit notes ("#'s in thousands"), stated once per tab near the top, in the label column |
| Medium blue, **bold** | `0067A5` | Sub-section labels ("Sales Growth %", "Cash Flows from Operating Activities:") |
| White, **bold** | `FFFFFF` | Text inside dark banner fills |

Provenance is mechanical and the script derives it live from cell content:
formula referencing another sheet → green; other formula → gray; constant → blue.
The classification you supply only needs to carry *semantic* judgment, not a color
per cell.

## 3. Fills = 3-tier total hierarchy + special fills

Identical on every tab. **All three total tiers use bold text.** Tier fills (and
banner fills) span from the **label column through the last data column** — never
just the label cell.

| Fill | Hex | Use |
|---|---|---|
| Light gray | `F2F2F2` | **Tier 1: subtotals** (Total Product Cost, Total Current Assets, Net cash from operating activities) |
| Light blue | `DBE8F4` | **Tier 2: major totals** (Total Cost of Sales, EBITDA, Total Liabilities, net change in cash) |
| Pale cyan | `CDF5F5` | **Tier 3: headline totals** (Gross Profit, Total Assets, Total Liabilities & Equity, Cash at end of period, Total Availability) |
| Pale yellow | `FFFFCC` | **Input cells** — ALWAYS triple-marked: this fill **+** blue `0000FF` font **+** thin blue `0000FF` box around the whole rectangle of inputs (one box, not per-cell) |
| Dark navy | `002855` | **Section banner rows** — white bold text; fill extends from the label column across ALL data columns of the row |
| Steel blue | `9FC3DA` | **Secondary / sub-group banners** ("Cost of Goods Sold", "Operating Expense" on an Assumptions tab) — white bold text |

## 4. Borders

| Element | Border |
|---|---|
| Totals / subtotals (all three tiers) | thin `BFBFBF` gray **above and below** the total row, spanning label col → last data col. The top border sits on the total row itself; the row above also carries a bottom border. |
| Input blocks | thin `0000FF` blue **box around the rectangle** of input cells (one box, not per-cell) |
| Everything else | no gridlines — turn worksheet gridlines **OFF** on all presentation tabs; leave them **ON** on source/working tabs |

## 5. Number formats

EXACT strings — copy character-for-character, never paraphrase. (Excel may store
these on disk with the `$` quoted and `( ) -` escaped; that is the same format.
openpyxl round-trips our literal strings faithfully.)

| Context | Format string |
|---|---|
| First `$` row of a block (Gross Sales, Cash, Net Income) | `$#,##0_);($#,##0);-;` |
| Subsequent rows in the block | `#,##0_);(#,##0);-;` |
| Subsequent rows — Balance Sheet / Cash Flow variant | `#,##0_);(#,##0);-_);` |
| Total rows (all tiers) | `$#,##0_);($#,##0);$-_);` |
| Margin / percent rows (italic, gray, under parent) | `0.0%_);(0.0%);-;` |
| Assumption % inputs | `0.0%_);(0.0%);-_);` |
| Capex-style `$` input | `$#,##0_);($#,##0)` |
| Month header dates | `[$-en-US]mmm-yy;@` |
| Check rows (3 decimals, exposes tiny breaks; italic red) | `$#,##0.000_);($#,##0.000);-;` |
| Rate-spread inputs (renders e.g. "S + 800") | `"S + "000` |

Conventions baked in: negatives in parentheses; zeros render as `-`; `$` sign only
on the first row of each block and on totals; units stated once per tab.

## 6. The universal sheet chassis

Every schedule tab follows this skeleton. **Adapt the proportions and the row count
to the target — keep the pattern, never the literal addresses.**

### Column skeleton
| Col | Role | Width (approx) |
|---|---|---|
| A | narrow empty margin | ~12.5pt (≈1.8) |
| B | narrow — bold red "X" markers next to each banner (nav anchors); on the IS also helper tags ("Cash" / "Non-Cash") used as SUMIFS criteria | ~12.5pt |
| C | wide primary label column | ~141.5pt (≈26) |
| D | secondary label / spillover (same width). Indent with leading spaces or alignment, **not** by changing column widths | ~141.5pt |
| E onward | uniform data columns | ~53.5pt (≈10 for monthly) |

If the target already puts labels in col A, **adapt to the existing skeleton** —
do not insert columns unless the user approves restructuring.

### Header rows (the example uses rows 1–8; adapt the count, keep the order/roles)
1. Empty top row.
2. "Print Toggle" / "Circ Toggle" labels with values beside them — yellow `FFFFCC` fill, blue box, centered. On the Assumptions tab the values are blue-font hardcodes (0/1); on every other tab they are green-font links `=Assumptions!$D$2` / `=Assumptions!$D$3`.
3. Year helper row: `=YEAR(first date cell)` per column, plain gray (SUMIFS key for annual columns).
4. FY band row: label like "FY2019" only in the first month-column of each year — white bold text on a dark fill spanning the whole year's columns.
5. Period-date row: `=EOMONTH(prev,1)` chained from a single anchor date; format `[$-en-US]mmm-yy;@`; bold; centered; lighter companion fill.
6. Spacer row ~1.5pt tall (a hairline).
7. "Actual" / "Forecast" row — italic, centered, no fill.
8. First section banner row — navy `002855` across label + all data columns, white bold, red bold "X" in col B.

### Header fill pairs by tab family → (FY band fill, date-row fill)
| Tab family | FY band | Date row |
|---|---|---|
| Core statements (IS, BS, CF) | `525766` | `B6BAC5` |
| Liquidity / borrowing base | `24B1B1` | `9CEBEB` |
| Ancillary forecasts (Sales, COGS, BS Forecast) | `BCBFC6` | `E3E4E8` |
| Assumptions "Live Case" block | `6C1E36` | `DF8AA4` (with "Live Case" white bold over the maroon) |

### Freeze panes
Freeze the header rows + label columns on every schedule (the example freezes rows
1–8 and columns A–D → `freeze_panes = "E9"`). On long monthly tabs where history is
closed, freeze so the live forecast months are visible on open.

### Time axis
Monthly columns from model start to end; **Actual** through the last closed month,
**Forecast** after. Then a 1-column separator holding a literal lowercase "x" in
each header row, then an **Annuals** block (one column per year) built with
`SUMIFS(data_row, year_helper_row, annual_year)`. On liquidity-style tabs, closed
historical month columns get `F2F2F2` fill down the body to signal "closed"; the
current/live column stays white with blue inputs.

> Annual-only / quarterly models: there is no monthly axis. Use a single annual (or
> quarterly) block — no "x" separator, no SUMIFS annuals. Quarterly uses the same
> chassis with quarterly EOMONTH stepping.

## 7. Workbook architecture

- **Divider tabs**: empty tabs named with a trailing `>` ("Statements>", "Output>",
  "Sources>") act as section dividers — no tab color, no formatting.
- **Tab color families per section** (6-digit RGB):

  | Section | Tab color |
  |---|---|
  | Cover / Assumptions / dividers | *(none)* |
  | Core statements (IS / BS / CF) | `0067A5` |
  | Ancillary forecasts | `BCBFC6` |
  | Carveout / special analyses | `24B1B1` |
  | Output tabs | `FFFF00` |
  | Source data | `00B0F0` |
  | Working tabs | `00B050` |

- **Hide** helper/system tabs (e.g. a linking-names tab). The script does not change
  tab visibility on its own — leave that to the reviewer.
- **Deliberately do NOT format working tabs** (bridges, scratch dumps, raw source
  paste tabs): default font, gridlines on, minimal styling. A tab color is still
  applied, but the body is left untouched. Formatting effort is reserved for tabs an
  outsider will see. See `tab-patterns.md`.

## 8. Page setup (presentation tabs)

Portrait, Letter, margins ≈ **0.7"** left/right and **0.75"** top/bottom, print
gridlines off, per-tab custom print scale chosen so the label columns plus the
relevant period fit one page wide (the script sets fit-to-width = 1 page as a safe
default). Cover: portrait, ~50% print scale.
