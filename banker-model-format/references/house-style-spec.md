# House style specification

Read from the Project Print model (Xerox RX, v18). The user built the financial statement, NWC, capital structure,
maturity and comps tabs; Claude built the analysis tabs by copying cell styles from the user's
`$XRX DCF Operating Assumptions` tab. All values are mirrored in `scripts/house_style.py`; use the constants
there rather than retyping anything below.

## Contents
1. Theme and colors
2. Fonts
3. Cell roles (what each kind of cell looks like)
4. Number formats
5. Borders, alignment, geometry
6. Tab colors and views

---

## 1. Theme and colors

The source workbook uses an Office theme with `dk2 = 0E2841` (navy) and `lt2 = E8E8E8` (light grey), and paints
many fills as theme tints. A new workbook built with openpyxl gets a different default theme, so the library uses
the resolved hex values:

| Constant | Hex | Source in the Xerox model | Use |
|---|---|---|---|
| `NAVY` | `002060` | band cells, e.g. Capital Structure C7 | Heading 1: section band |
| `GREY_H2` | `808080` | Cover legend "Heading 2" | secondary band (rare) |
| `HDR` | `E8E8E8` | theme lt2, header rows (Capital Structure row 8 / 9) | column headers, memo rows |
| `HDR_DK` | `D1D1D1` | lt2 -10%, comps header row 6 | dense headers on wide tables |
| `PROJ` | `215F9A` | dk2 +25%, DCF Op Assumptions I8 "Projected" | projected part of a band |
| `LTM` | `4E95D9` | dk2 +50%, "LTM" header cells | LTM / stub column header |
| `ASIDE` | `A6CAEC` | dk2 +75%, DCF "Exit Multiple Method", sensitivity headers | side blocks, method sub-bands |
| `KEY_TOTAL` | `DCEAF7` | dk2 +90%, Capital Structure "Total Debt", "Net Debt" | key totals |
| `INPUT` | `FFF2CC` | Liquidation % recovery inputs, Cover "Input/Driver" | drivers the user changes |

Font colors (the Cover Page legend, rows 13 to 21):

| Color | Hex | Meaning |
|---|---|---|
| Black | `000000` | formulas on the same tab |
| Blue | `0000FF` | hard-codes: typed inputs and historical data |
| Green | `008000` | links to other sheets |
| Red | `FF0000` | highlight values / checks |
| Grey | `A0A0A0` | A1 title echo, divider-tab labels, Cover "Return to Cover" |
| Hyperlink | theme hlink / `0563C1`, underlined | B1 link back to the Cover Page |

Output tabs (`Income Statement Output`, `Balance Sheet Output`, ...) are presentation pages that only restate the
working tabs: the user keeps them black throughout. Working tabs color-code every link green.

## 2. Fonts

- Garamond 10 for every cell (12,655 of 12,684 populated cells in the source). Bold for titles, band titles,
  header rows, first lines of blocks and totals. Italic for units lines, band sub-text, % / growth / margin lines,
  memo and check lines.
- Cover Page subtitle: Garamond 12 italic, grey `808080`.
- B1 link: Garamond 11, underlined, hyperlink color.
- Set the workbook's default font to Garamond 10 too (`new_workbook()` does), so blank cells and new rows match.

## 3. Cell roles

| Role | Font | Fill | Border | Default format | Align / indent |
|---|---|---|---|---|---|
| `label` (line item) | black | - | - | - | left, indent 1 |
| `label_head` (first line, e.g. Revenue) | bold | - | - | - | left, 0 |
| `label_sub` (% Growth, % Margin) | italic | - | - | - | left, indent 2 |
| `label_total` | bold | - | thin top | - | left, 0 |
| `label_key` | bold | `DCEAF7` | thin top + bottom | - | left, 0 |
| `label_memo` ("Memo: LTM Adj. EBITDA") | bold | `E8E8E8` | hair top + bottom | - | left, 0 |
| `label_check` ("Check: ...") | italic | - | - | - | left, indent 2 |
| `input` | blue | `FFF2CC` | - | num | right |
| `hard` (historical / sourced) | blue | - | - | num | right |
| `link` | green | - | - | num | right |
| `calc` | black | - | - | num | right |
| `calc_sub` (%) | italic | - | - | pct | right |
| `total` | bold | - | thin top | dol | right |
| `key` | bold | `DCEAF7` | thin top + bottom | dol | right |
| `memo` | bold | `E8E8E8` | hair top + bottom | dol | right |
| `check` | italic, red when non-zero | - | - | num | right |

`Tab.row()` picks the font color from the value, so a row can mix them: a typed number is blue, a formula that
reads another tab is green, any other formula is black. A total or key row keeps its bold and fill whatever the
value.

## 4. Number formats

Copy these strings exactly (they are `NF[...]` in the library). The pattern: positive with a trailing space for a
parenthesis, negative in parentheses, zero as an en dash, text aligned with the numbers.

| Key | Format string | Shows | Use |
|---|---|---|---|
| `num` | `_(#,##0_)_%;\(#,##0\)_%;_("–"_)_%;_(@_)_%` | 1,234 / (1,234) / – | most rows |
| `dol` | `_([$$]#,##0_)_%;\([$$]#,##0\)_%;_("–"_)_%;_(@_)_%` | $1,234 | first row of a block, totals |
| `num1`, `dol1` | one-decimal versions | 77.4 / $1,234.5 | bond prices, $ per SF |
| `dol2` | two decimals | $1.25 | per-share, rent per SF |
| `pct` | `_(#,##0.0%_);\(#,##0.0%\);_("–"_)_%;_(@_)_%` | 4.5% | growth, margins, recoveries |
| `pct2` | two decimals | 5.75% | rates, cap rates, input % |
| `pct3` | three decimals | 10.250% | coupons, SOFR |
| `mult` | `_(#,##0.0"x"_)_%;\(#,##0.0"x"\)_%;_("–"_)_%;_(@_)_%` | 5.3x | leverage, EV multiples |
| `mult2` | two decimals | 1.25x | DSCR, equity multiple |
| `year_a` / `year_e` | `General"A"` / `0"E"` | 2025A / 2026E | year headers (numbers, so `=prev+1` works) |
| `year_n` | `"Year "0` | Year 3 | hold-period headers |
| `month` | `mmm-yy` | Mar-28 | maturities, monthly headers |
| `date` | `m/d/yy;@` | 6/30/26 | dates |
| `spread` | `"S + "#` | S + 400 | floating-rate spreads |
| `hidden` | `;;;` | (nothing) | A1 title echo |
| toggles | `[=1]"Bull";[=2]"Base";"Bear"` | Base | numeric case index shown as a name |

Use `$` on the first number row of a block and on totals; plain numbers in between (Maturity Schedule rows 9 and
16, Liquidation rows 11 to 23). One decimal for percentages unless the input itself needs two.

## 5. Borders, alignment, geometry

- Thin rules only: under the title row (C3 across to the last column), on top of a band, under the last header
  row, on top of totals, top and bottom of key totals. Hair rules around memo rows. No vertical lines, no boxes
  except the case selector cell (thin all round).
- Bands: title left-aligned in C; any band text in a data column is `centerContinuous`, so it centers across the
  empty band cells to its right. Header labels centered. Numbers right-aligned (centered in the capital structure
  table, which is the user's own choice for that tab).
- Widths: A and B 1.73 (marker columns), label column C 40 to 60 (wide enough for the longest label, never
  wrapping), data columns 13, one 1.73 column closing the right margin.
- Rows 12.75; title row 14.25; a 3pt spacer row under a header row; one blank row between blocks.

## 6. Tab colors and views

- Cover Page tab: `D1D1D1`. Sections shade by depth: `0E2841` (financial statement inputs), `215F9A` (outputs),
  `4E95D9` (credit analysis), `A6CAEC` (analysis), `E8E8E8` (stock performance), `747474` (valuation). Divider
  tabs share their section's color.
- Gridlines off on every tab. Deliver every tab at A1, cursor at A1, Cover Page active. The source used Page
  Break Preview at 81 to 100%; `finalize()` sets Normal view at 85%, the delivery standard in the user's
  `hl-model-format` skill.
