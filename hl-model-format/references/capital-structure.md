# Capital Structure tab — the canonical playbook

A Capital Structure tab is **always** built like this (reference: DB Deep Dive, the
designated standard). When you see a cap-structure / cap-table tab — a debt-tranche
table with leverage, maturities, and a liquidity build — reproduce this exactly,
adapting only the tranche names and figures. All values below are the HL house theme
(`dk1 525766`, accents `508BC9/BCBFC6/7E8597/24B1B1/6C1E36/9FC3DA`); fills are theme
tints. Body font **Segoe UI 8**; gridlines off; portrait, fit ~one page (scale ≈47%),
margins 0.7"/0.75"; tab color accent2 `BCBFC6`.

## Layout / chassis
- **Col A**: narrow margin (~1.3). **Col C**: wide label column (~28). **Cols D–N**:
  the tranche-table data columns (~8.7 each). **Col S**: narrow gap (~1.7). **Cols
  P–R**: a small right-side effective-rate calc block. Row heights ~15 for data rows;
  thin **spacer rows** (3–5 pt) separate sections.
- Indentation is real (`alignment.indent`), not spaces.

## Title block (rows ~2–7), Segoe UI 8
- **Print/Circ toggle**: label **bold** gray; the value cell is an input — **yellow
  `FFFFCC` fill + a DASHED blue `0000FF` box, centered, blue font**.
- **Company name**: **bold** gray. **Subtitle "Capital Structure"**: *italic* gray,
  with a **medium `BCBFC6` bottom rule** under the title block.
- **"Data as of …"**: gray. **"($ millions)"**: *italic* gray (unit note — gray, not
  purple).
- **Ticker** (e.g. "NASDAQ:PLAY") and **as-of date** (`=TODAY()`): inputs → yellow
  fill + dashed blue box, centered; date format `mm-dd-yy`.
- **"DRAFT" / "CONFIDENTIAL"**: **bold `C00000` red, right-aligned**, sharing the
  medium bottom rule.

## Column headers (two rows), Segoe UI **Semibold** 8, gray, centered
A group-label row over a sub-label row; the **second header row carries a medium
`BCBFC6` bottom rule**. The standard columns (left block):

| Col | Header | Number format |
|---|---|---|
| C | Description | text, left, indent 1 |
| D | Face Value | **first tranche** `"$"#,##0_);("$"#,##0);"$"-_)`; **rest** `#,##0_);(#,##0)` |
| E | Leverage (x) | `#,##0.0\x ;"NMF"_x` |
| F | Interest Rate (spread) | `"S+"0`  → renders "S+325" |
| G | Effective Rate | `0.0%` |
| H | Implied Yield | `0.0%` |
| I | Annual Interest | `"$"#,##0` first / `#,##0` rest |
| J | Maturity | `[$-409]mmm-yy;@` |
| K | Credit Rating | text, centered |
| L | Trading Price | `0.0` |
| M | Market Value | `"$"#,##0…` first / accounting `_(* #,##0…)` rest |
| N | Leverage (Market) | `#,##0.0\x ;"NMF"_x` |

(Number format follows the column's unit — this is the per-column rule; use
`column_numfmt`. The "first row of the block gets the $" convention applies within
the $ columns.)

## Debt table → totals (the rollup, boxed)
Tranche rows: label in C (indent 1), figures per the column formats above. Then:

- **Total Debt** = sum of tranches → **bold**, fill **`F2F2F2`/`F2F2F4`** (subtotal),
  `"$"#,##0` value.
- **Less: Cash and Equivalents**, **Less: Debt Issuance Costs/Discounts** → indent-1
  adjustment lines.
- **Net Debt** = Total Debt − cash → **bold**, fill **`DCE8F4`** (major).
- **Market Cap** (links to a shares calc) → plain line.
- **TEV** = Net Debt + Market Cap → **bold**, fill **`D3EFEF`** (headline).

Every total row is **boxed**: thin `BCBFC6` **top + bottom across the row, left on the
first cell, right on the last** (a closed box around the row's span). This is the
cap-structure tier ladder: subtotal `F2F2F2` → major `DCE8F4` → headline `D3EFEF`,
all bold + boxed.

## Liquidity block (below a spacer)
**RCF Commitments** (bold) → **Less: Outstanding Borrowings**, **Less: Letters of
Credit** (indent 1) → **RCF Availability** (bold, fill `F2F2F2`, boxed) → **Plus:
Cash and Equivalents** → **Total Liquidity** (bold, fill `DCE8F4`, boxed).

## Memo & footnotes
- **Memo: LTM … Adj. EBITDA** (the leverage denominator) → its own row in a **dotted
  `BCBFC6` box**, gray, value `"$"#,##0`.
- **Footnotes** at the bottom (amort schedule, pricing source) → Segoe UI 8 gray,
  plain; referenced by `(1)`/`(2)` tags next to the relevant headers/labels.

## Right-side effective-rate block (cols P–R), Segoe UI 10
A small SOFR + spread → effective-rate calc: **SOFR** = blue `0000FF` input,
`0.000%` (3-dp); **Spread** (=F/10000) gray; **Effective Rate** (=SOFR+spread) gray.
Styled as an exhibit: **thick WHITE left/right borders** as column separators + a
**medium `7E8597` top rule**.

## Color convention on this tab (important — and the one nuance to confirm)
- **Gray `525766`** — tranche names, **contractual face amounts you type**, credit
  ratings, and the book/face totals (Total Debt/Net Debt/TEV face `$`), all labels.
- **Blue `0000FF`/`0000E1`** — pure model inputs only: the toggles, ticker, as-of
  date, and the SOFR rate.
- **Maroon `6C1E36` (accent5)** — the **market / as-of-date analytics refreshed each
  period**: leverage multiples, interest spreads, effective & implied-yield rates,
  annual interest, maturities, trading prices, market values, market cap, and
  balances that move (the drawn RCF). It marks "update these from the market/filings
  at each refresh," distinct from blue model inputs and gray static structure.

> This maroon-for-market-data convention is strong and deliberate in the canonical
> file but was **not** present in the older cap structures (Bingo/Traeger used
> blue inputs + gray only). Treat it as the current standard; if the intent differs,
> it's the one rule to adjust. Everything else here is unambiguous.

## How the engine produces this
Use Mode B with: `column_numfmt` for the per-column formats; row classes
`subtotal`/`major_total`/`headline_total` for Total Debt / Net Debt / TEV (palette
tier1 `F2F2F2` → tier2 `DCE8F4` → tier3 `D3EFEF`); `section_header`/labels with
indent for tranches and adjustments; input blocks (yellow + dashed-blue box) for the
toggle/ticker/date and the SOFR cell; `total_border` `BCBFC6`. Set a palette with
`body=525766`, `flag=C00000`, `font_body="Segoe UI"`, `font_size=8`, and use accent5
`6C1E36` for the market-analytics cells.
