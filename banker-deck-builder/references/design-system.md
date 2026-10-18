# Design system

Values in this file come from the Project Print deck (Garamond, near-black ink, red accent). They are evidence of a
system, not constants. For a new deck, keep the ROLES and RULES, and take the actual colors, font and geometry from
the user's template and theme.json.

## Contents
1. Page anatomy and grid
2. Typography
3. Color roles and accent discipline
4. Components (what each element is for)
5. Charts: visual rules
6. Tables: visual rules
7. Density and balance
8. Anti-patterns (what makes slides look AI-made)
9. Reference deck lessons

---

## 1. Page anatomy and grid (16:9, 13.33 x 7.5 in)

Top to bottom on every content slide:
1. **Title** (template placeholder): short noun phrase, e.g. "Valuation Summary". Accent vertical tab at its left
   comes from the master.
2. **Takeaway band** (template placeholder, italic, light grey band): ONE line, about 100 to 120 characters, stating
   the slide's claim with numbers.
3. **Body**: two panels by default.
   - Left panel: x 0.23 to 8.33 (main exhibit, ~60%).
   - Right panel: x 8.62 to 13.11 (supporting exhibit + table or callout, ~35%).
   - Gutter ~0.29 in. Panel headers sit at y 1.30 with a rule at y 1.59; content starts ~1.66 (caption) / ~1.9.
   - Alternatives: 2x2 modules (situation-style), full-width table (debt table) + side panel, or three columns
     for appendix pages (image + image + commentary).
4. **Closing line** (y 6.43, h 0.36): full-width dashed rounded box, one sentence, the conclusion + hand-off.
5. **Footer** (from the master, top ~6.93): rule, logo, Notes line(s) and Sources line in grey 10 pt, page mark.
   Notes are numbered "(1) ...; (2) ..." matching superscripts placed exactly on the numbers they explain.

The body must end above ~6.30 so the closing line and footer breathe.

## 2. Typography

- One font throughout (the deck's: Garamond here). Set it on every run and in chart text (txPr), because theme
  fonts are often Office defaults.
- Sizes: panel header 14 bold; body/table 10.5; chart labels 10 to 10.5 bold; captions 10 italic grey; category
  labels 9 to 9.5; footnotes 10 (master); closing line 11.5; takeaway as the template sets it.
- Bold for the key phrase or number in a sentence, not whole sentences.
- Italic grey for captions ("EV / LTM Adj. EBITDA", "Recovery at mid EV (low to high)") and secondary notes.
- Superscript footnote markers right after the number or label they qualify.

## 3. Color roles and accent discipline

Roles (theme.json keys): ink (text, dark bars, rules), accent (emphasis text, lines, pill borders, markers),
accent_fill (bar fills for the subject/distress), grey_txt (captions), grey_mid (secondary series, connectors,
"lost value"), grey_lt (peer/reference bars), hair (table rules), band (shaded ranges, highlighted rows),
callout_fill (commentary callouts), white.

Rules:
- **The accent is a signal, not decoration.** Use it only for the subject company (its bar, its row, its label)
  and for distress or bad numbers (shortfalls, impaired recoveries, maturities that cannot be met). Brackets,
  supporting text and chevrons stay ink or grey. Six red chevrons on an agenda spends the signal.
- Peers grey, subject accent. Core methods ink, reference methods light grey.
- Status quo vs. proposal: status quo in accent (the problem), proposal in ink (the fix).
- Positive vs. negative: negative numbers in parentheses "($990)", accent only when they are the point.
- A deck has ONE accent. Do not introduce extra hues (no blue/green highlights) unless the brand demands it.

## 4. Components

| Component | Look | Use for | deck_kit |
|---|---|---|---|
| Panel header | bold 14 + full-width ink rule | every exhibit | `section_header` |
| Caption | italic grey 10 | metric/units under a header | `caption` |
| Commentary callout | grey fill (`callout_fill`), accent round-dot border, centered 10.5 | the "so what" sentence beside the evidence | `callout` |
| Pill | white fill, round-dot border (accent for bad news, ink for neutral) | flag on a chart point: "ABL matures May 2028", "+$857M preserved", "the fulcrum" | `pill` |
| Highlight row | `band` or translucent grey box behind a chart row/table rows, sometimes with accent dotted border | the subject company row in a peer chart; fulcrum classes in a table | `box`, `grid_table(highlight_rows)` |
| Outlined column | thin ink rounded rectangle around a table column | the base case / the recommended plan column | `grid_table(outline_col)` |
| Claim line | vertical dashed line across a value chart, labeled at the top | total debt, secured debt, market price | `line(..., dash=DASH)` |
| Band | `band`-filled rectangle behind bars | a selected range or a peer Q1 to median range | `box(... shape=RECTANGLE)` |
| Bracket | horizontal line with end ticks, or a brace | tie a range to its label: "Unsecured notes: the fulcrum", "Receivable runoff $663M" | `hbracket`, `brace` |
| Connector | thin grey/accent line from a tag to a bar label | deal tags over bars | `line` |
| Numbered items | dark circles with white numbers + bold lead phrase | 2 to 4 issues/considerations | `numbered_items` |
| Card | dark rounded header bar over white body box | entity charts, business-model boxes | `card` |
| Closing line | dashed ink rounded box, full width | the conclusion of every main slide | `closing_line` |
| Timeline | dots on a line, dates above, steps below; dashed accent segment to the deadline | path to emergence, process steps | `box(OVAL)`, `line` |

Callout vs. pill fill rule (agreed with the user): grey fill for commentary boxes that hold sentences; white fill
for pills that mark a point on a chart. Grey under small text hurts contrast and blurs the warning signal.

## 5. Charts: visual rules

- Native, editable PowerPoint charts with data in the embedded workbook (deck_kit.add_chart). Never pictures of
  charts on main slides.
- Hide axes and gridlines; label values directly. Draw your own baseline (grey_mid 0.75 pt).
- Data labels are written out with units: "$1,739M", "$16.7B", "5.9x", "36%", "52c". If labels wrap ("$1,739" /
  "M"), move units into the header ("($M)") and show bare numbers.
- Put labels where nothing else is: above bars, inside bars (white on dark) when a line or band crowds the top,
  right of horizontal bars. Labels must never sit on a rule, a dashed line or a marker. White-fill a label that
  must sit on a line.
- Logos instead of tickers in peer charts when the user has them (copy them from slides where the user already
  placed them). If one is missing, put the company name in its slot and tell the user.
- One message per chart. Add exactly the annotation that makes the message unmissable: a claim line, a band, a
  bracket, a "+$857M preserved" pill, a break line "Value breaks here".
- Category labels as textboxes under the plot (two lines: name bold / date or detail grey) give full control and
  avoid rotated or wrapped axis labels.
- Truncate outliers with a break mark and footnote rather than squashing every other bar.

## 6. Tables: visual rules

- No PowerPoint table styles. Either the textbox grid (`grid_table`) or a fully restyled native table
  (`style_cell`) for dense data like a debt table.
- Header row bold with an ink rule under it; hairline rules between rows only when rows are dense; a heavier ink
  rule above totals; totals bold.
- First column left, numbers right. "($M)" in the top-left header cell.
- Accent only for the cells that carry the message (shortfall, impaired recovery). Outline the column the slide is
  arguing for.
- 9 pt is the floor for anything that must read on a recorded video.

## 7. Density and balance

- Two to four exhibits per main slide, each with one visual and very little text.
- No blocks of prose. A callout is one sentence; a commentary list is 2 to 4 numbered points.
- Fill the page evenly: no large empty quadrant, no exhibit crammed into a corner. If a panel is sparse, enlarge
  the exhibit or add the one supporting fact that strengthens the story (never filler).
- Repeat a strong visual language across slides (the same claim lines, the same highlight row, the same bar
  colors) so the deck reads as one system and each slide is faster to read.

## 8. Anti-patterns (look AI-made; avoid)

Letter-spaced all-caps kickers; thin accent stripes on box edges; big "01 / 02" numerals; rows of uniform stat
tiles; tiny grey labels; gradients; drop shadows on everything; perfectly flat generic shapes with no purpose;
icon soup; emoji; centered paragraphs of prose; a headline strip repeated on every slide; "Key Takeaways" boxes
that restate the title; overly symmetric layouts that ignore the data. Rebuilt "Excel-looking" tables in the
appendix look AI-made too: paste the real model outputs (appendix-model-outputs.md).

## 9. Reference deck lessons (from the user's banker and case-competition decks)

What the strongest decks share (Twelve Oaks, Whistler, Quinnipiac style): one-line italic takeaway in a grey band;
section header bars above each exhibit; two to four exhibits per slide with superscript footnotes tied to one
Notes line; dark rounded callouts for the one number that matters; dashed takeaway boxes beside the evidence;
logos in place of tickers; football fields with reference lines; recommendations built from scenario columns or
lettered theses with chevrons into catalysts.

What fails: overcrowding (five or six exhibits with colliding labels); blocks of prose (memo-in-slides); rows of
identical stat tiles; inconsistent fonts; overlapping chart labels; unscrubbed leftovers (placeholder slides,
joke footnotes). Always do a final scrub pass of notes and placeholders.

Last year's competition winner (restructuring case): entity-level precision, full debt tables, liquidity runway,
pro forma comparisons, and an appendix of pure model outputs (revenue build, DCF assumptions by case,
post-transaction waterfall, sources and uses, liquidity builds). Beat it with real charts and a clear story; match
it on precision.
