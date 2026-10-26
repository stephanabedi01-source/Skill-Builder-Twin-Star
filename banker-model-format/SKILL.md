---
name: banker-model-format
description: Build and format Excel financial models in the user's banker house style (the style of the Project Print / Xerox restructuring model) - Garamond 10, navy section bands with x markers, grey column headers, blue hard-codes on a pale-yellow input fill, green cross-tab links, black formulas, light-blue key totals, en-dash zeros, a Cover Page with color legend, case selector and linked contents, section divider tabs, notes in cell comments, check rows and an End marker on every tab. Use whenever the user wants to build, lay out, extend, clean up or format a model or a tab in "my format" / "my house style" / "the way the Xerox model looks", including real estate models (acquisition, development, value-add, pro forma, rent roll, debt, waterfall, returns), three-statement, DCF, LBO, liquidity, liquidation, recovery waterfall and capital structure tabs, or asks how another chat should format a model - even if they do not name the skill.
---

# Banker Model Format

Build models that a reader can follow top to bottom without asking what anything is: one assumption in one cell,
every input visibly an input, every link visibly a link, every block closing on a total, every source in a
comment, nothing on the page that does not drive a number the deck uses.

## Generalization first: the values are evidence, the discipline is the rule

Everything concrete here (Garamond, `002060`, `FFF2CC`, column widths, row numbers, tab names) was read from one
workbook: the Project Print model (Xerox RX, v18), whose tabs the user built and Claude extended to match. Treat
those values as the default for a NEW model in this user's style. When the user hands over a workbook that already
has its own format, that workbook wins: read its fonts, fills, number formats, widths and layout and reproduce
them exactly (the `hl-model-format` skill's Mode A clones an established format losslessly). Never repaint an
existing model with these constants.

## The user's rules (from their feedback on the Xerox build)

1. **Easily followable, readable, clean.** "In all of the tabs that I have built ... the model is very easily
   followable and readable and clear and clean." If a tab needs explaining, simplify it: fewer rows, one flow,
   plain labels. Simpler means easier to read, not less analysis.
2. **Notes go in comments, not columns.** A "notes" or "source" column was moved to cell comments on every tab.
   Exception: a displayed `Source` column the user explicitly asks for (e.g. a precedent deal list that will be
   pasted into an appendix).
3. **No summary blocks that repeat the tab.** "Just link that to the operating assumptions and not have that
   summary thing at the bottom. That is unnecessary." Link to the tab that owns a number; never rebuild it.
4. **Match the user's own tabs exactly.** "Some of it does not match my perfectly formatted ones ... it is a little
   sloppy in those places." Copy styles from the user's best tab rather than approximating them.
5. **Debt schedules in the user's layout** (see references/tab-anatomy.md, "Schedule by year"), not a bespoke
   structure.
6. **Garamond everywhere**, including model outputs pasted into decks.
7. **Real, sourced numbers only.** Every hard-code carries its source in a comment; estimates are labelled as
   estimates; never invent a figure.
8. **Build with the user, not around them.** Check their existing work, improve assumptions, add notes, correct
   errors, and say what you changed and why. Ask before restructuring a tab they built.

## Workflow

1. **Plan the tabs before building.** List the tabs in order (inputs, then calculations, then outputs), what each
   one owns, and which tab feeds which. One tab owns each number; every other tab links to it. Get the user's
   approval on the tab list for a new model.
2. **Build the skeleton with `scripts/house_style.py`.** `new_workbook()`, `cover_page()`, `divider()` for each
   section, then a `Tab` per sheet: `header()`, optional `case_row()`, `band()` per block, `col_header()` /
   `year_header()`, then `row()` lines, `check()` lines and `end()`. Never type a hex code or number format by hand.
3. **Fill the formulas.** Inputs live on the assumptions tab (blue on the input fill). Calculations reference
   them; nothing typed inside a formula except structural constants (12 months, 100, 365).
4. **Comment every hard-code and judgment call** with `note()`: source, date, and the reasoning if it is an
   assumption.
5. **Add checks** where totals must tie (balance sheet balances, sources = uses, units tie to the rent roll,
   waterfall distributes 100%). `Tab.check()` turns red when not zero.
6. **Recalculate and look.** Run LibreOffice recalc (the xlsx skill's `recalc.py`): zero errors. Render the key
   tabs (`banker-deck-builder/scripts/render_xl.py`) and look at them: alignment, clipped labels, empty gaps,
   stray fills.
7. **Audit:** `python3 scripts/audit_model.py model.xlsx`. Fix every finding or say why it stays (a displayed
   source column, a typed counter row).
8. **Deliver:** `finalize(wb)` (Garamond default, gridlines off, every tab at A1, 85% zoom, Cover Page active).
   Report what was built or changed, what is an estimate, and anything the user must confirm.

## House style at a glance (full spec: references/house-style-spec.md)

| Element | Rule |
|---|---|
| Font | Garamond 10 everywhere (cover subtitle 12 italic grey) |
| Tab top | A1 `=C3` hidden (`;;;`, grey); B1 link to Cover; C3 title bold with a thin bottom rule across to the last column, entity right-aligned on the same row; C4 units in italics |
| Section band | Navy `002060`, white bold title in C, italic white text centered across (`centerContinuous`); "Projected" part in `215F9A`; side blocks in `A6CAEC`; "x" in A and B on the band row |
| Column header | Grey `E8E8E8`, bold, centered, thin bottom rule on the last header row, 3pt spacer row below |
| Line items | Label indented 1; first line of a block bold (no indent); % / memo lines italic indented 2 |
| Inputs | Blue `0000FF` font on `FFF2CC` fill (drivers the user changes); sourced historicals blue, no fill |
| Formulas | Black; links to other tabs green `008000`; checks red when non-zero |
| Totals | Bold over a thin top rule; key totals (Total Debt, Net Debt, NOI, Levered CF) bold on `DCEAF7` with top and bottom rules |
| Numbers | `$` on the first row of a block and on totals, plain numbers between; negatives in parentheses; zero shows an en dash; one decimal for % and multiples |
| Geometry | A, B and the right margin column 1.73 wide; labels 40 to 60; data 13; rows 12.75 |
| Close | x markers in A and B the row above `End` in column A |
| Workbook | Cover Page first; "Section >" divider tabs; tab colors shade from navy to grey by section; no gridlines |

## Scripts

- `scripts/house_style.py`: constants, number formats, `Tab` builder, `cover_page`, `divider`, `note`,
  `finalize`. `python3 house_style.py --selftest out.xlsx` writes a one-tab specimen of every row type.
- `scripts/audit_model.py model.xlsx`: font, color-coding, typed constants inside formulas, number formats, long
  text that belongs in a comment, layout markers, delivery state.
- `scripts/example_re_model.py out.xlsx`: a complete small multifamily acquisition model (assumptions with a case
  selector, unit mix, 10-year pro forma, loan schedule, returns with an exit-cap sensitivity) built only with the
  library. It recalculates with zero errors and audits clean: copy its patterns.

## References

- `references/house-style-spec.md`: every color, font, border, number format and geometry value, with the Xerox
  cell it came from.
- `references/tab-anatomy.md`: row-by-row layouts for the standard tab, time series, Low / Mid / High matrices,
  assumption blocks, schedules by year, sensitivities, the Cover Page and divider tabs.
- `references/build-principles.md`: how to keep a model followable, formula conventions, checks, case toggles,
  and the lessons log from the Xerox build.
- `references/real-estate-model.md`: tab blueprint and conventions for real estate models in this style.

## Delivery checklist

- [ ] Recalculates with zero errors; every check row reads an en dash (zero)
- [ ] `audit_model.py` clean, or each remaining finding explained
- [ ] Every hard-code blue with a source comment; every cross-tab link green; no typed constants in formulas
- [ ] No notes columns, no duplicate summary blocks, no orphan rows or stray fills
- [ ] Every tab: title row, units line, bands with x markers, End marker, gridlines off, opens at A1
- [ ] Cover Page: data-as-of date, sources, case selector (if cases exist), working contents links
- [ ] Estimates flagged in comments and in your message to the user
