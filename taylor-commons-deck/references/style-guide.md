# Style guide: Taylor Commons AM deck

The slide master already does the header, footer, logo and colors. This file describes what is on the master
and the conventions for everything you put on a slide.

## Contents
1. Master (what you get for free)
2. Page grid
3. Type scale
4. Colors
5. Components
6. Charts and tables
7. Photos and maps
8. Writing
9. Checklist before handing a slide back

## 1. Master

- **Cover (TITLE layout):** solid charcoal background, white Garamond title, short brick rule under it, presenters
  below, fund name and month bottom left, white SMU logo bottom right. No footer.
- **Content slides (CONTENT layout):**
  - Thin charcoal bar across the very top edge.
  - Running head top right: "Taylor Commons | Asset Management Update", 9pt grey italic.
  - Title placeholder (idx 106): 28pt bold charcoal, left at 0.50 in, top 0.19 in.
  - Takeaway placeholder (idx 107): 14pt italic slate (`4A5560`), at 0.70 in; one line, a claim with numbers.
  - Plain grey hairline under the takeaway at 1.12 in, full width (no colored segment).
  - Footer: grey rule at 6.82 in; charcoal SMU logo bottom left; fund name and page number bottom right, 9pt.
- Always use the **CONTENT** layout for new slides. The other layouts (5_CONTENT, 6_CONTENT, ...) carry the prior
  group's pre-drawn panel rules and placeholder positions and will fight your layout.

## 2. Page grid (inches, 13.33 x 7.5 slide)

| Element | Position |
|---|---|
| Left margin / right margin | 0.50 / 12.83 |
| Left panel (main exhibit) | x 0.50 to 7.85 (7.35 wide) |
| Right panel (support) | x 8.20 to 12.83 (4.63 wide) |
| Panel headers | top 1.32; rule 0.29 below |
| Content starts | ~1.72 |
| Body must end by | ~6.0 |
| Closing line (dashed box) | y 6.22, height 0.36, full width |
| Footer notes and sources | x 1.45, y 6.87, width 8.5, 9pt, two short lines |

Two-panel is the default. Alternatives: a full-width table plus side panel, or a photo grid (3 x 2) on the left
with an assessment on the right (site visit slides).

## 3. Type scale (Garamond only)

| Use | Size |
|---|---|
| Title | 28 bold (master) |
| Takeaway | 14 italic slate (master) |
| Panel header | 15 bold with ink rule |
| Body text, tables, callouts, numbered items, closing line | 12 |
| Chart value labels | 12 bold |
| Chart category labels, captions, axis labels, photo captions | 11 (captions italic grey) |
| Footer notes and sources | 9 grey (sources italic) |

Bold the key phrase or number in a sentence, not whole sentences. Superscript footnote markers go right after the
number they qualify: `$13.7M(2)`.

## 4. Colors

| Role | Hex | Use |
|---|---|---|
| Ink | `22262B` | titles, text, dark bars, rules, totals |
| Slate | `4A5560` | takeaway line |
| Accent (brick) | `B4442C` | Taylor Commons only, and bad news: losses, shortfalls, "Weak" grades |
| Grey text | `5F6670` | captions, sub-labels, footnotes |
| Grey mid | `8C939A` | baselines, connectors, secondary series |
| Grey light | `C9CDD1` | comparison bars (IC plan, peers, exemption-kept case) |
| Hairline | `D5D8DB` | table rules |
| Band / callout fill | `F2F1EE` / `EEEDEA` | highlighted rows, callout boxes |

One accent per deck. In a table, color only the cells that carry the message (e.g. 4 cells, not a whole column).

## 5. Components

- **Panel header:** bold 15pt, full-width ink rule under it. Optional superscript note marker.
- **Callout:** light grey fill, brick round-dot border, one sentence, centered, 12pt. Use at most one per panel.
- **Pill:** small rounded white label with dotted border, flags one point on a chart ("Ends at YE2026 under HB 21").
- **Closing line:** full-width dashed box at the bottom: the finding plus the hand-off to the next slide, with the
  key clause in bold. One line.
- **Numbered items:** dark circles with white numbers, a bold lead phrase, then a short line. Two to four items.
- **Assessment rows (site visit):** area name bold left, grade right ("Weak" in brick, "Fair"/"Good" in ink), one
  or two lines of finding beneath, hairline between rows.

## 6. Charts and tables

- Native PowerPoint charts with data in the embedded workbook. Axes hidden; draw your own baseline; label values
  directly ("$1.43M", "0.93x", "(7.0)"). Category labels as text boxes below the plot, bold name plus an italic grey
  sub-label.
- Waterfalls: invisible base + increases + decreases; dashed grey connectors at running totals.
- Claim lines: dashed line across the chart for the number that matters ("$16.5M owed, due Mar-27", "IC stabilized
  occupancy 95%"), labeled where nothing else is.
- Tables: no PowerPoint table styles. Bold header row with an ink rule, hairlines between rows, a heavier rule above
  totals, first column left, numbers right.
- Never let a label sit on a line or marker; move it or put it inside the bar.

## 7. Photos and maps

- Photos are cropped to exact boxes (the `photo` helper does center crops; `focus` shifts the crop vertically).
  **Photo rule (user instruction, Nov-18): use ONLY the team's Nov-26 secret shop photos (Seth's; `scripts/img/v26_*.jpg`,
  also on the seven "Secret Shopping - Asset Photos" slides). Never use photos from the Nov-25 inspection report, the
  IC deck, the sponsor, prior groups or past assignments.** Maps and logos are not photos and are allowed.
  Captions credit "Team site visit and secret shop (Nov-26)".
- Maps use clean pins, not logo collages: brick dot plus label for Taylor Commons, dark numbered circles for
  competing apartments, white circles (S1..) for schools, grey circles (R1..) for retail, and a matching key in
  the right panel. Illustrative overlays (county outline, migration arrows) are labeled "Illustrative" and
  footnoted as approximate.

## 8. Writing

- Takeaway: one line, about 100 to 120 characters, a claim with numbers, not a topic ("The HFC refinancing added
  $1.04M of annual interest to save $428K of taxes net of fees, a $612K yearly loss").
- Closing line: finding plus hand-off ("... HB 21 ends the exemption at YE2026 while the 9% debt remains").
- Third person only; no we/our/you on slides. No em dashes; ranges read "$13.4M to $15.6M".
- Name every figure's period (T12 to Feb-26, Nov-26 box score, 2027 base case) and basis (with or without the tax
  exemption, levered or unlevered).
- Say what is true, not what a source says: if a number is an estimate or from a secondary source, the footnote
  says so.

## 9. Checklist before handing a slide back

- [ ] Built on the newest deck; saved as a new version.
- [ ] CONTENT layout; title and takeaway in the placeholders; slide number shows.
- [ ] Takeaway and closing line each fit on one line.
- [ ] Body text 12pt (10pt minimum); nothing overlaps; nothing runs past 6.0 in.
- [ ] Brick used only for Taylor Commons and bad news.
- [ ] Every number matches `deal-facts.md` and other slides; footer notes and sources filled in.
- [ ] Speaker notes: spoken script plus reference notes.
- [ ] `lint_deck.py` clean for your slide (pronouns, dashes, placeholders).
