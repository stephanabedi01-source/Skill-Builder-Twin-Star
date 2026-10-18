# QA checklist and delivery

Run this every time a slide is built or changed. The render loop is not optional: almost every approved slide went
through two or three rounds of render, look, fix.

## Commands
```bash
python3 <skill>/scripts/render_slides.py out/Deck_v13.pptx out/r13a --pages 10 10      # new folder each round
python3 <skill>/scripts/lint_deck.py out/Deck_v13.pptx
python3 <pptx-skill>/scripts/office/validate.py out/Deck_v13.pptx --original <user base deck>.pptx
```
Find the pptx skill's validate.py with a glob if the path is unknown (e.g. `**/pptx/scripts/office/validate.py`).

## Visual checks (look at the PNG)
1. Collisions: labels on header rules, on claim lines, on markers, on other labels; legend text over chart.
2. Wraps: takeaway on one line; closing line on one line; callouts without an orphan last word; category labels
   not breaking mid-word ("InnerWorkin / gs").
3. Overflow: text spilling out of boxes, cells, pills; labels cut off at the chart's right edge.
4. Values: every label matches the plan and the model; bars match labels (a "$3,175" anchor cell where the grid
   should start at "$2,602" means a stale data-table value).
5. Balance: no large empty quadrant, nothing crammed; exhibits aligned on the grid; consistent gaps.
6. Color discipline: accent only on the subject/distress; nothing decorative in accent.
7. Consistency with earlier approved slides: same header style, callout style, label style, number formats.
8. Readability on video: nothing important below 9 pt.

Font caveat: if the deck font is missing in the sandbox, renders use a wider substitute. Marginal wraps may not
happen in PowerPoint; say so when reporting, but fix anything clearly broken.

## Content checks
- Numbers tie to the model and to other slides; derived numbers recomputed (show the arithmetic in chat when
  useful).
- Every superscript has a note; every note is used; notes in order.
- Lint clean: no pronouns, no em/en dashes in prose, no placeholders, no "$(x)" negatives.
- Shapes named meaningfully.
- The slide sits at the right position; slide numbers in the footer update.

## Delivery
- Save as a new version (`_v14.pptx`), never overwrite the user's file. Keep the builder chain rerunnable.
- Send the file (and the new slide PNG when helpful).
- Report: what the slide shows and how it advances the story; changes to other slides; open items for the user;
  the next slide. Keep it short and plain.
- Before final submission: scrub every slide for leftovers (placeholder text, empty Notes labels, stale footnotes,
  hidden off-slide shapes), confirm page limits, export to PDF and look at every page once more.
