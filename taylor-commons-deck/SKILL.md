---
name: taylor-commons-deck
description: Build, edit, check and present slides for the SMU Cox Real Estate Impact Fund's Taylor Commons asset management (AM) update deck (224-unit, 1979 multifamily in east Fort Worth; TexSun sponsor; March 2025 Pecos HFC refinancing; HB 21 tax cliff; W Financial bridge loan). Carries the team's slide master and formatting rules (Garamond, charcoal and brick, 12pt body, takeaway line, dashed closing line, 9pt footer notes), every sourced deal number, the status of each slide, open questions for TexSun, and runnable builder scripts plus the team's Nov-26 property photos and maps. Use this whenever anyone mentions Taylor Commons, the TC deck, the AM update or AM pitch, the HFC or HB 21 slides, the refinancing trade, the 2027 cliff, valuation and recovery, rent or sales comps for Taylor Commons, the site visit slides, speaker notes for the presentation, or asks to add, fix, restyle, check or rebuild any slide in that deck, even if they do not say "skill".
---

# Taylor Commons AM Deck

A shared playbook so every teammate's slides look like one deck and quote the same numbers. It holds three things:
how slides look (the house style), what is true about the deal (sourced facts), and where the deck stands (what
is built, what is left, what is still open with the sponsor).

Presentation: Taylor Commons AM update to the fund, November 18 (the assignment brief schedules it then; deck and
model were due the Monday before at 5 PM CT).

## Start here

1. **Get the latest deck from the team** (the file name has been `v2_Taylor_Commons_AM_pitch_11.16.2026.pptx`).
   Always build on the newest file someone shared, never an older copy, and save your work as a new version
   (e.g. `v3_...`) so nobody's edits get overwritten.
2. **Read `references/slide-log.md`** to see which slides are finished (and should not be redone), which are still
   the prior group's leftovers, and the plan for the rest.
3. **Read `references/deal-facts.md` before quoting any number.** Every figure on a slide must come from there or
   from a named document, and must match how other slides already state it. If you find a better source or a
   conflict, update the facts file and tell the team which slides quote the old number.
4. **If you are editing the live deck through Claude in PowerPoint** (teammates editing at the same time), use the
   copy-paste batches in `references/claude-in-ppt.md`: house rules first, then one batch at a time, matching
   slides by title.
5. **Follow `references/style-guide.md`** for layout, fonts, colors and wording. The master already carries the
   header, footer, logo and colors, so new slides only need the body.

## The story the deck tells

One sentence: Taylor Commons' common equity is out of the money. Operations missed first (occupancy and NOI well
below the IC plan), then the March 2025 refinancing swapped a 3.57% agency loan for a 9.5% bridge to capture an HFC
tax exemption that saved less than it cost; HB 21 ends that exemption at YE2026 while the 9% debt stays, the loan
matures in March 2027, and value sits below the loan, so SMU's $350K recovers $0.

Arc: Asset and location -> Then vs. now (IC vs. today) -> Site visit -> Operations -> Market and comps ->
Refinancing trade -> HB 21 -> 2027 cliff -> Valuation and recovery -> What the fund does now (the team wants this
closing section to be creative but rigorous) -> Lessons.

Every new slide should push that argument one step and hand off to the next slide in its closing line.

## How to work on a slide

1. **Plan first, in a few lines:** the claim (takeaway with numbers), what goes in the left panel (main exhibit,
   ~60% width) and right panel (~35%), the closing line, and the numbers with sources. Do the arithmetic that
   tests the claim before designing anything; if the numbers do not support the claim, change the claim.
2. **Build it.** Two ways:
   - *In PowerPoint by hand:* insert a new slide on the **CONTENT** layout, type the title and the italic takeaway
     into the two placeholders, and follow the grid in the style guide.
   - *With the scripts (Python + python-pptx):* copy the pattern of the closest builder in `scripts/` (each one
     builds one slide; see "Scripts" below). Run it on the latest deck; it writes a new file.
3. **Check it.** Render it and look (`scripts/render_slides.py`, needs LibreOffice), or open it in PowerPoint and
   look at 100%. Fix collisions, wrapped one-line elements (takeaway, closing line, callouts leaving an orphan
   word), text over chart lines, and anything below 10pt in the body. Then run `scripts/lint_deck.py` for
   pronouns, dashes and placeholders.
4. **Write speaker notes** as a spoken script (presenter name on line one, plain talk), then a "Reference notes"
   block with the facts and sources the presenter may be asked about.
5. **Hand it back** as a new version with a two-line summary: what the slide shows and which numbers changed.

## Rules the team agreed on

- Garamond everywhere; body text about 12pt; nothing in the body under 10pt (footer notes are 9pt).
- Brick red (`B4442C`) only for Taylor Commons itself and for bad news (losses, shortfalls, weak grades). Everything
  else is charcoal ink and greys. Do not reintroduce SMU navy or the red corner swoosh; that template was used by
  another group.
- Notes and sources live in the footer band next to the SMU logo, short: one line of notes, one line of sources.
  Longer explanations go in speaker notes.
- No personal pronouns on slides (we, our, us, you, I); slides speak in the third person. No em dashes. Numbers
  formatted like `$1.43M`, `($612K)`, `0.93x`, `74.9%`, `$13.4M to $15.6M`.
- Photos: only the team's Nov-26 secret shop photos (`scripts/img/v26_*.jpg`). No photos from the Nov-25 inspection
  report, the IC deck, the sponsor, prior groups or past assignments. Maps and logos are fine.
- Real, sourced numbers only. Mark estimates as estimates. Never fill a gap with a plausible-looking number.
- Native, editable charts (not pictures of charts). Label values directly; no gridlines or legends unless needed.
- The prior group's facts are not trusted by default: their Asset Overview said "24 affordable units" (wrong; the
  regulatory agreement requires 50% of units under 80% AMI and 90% under 140%) and 178,923 SF (IC says 171,715 SF).

## Scripts (optional, for teammates using Claude or Python)

`scripts/` holds a small toolkit:
- `deck_kit.py`: building blocks (textboxes, panel headers, banker tables, native charts with exact geometry,
  callouts, closing line, pills, braces).
- `tc_common.py`: deck-specific helpers on top of it: `new_slide` (CONTENT layout, title + takeaway; if a slide with
  that title already exists it is replaced in place), `place`, `header` (15pt panel header), `closing` (12pt dashed
  closing line), `caption`, `footer_notes` (notes + sources in the footer band), `photo` (crops a photo from
  `scripts/img` to an exact box), `speaker_notes`, `add_slide_number`, `drop_slide_titled`.
- `theme.json`: palette and grid used by every builder.
- One builder per finished slide (`build_snapshot.py`, `build_operations.py`, `build_site.py`, `build_location.py`,
  `build_migration.py`, `build_refi_trade.py`, `build_hb21.py`, `build_cliff.py`, `build_valuation.py`). Each has
  a docstring listing every number and its source. `build_master.py` documents how the slide master was rebuilt
  from the original template (only needed if the master must be rebuilt from scratch).
- `render_slides.py` (PNG renders via LibreOffice) and `lint_deck.py` (writing-rule checks).

Usage: `python3 scripts/build_<slide>.py <latest_deck.pptx> <new_version.pptx> [position]`. Rerunning a builder on
a deck that already has that slide updates it in place. To make a new slide, copy the closest builder, change the
title, takeaway, exhibits, notes and sources, and keep the same helpers so spacing and fonts stay consistent.
Requirements: Python 3, `python-pptx`, `lxml`, `Pillow`; LibreOffice and `pdftoppm` for renders. Install Garamond
(or EB Garamond) for renders that match PowerPoint.

## Where things come from (the source files the team has)

Assignment brief; IC presentation, IC memo and IC model (3/18/23); prior AM models (v16/v17, May 2025) and the
current AM model (`2026.11.16 - TC_AM_Model_v1_1.xlsx`, whose Rent Roll tab holds the 11/15/2026 box score);
fund 2025-26 annual report; TexSun Q1-26 commentary; T12 income statement (Mar-25 to Feb-26), balance sheet
(Feb-26), rent roll (Apr-26), box score data (weekly, Jul-24 to Dec-25), trending occupancy, retention,
receivables and trade-out reports; the March 2025 closing documents (special warranty deed, memorandum of ground
lease, fee and leasehold deed of trust, assignment of leases, UCC-1, regulatory agreement, purchase agreement);
the inspection report; rent comps. Details and page references are in `references/deal-facts.md`.
