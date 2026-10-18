# Workflow: from brief to a finished deck, one slide at a time

This is the collaboration protocol that produced the Project Print deck. The user's satisfaction came from the
process as much as the pixels: plan first, get a yes, build, render, fix, explain. Follow it.

## Contents
1. Intake
2. Storyline and page plan
3. The per-slide plan (what to propose before building)
4. Build
5. Render, inspect, fix (the QA loop)
6. Deliver and report
7. Revisions, comparisons with other versions, and pushback
8. Working rules the user set (keep honoring them)

---

## 1. Intake

Collect, and read before proposing anything:
- **The brief**: who the audience is (judges, IC, client), the role being played (e.g. debtor advisor), deadline,
  page limits (main pages vs. appendix), required topics and their order, deliverable format (PDF + video?).
- **The template / master**: the .pptx the user is working in. Run `scripts/inspect_template.py deck.pptx` to find
  the content layout, placeholder idx values, footer, logos already placed, theme fonts and the palette in use.
  Write `theme.json` (start from `assets/theme.example.json`). If theme fonts are Office defaults (Aptos/Calibri),
  tell the user: inserted charts inherit them unless the builder sets the font explicitly (deck_kit does).
- **New palette or brand on an existing master**: change it once in the slide master and theme (title tab color,
  logo, theme colors and fonts), not by hiding and redrawing elements on every slide. If the user has not done it,
  say so and offer to edit the master; only as a stopgap redraw on the slide, and tell the user.
- **Reference decks**: study them for what works and what does not (see design-system.md "Reference deck lessons").
  Summarize the common strengths and failure modes back to the user once, briefly.
- **The data source**: usually the user's model. Read values from a recalculated copy (formulas written by
  openpyxl have no cached values until recalculated). Never invent numbers; flag estimates as estimates.
- **The user's existing slides**: the first content slides the user approves become the style benchmark for every
  later slide ("stick to the clean formatting of the first 2 content slides"). Re-read them before each build.

Do not build anything during intake unless asked. "Keep holding" means read and wait.

## 2. Storyline and page plan

Before any slide, write the deck thesis in one sentence and the arc in 4 to 6 numbered beats (see
rx-deck-playbook.md for the restructuring arc). Then map beats to pages:
- Count pages against the limit. If required topics + agenda exactly fill the main section, no topic gets a second
  page; overflow goes to the appendix. Say this explicitly when it constrains a request (e.g. "no industry slide").
- Each slide must advance the argument and hand off to the next. Write each slide's takeaway (the claim) and its
  closing line (the hand-off) before designing visuals.
- Note which slides need model outputs and which do not, so non-dependent slides can be built while the model is
  finished.

## 3. The per-slide plan (propose before building)

When the user asks "what's next" or "lay it out", answer with this structure, in plain prose and short tables,
then stop and wait for a go-ahead:

1. **The story job**: one or two sentences on what this slide proves and how it connects the previous slide to the
   next one. ("The DCF answers the objection 'what if the plan works?'... hands off to Liquidation.")
2. **Takeaway**: the exact one-line text (a claim with numbers, not a topic label).
3. **What goes where**, panel by panel:
   - Left panel (main exhibit, ~60% width): chart type, series, colors (accent only for the subject/distress),
     annotations (claim lines, brackets, bands, pills), and the point the reader should see.
   - Right panel(s): the supporting chart/table and its point.
   - Callout(s): the exact sentence.
   - Closing line: the exact sentence.
4. **Test the claim against the numbers first.** Before designing anything, do the arithmetic that proves or
   breaks the requested message (e.g. break-even multiple = net debt / EBITDA vs. the peer median). If the numbers do
   not support the user's intended claim, say so up front, propose the claim they do support, and build that one
   (or ask, if the user is available). Never build a slide whose takeaway the data contradicts.
5. **Numbers**: list the figures that will appear, with their source (model tab / filing). Recompute anything
   derived and show the arithmetic when it matters.
6. **Flags before building**: inconsistencies between slides or with the model, data the user must supply
   (logos, quotes, sources), judgment calls the user may want to reverse, and the questions a judge will ask
   with the answer the slide gives.
7. **What moves out**: anything dropped from the main page and where it goes (appendix, video, footnote).

Keep the plan concrete enough that the build is mechanical. If a decision is genuinely the user's (e.g. reveal
the recommendation early or not), offer the alternative in one line rather than asking an open question.

## 4. Build

- One builder script per slide: `build_<slide>.py <in.pptx> <out.pptx>`. It loads the theme, adds the slide with
  `new_content_slide`, draws every element, moves the slide to its position and saves a NEW versioned file
  (`Deck_v13.pptx` from `Deck_v12.pptx`). Never overwrite the user's upload; never edit in place.
- Keep builders deterministic and rerunnable: when numbers change upstream (model update), rerun the chain of
  builders in order on the last user-approved base deck. Keep the chain documented (base -> builder A -> builder B).
- Hard-code slide numbers in the builder from the model *with a comment naming the source cell/tab*, or read them
  from the recalculated model directly. Either way, every number traces to a source.
- Name every shape meaningfully (`"A football field chart"`, `"C callout"`). It makes later edits by the user and
  by you far easier.
- When the user has edited the deck since your last version, build on THEIR latest file, not yours.

## 5. Render, inspect, fix (never skip)

After every build:
1. `python3 scripts/render_slides.py deck.pptx out/rNNa --pages K K` (new folder each round: r07a, r07b, ...).
2. Look at the PNG. Check, in order: text collisions (labels on rules, labels on markers, labels over lines),
   wrapped one-liners (takeaway, closing line, callouts leaving an orphan word), overflowing boxes, labels cut off
   at chart edges, values that do not match the plan, unbalanced whitespace, and anything that looks generic.
3. Fix and re-render into a new folder. Typical fixes: move a chart down to clear a header rule, put a label
   inside a bar instead of beside it, shorten a phrase to kill a wrap, white-fill a label that sits on a dashed
   line, widen a column, drop a redundant reference line.
4. Repeat until clean. Two or three rounds is normal.
5. Validate the file (`validate.py deck.pptx --original <user's base deck>` from the pptx skill) and run
   `scripts/lint_deck.py deck.pptx`. Both must pass (lint warnings on the user's own slides can be reported, not fixed).

Remember the render uses a substitute font when the deck font is not installed: if text only just fits or only
just wraps, say so; if it clearly collides, it is real.

## 6. Deliver and report

Send the new deck file (and a PNG of the new slide when useful). Then report briefly:
- What the slide shows, panel by panel, and **how it advances the story** (the user explicitly values this).
- Fixes made during QA only if they matter to the user.
- Numbers that changed versus earlier slides, and any knock-on edits made to other slides for consistency.
- Open items the user must handle (logos, sources, quotes), in a short list.
- The next slide and a one-line offer to lay it out.
No pronoun-free requirement applies to chat, but keep chat plain, direct, and free of AI tells (no em dashes).

## 7. Revisions, comparisons, pushback

- **User edits**: when the user edits a slide and asks "did I make it better?", give an honest split: what improved
  (and keep it everywhere), what to reverse or adjust (with the precise reason, e.g. a figure that is now wrong).
- **Another model's version**: compare item by item: what theirs does better (adopt it), what yours does better
  (keep it), what in theirs is wrong or unsourced (do not carry it over). Then propose the merged build.
- **Pushback**: when guidance will not get to the needed story, say so directly with the reason and a better
  option. The user asked for this. Example: rejecting an out-of-court exchange because holdouts and legacy notes
  stay ahead of new equity; fixing the WACC because cost of equity was below pre-tax cost of debt.
- **Upstream fixes first**: if a slide exposes a model problem, propose the fix, and when approved fix the model
  first, re-derive the numbers, update every slide that quotes them, then build the new slide.
- **Keep alternatives instead of deleting**: when the user wants to choose ("don't delete the old one"), build both
  versions side by side and say which you recommend and why.

## 8. Working rules the user set (persist across the whole engagement)

- Do not build or edit the deck or model unless explicitly asked. "What do you think?" means opinion only.
- Push back when guidance will not reach the story. Be direct.
- No em dashes or other AI-writing tells anywhere (chat, slides, files).
- No personal pronouns on slides (we / our / us / you / I). Slides speak in the third person.
- Real, sourced numbers only. Flag estimates. Never invent figures, market shares, quotes or logos.
- Native, editable charts. Accent color reserved for the subject company or distress.
- No footer (logo, rule, page number) on title or cover slides.
- Slides must not look AI-generated (see design-system.md "Anti-patterns").
