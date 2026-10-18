---
name: banker-deck-builder
description: Plan and build investment-banking-quality PowerPoint decks slide by slide (restructuring / RX pitch decks, case-competition decks, M&A and valuation pitch books, IC decks) on the user's own template, the way the Project Print deck was built - storyline first, a per-slide layout plan the user approves, then native editable charts with claim lines, bands, brackets, callouts and footnotes, a render-and-fix QA loop, and an appendix of real Excel model outputs. Use this whenever the user wants to plan, lay out, build, rebuild, critique or polish banker-style slides or a pitch deck in .pptx, asks "what's our next slide", "lay out the X slide", "build the valuation summary / comps / precedents / DCF / liquidation / waterfall / recommendation slide", wants slides that "tell a story" and do not look AI-generated, or wants appendix pages from their model - even if they do not say "skill" or name the template.
---

# Banker Deck Builder

Build decks the way a strong analyst and a picky VP would together: a clear argument, one claim per slide, charts
that prove it, nothing decorative, every number sourced, and every slide rendered and checked before it is shown.

## Generalization first: the specifics are evidence, not constants

This skill was distilled from one engagement (Project Print: a Xerox restructuring deck for the Pari Passu case
competition, Garamond on a near-black/red master). Colors, font, slide master, placeholder indices, page grid,
content and even the storyline will differ next time. Keep the PROCESS, the RULES and the CONSTRUCTION TECHNIQUES;
take every concrete value (hex codes, inches, idx numbers, section list) from the new user's template, brief and
reference decks. The Xerox numbers in the examples are illustrations of how to build, never data to reuse.

## The workflow (read references/workflow.md for the full protocol)

1. **Intake.** Read the brief (audience, role, page limits, required sections, deadline), the user's template and
   their reference decks. Run `scripts/inspect_template.py <deck.pptx>` and write `theme.json` from
   `assets/theme.example.json` (palette roles, font, template slide, placeholder idx, footer, grid). If the user says
   "keep holding", only read.
2. **Storyline.** One-sentence thesis, 4 to 6 beats, mapped to pages within the limit (references/rx-deck-playbook.md
   for the RX arc). Each slide gets a claim (takeaway) and a hand-off (closing line) before any visual is designed.
3. **Per-slide plan, then wait.** For each slide, propose: the story job; the exact takeaway; what goes where (left
   panel, right panels, callouts, closing line); the numbers with sources; flags (inconsistencies, missing inputs,
   judge questions); what moves out. Build only after the user says go. If the user asks "what do you think", give
   an opinion only.
4. **Build** with one script per slide using `scripts/deck_kit.py` on the user's latest deck; save a new version.
5. **Render, inspect, fix** until clean: `scripts/render_slides.py` (new folder each round), look at the PNG,
   fix collisions/wraps/overflow/balance, re-render. Then the pptx skill's `validate.py --original <base>` and
   `scripts/lint_deck.py`. Never show an unrendered slide.
6. **Deliver and report**: send the file; explain what the slide shows and how it advances the story; list changes to
   other slides and open items; offer the next slide's plan.

## Rules the user set (honor them for the whole engagement)

- Do not build or edit the deck or model unless explicitly asked; "what do you think" means opinion only.
- Push back, directly, when guidance will not get to the story; propose the better option.
- No em dashes or other AI-writing tells anywhere. No personal pronouns on slides (we / our / us / you / I).
- Real, sourced numbers only; flag estimates; never invent figures, quotes, market shares or logos.
- Native, editable charts. The accent color is reserved for the subject company and distress.
- No footer elements on title/cover slides.
- Slides must not look AI-generated (references/design-system.md, Anti-patterns).
- When a slide exposes an upstream (model) problem: propose the fix, fix the model first when approved, update every
  slide that quotes the changed numbers, then build.
- When the user wants to choose between versions, build both and keep the old one.

## What a finished main slide looks like (references/design-system.md)

- Title (master placeholder) + one-line italic takeaway band that states the claim with numbers.
- Body on a grid: main exhibit left (~60%), supporting exhibit/table/callout right (~35%); or 2x2 modules.
  Two to four exhibits, each under a bold panel header with a full-width rule, each making one point.
- Charts native and annotated: claim lines (dashed, accent for the critical claim), bands for ranges, brackets tying
  a range to its label, pills flagging a point, highlight row for the subject, logos for peers, labels written out
  with units, placed where nothing else is.
- One grey-filled, accent-dotted commentary callout with the "so what"; white pills for chart flags.
- Closing line: full-width dashed box, one sentence, conclusion plus hand-off to the next slide.
- Footer: numbered Notes with superscripts on the exact numbers, and Sources.
- Accent used only on the subject/distress. One font. Nothing below 9 pt that matters.

Slide archetypes with layouts, construction and QA lessons: references/slide-archetypes.md (snapshot, situation
2x2, capital structure, business plan, valuation summary/football field, comps, precedents, DCF, liquidation,
recovery waterfall, recommendation, appendix page, agenda, cover).

## Building with deck_kit (scripts/deck_kit.py)

```python
import sys; sys.path.insert(0, "<skill>/scripts")
from deck_kit import *
load_theme("theme.json")
prs = Presentation("out/Deck_v12.pptx")              # the user's latest file
g = THEME["grid"]
s = new_content_slide(prs, "Valuation Summary", "<claim with numbers>", notes="Notes: (1) ...", sources="Sources: ...")
section_header(s, g["L0"], g["ROW1"], g["LP1"] - g["L0"], "Enterprise Value Range vs. Claims ($M)", "A header")
P = Plot(2.55, 2.17, 5.55, 3.5, n=len(cats), vmin=1500, vmax=4800, horizontal=True)
chart(s, P, cats, [("Low", lo), ("Range", span)],
      [nofill_ser(0, "Low", cats, lo), bar_ser(1, "Range", cats, span, "ink", dpt={5: "grey_lt"})],
      "A football field", grouping="stacked", gap=55)
line(s, P.v(4355), P.y - 0.3, P.v(4355), P.y + P.h, "accent", 1.25, dash=MSO_LINE_DASH_STYLE.DASH)
grid_table(s, g["RP0"], 1.72, [1.84, .87, .87, .88], ["($M)", "Low", "Mid", "High"], rows, "B table", outline_col=2)
callout(s, g["RP0"], 5.0, g["R1"] - g["RP0"], 0.62, [("Market cap is ", {}), ("$324M", {"b": True})], "C callout")
closing_line(s, [("Value covers secured debt ", {}), ("~1.5x", {"b": True}), ("...", {})])
move_slide(prs, s, 6); prs.save("out/Deck_v13.pptx")
```
Key pieces: `new_content_slide`, `move_slide`, `drop_slide`, `section_header`, `caption`, `textbox` (runs as
`(text, {"b","i","c","sup","size"})`, colors by role name), `box`, `line`, `arrow`, `callout`, `pill`, `closing_line`,
`hbracket`, `brace`, `numbered_items`, `card`, `grid_table`, `style_cell` (native tables), `copy_picture` (reuse
the user's logos), `add_image`, `Plot` + `chart` / `add_chart` + `bar_ser` / `nofill_ser` / `line_ser` /
`bar_chart_xml` for native charts with exact geometry. Chart technique and recipes: references/charts-and-geometry.md.

Builder hygiene: one script per slide, rerunnable, reading from the user's latest deck and writing a new version;
numbers from the recalculated model with a comment naming the source; meaningful shape names; keep the builder chain
documented so an upstream change can be replayed across slides.

## QA before showing anything (references/qa-and-delivery.md)

Render the slide, look at it, and fix: collisions (labels on rules/lines/markers), wraps (takeaway, closing line,
orphan words), overflow, wrong or stale values, empty or cramped areas, accent misuse, inconsistency with approved
slides. Re-render into a new folder each round (never rm render folders). Validate and lint. Renders may use a
substitute font; marginal wraps may not occur in PowerPoint, clear collisions are real.

## Writing (references/writing-and-numbers.md)

Takeaway = one-line claim with numbers. Closing line = finding + implication/hand-off. Callout = one sentence. No
pronouns, no dashes as punctuation, no AI phrasing. "$1,739M", "($990)", "5.9x", "52c", "4.9x to 6.4x". Every
figure has its period. Numbers tie across slides or a footnote explains the different basis.

## Appendix (references/appendix-model-outputs.md)

Appendix pages are real model outputs rendered with `scripts/render_xl.py` (the model's own formatting), with a short
numbered commentary column and one key-number callout. Choose pages that back the main slides and pre-answer judge
questions; never paste a range that contradicts a main slide.

## Reference map

| Need | Read |
|---|---|
| Collaboration protocol, plan format, revisions, pushback | references/workflow.md |
| Grid, type, color roles, components, anti-patterns, reference-deck lessons | references/design-system.md |
| How each slide type was laid out and built | references/slide-archetypes.md |
| Native chart construction and label placement | references/charts-and-geometry.md |
| Takeaways, closing lines, number formats, footnotes, sourcing | references/writing-and-numbers.md |
| Render/validate/lint checklist and delivery | references/qa-and-delivery.md |
| Appendix from model outputs | references/appendix-model-outputs.md |
| RX storyline, judge questions, adapting to other deck types | references/rx-deck-playbook.md |
| What this user praised and corrected | references/feedback-log.md |
| Working builders + rendered results for every archetype | examples/project_print/ (README.md) |
