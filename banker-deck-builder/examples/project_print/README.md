# Project Print example (Xerox restructuring deck, Pari Passu case competition 2026)

Working builders and rendered results for every slide archetype. Use them to see exactly how a slide was
constructed; reuse the construction, not the numbers.

| Slide | Render | Builder | Archetype |
|---|---|---|---|
| 1 Cover | renders/s-01.png | (user-built) | cover |
| 2 Agenda | renders/s-02.png | (user-built) | agenda |
| 3 Company Snapshot | renders/s-03.png | (user-built benchmark) | snapshot, logo bar chart |
| 4 Situation Overview | renders/s-04.png | builders/build_situation.py | 2x2 modules, deal tags, braces, numbered issues |
| 5 Capital Structure | renders/s-05.png | builders/build_capstruct_v3.py | restyled native table, entity cards, arrows |
| 6 Business Plan | renders/s-06.png | builders/build_bizplan.py | bridge waterfall, case lines, track-record table |
| 7 Valuation Summary | renders/s-07.png | builders/build_valsum.py | football field vs. claim lines, equity bridge |
| 8 Public Comps | renders/s-08.png | builders/build_comps.py | logo rows (copied pictures), band, outline segment |
| 9 Precedents | renders/s-09.png | builders/build_precedents.py | eras, medians behind bars, truncation mark |
| 10 DCF | renders/s-10.png | builders/build_dcf.py | combo chart, value vs. claims, shape grid |
| 11 Liquidation | renders/s-11.png | builders/build_liquidation.py | book-to-net waterfall, percent-stacked funnel |
| 12 Recovery Waterfall | renders/s-12.png | builders/build_waterfall.py | claim outline vs. recovery fill, whiskers, diamonds |
| 13 Recommendation | renders/s-13.png | builders/build_rec2.py | before/after minis, timeline, terms table |
| 14-21 Appendix | renders/s-14..21.png | builders/build_appendix.py + render_xl.py | Excel outputs + commentary |
| (alt) Appendix recovery | - | builders/build_appx_rec.py | native-table version with toggle sensitivity |

Notes on the builders:
- They import `pp_helpers.py` (the earlier, hard-coded-palette version of `scripts/deck_kit.py`). New work should use
  deck_kit, which is the same API with theme-driven colors plus the helpers that were repeated across builders
  (callout, pill, closing_line, grid_table, new_content_slide, move_slide, copy_picture, Plot/chart).
- Each builder takes `<in.pptx> <out.pptx>`, appends its slide, moves it into position and saves a new version.
  The rebuild chain used after a model change: v9 -> valsum -> comps -> precedents -> fix_pronouns -> dcf ->
  liquidation -> waterfall -> rec2 -> appendix.
- Paths such as `model/v16_verify.xlsx` and `xl_img/` refer to that engagement's scratch folder.
- The renders use a substitute serif for Garamond, so spacing is slightly wider than in PowerPoint.

See plan_examples.md for two per-slide plans exactly as proposed to (and approved by) the user.
