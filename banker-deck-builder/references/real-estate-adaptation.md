# Adapting to real estate decks (IC / BOE pitches, acquisitions, development, LP positions)

The method of this skill carries over unchanged: storyline first, a per-slide plan the user approves, the claim tested
against the numbers before designing, native editable charts with exact geometry, sourced footnotes, the render and fix
loop, lint, appendix from model outputs. What changes is the storyline, the slide types, the vocabulary, and often the
house style.

## Contents
1. Which skill decides what (precedence)
2. What to skip from this skill
3. Real estate storyline
4. Real estate slide types (built with deck_kit)
5. Numbers and wording conventions
6. Questions an investment committee will ask

---

## 1. Which skill decides what

- If a dedicated real estate deck skill is installed (e.g. `real-estate-deck`, built on the SMU Cox fund's East Penn
  deck), **its house style, slide sequence, speaker-note voice, comp-map script and merge workflow win**. Read its
  SKILL.md and references first. Use this skill for what that one does not cover: the plan-before-build protocol,
  testing each claim against the numbers, the native chart construction (Plot / chart / bar_ser recipes in
  charts-and-geometry.md), the render/validate/lint loop and the appendix-from-model method.
- Where the two disagree on looks, the real estate skill wins. Example: it rules "nothing is tinted, pilled or tagged;
  rules and white space do the work". In that case do NOT use `callout`, `pill` or highlight rows; carry the
  emphasis with the accent color, bold, rules and placement instead. Its own colors (e.g. fund navy for "ours"),
  type scale and footer apply through theme.json.
- If no real estate skill is installed, use this skill's design system with the user's template, and this file's
  storyline and slide types.
- When editing a deck the team is also editing, prefer in-place edits of existing shapes over rebuilding the slide,
  so teammates' hand edits survive (the real estate skill's helpers do this).

## 2. What to skip from this skill

The restructuring material does not apply: rx-deck-playbook.md, and the snapshot / situation (balance-sheet
distress) / capital structure / football field / liquidation / recovery waterfall / prepack recommendation archetypes,
except as construction references (e.g. the football-field code is the template for any range chart; the bridge
waterfall code is the template for an IRR bridge).

## 3. Real estate storyline

Thesis in one sentence (e.g. "A well-located, entitled multifamily development at a basis below replacement cost, with
returns that clear the target even if lease-up runs slow"), then the beats:

| # | Slide | What it must prove |
|---|---|---|
| 1 | Executive summary and recommendation | The ask, the answer (advance / pass / invest at X), and the three reasons |
| 2 | Sponsorship | The sponsor can execute: track record on comparable deals, alignment (co-invest, fees) |
| 3 | Asset overview | What is being bought or built: units/SF, mix, amenities, status (say what is true, not what the OM says) |
| 4 | Location and submarket | Demand drivers with numbers: jobs, population, income, supply pipeline |
| 5 | Market and lease-up | Absorption and supply vs. the lease-up assumed |
| 6 | Rent comparables | Underwritten rents vs. comps, like-for-like by floor plan, trended vs. untrended named |
| 7 | Sales comparables | Basis and exit cap vs. recent trades ($/unit, $/SF, cap rate) |
| 8 | Project timeline / business plan | Entitlements, construction or renovation, lease-up, stabilization, exit dates |
| 9 | Sources and uses | Capital stack, where the position sits, debt terms |
| 10 | Key assumptions: sponsor vs. ours | Every material assumption side by side, ours in the deck's "ours" color |
| 11 | Returns and breakeven | IRR / multiple, the bridge from sponsor IRR to ours, breakevens and the rent-by-exit-cap grid |
| 12 | Key risks, diligence and recommendation | Risks with mitigants, open diligence items, the decision |

Page counts and order come from the user's brief or the real estate skill's slide specs; adapt the table to them.

## 4. Real estate slide types (construction with deck_kit)

- **Rent comps by floor plan**: grouped horizontal bars per floor plan (studio / 1BR / 2BR / 3BR), comps grey, subject
  in the accent, the underwritten rent as a dashed claim line per group; a $/SF column. Table of comps beside it.
  Like-for-like by floor plan, never one blended average that hides unit mix.
- **Comp map + table**: if the real estate skill's `comp_map.py` is available, use it (satellite basemap, numbered pins,
  photo callouts with leader lines). Otherwise ask the user for a map image and place numbered pins as shapes. The table
  or chart sits beside the map with matching numbers.
- **Sales comps**: $/unit (or $/SF) floating range or bars with the subject's basis as a claim line; cap rates as a
  second series or column; a band for the comp range.
- **Sources and uses**: a stacked column (senior debt, mezz/pref, LP equity, GP equity) beside a two-column table;
  highlight where the user's position sits in the stack and its attachment / detachment points.
- **Sponsor vs. ours table**: `grid_table` with assumption, sponsor, ours, delta, rationale; the "ours" column in the
  deck's ours color (outline_col or color); deltas that drive returns bolded.
- **IRR bridge**: the business-plan bridge recipe (invisible base + ups + downs + totals): sponsor IRR, "model mechanics"
  bar first (same inputs, two calculators), then one bar per assumption group, our IRR. Be ready to split any bar.
- **Breakevens**: horizontal bars or a small table: "rent decline to 1.0x", "exit cap to hit the 8% target", "months of
  lease-up delay to breakeven". Prefer breakevens over bear/bull cases when the real estate skill or user says so.
- **Rent x exit-cap sensitivity grid**: the shape-grid recipe (DCF slide) with the base case dark and cells below the
  target in the accent; steps as the user specifies (e.g. 2.5% rent, 25 bps cap).
- **Annual cash flow and returns**: columns of levered cash flow by year with the exit year distinguished; IRR and
  equity multiple labeled; optional cumulative line.
- **Timeline**: the recommendation timeline recipe (dots, dates above, steps below), with the critical-path item in
  the accent.
- **Supply and absorption**: combo (deliveries as columns, occupancy or absorption as a line) with the subject's lease-up
  window shaded as a band.

## 5. Numbers and wording conventions

- Cap rates and yields in % to two decimals ("5.25%"); changes in bps; yield on cost vs. exit cap spread in bps.
- $/unit, $/SF, $/key written out ("$285K/unit", "$312/SF"); rents monthly ("$1,850/mo") unless the market quotes
  annual $/SF.
- Name every rent as trended or untrended; every return as levered or unlevered, gross or net of fees/promote.
- IRR to one decimal, equity multiple to two ("14.2% IRR, 1.85x").
- Say what is true: "entitled, land closed" rather than marketing language from the OM, unless verified.
- All other rules of writing-and-numbers.md apply (no pronouns on slides, no dashes as punctuation, sourced and dated
  data, every figure's period named).

## 6. Questions an investment committee will ask

- What is our basis vs. replacement cost and vs. recent trades?
- What rent growth and exit cap are assumed, and what do comps and the forward curve support?
- How long is lease-up and what happens to returns if it runs 6 months longer?
- What does it take to lose money (the 1.0x breakeven)? To miss the target return?
- Where does the position sit in the capital stack, and what are the debt terms (rate, hedge, maturity, extension
  tests)?
- How is the sponsor paid (fees, promote) and how much of their own capital is in?
- For a small LP position: is the position offered as is (advance / pass) rather than negotiable?
Put the answer on the relevant slide or in the appendix.
