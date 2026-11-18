# Taylor Commons deck: instructions for Claude in PowerPoint

These prompts apply the Nov-18 changes to the live team deck (`v1_Taylor_Commons_AM_pitch_11.11.2026`). Others are
editing the same file, so paste **one batch at a time** and check each one before sending the next.
Always refer to slides **by title, not number**. The deck has hidden slides, so the numbers PowerPoint shows don't
match the PDF.

Three changes need a person, not Claude in PowerPoint:
1. **Paste in the rebuilt slides.** The Market, Submarket and Valuation slides in the live deck are older
   versions. Their newer versions have native charts and are in `TaylorCommons/v2_Taylor_Commons_AM_pitch_11.16.2026.pptx`.
   1. In PowerPoint, go to Home > New Slide > Reuse Slides and pick that file.
   2. Check **Keep source formatting**.
   3. Insert "Market Overview: Dallas-Fort Worth", "Submarket Overview: East Fort Worth" and "Valuation and
      Recovery".
   4. Delete the old copies.
2. **Fill the bracketed items.** Anything in [square brackets] below is something only the person who did the
   Nov-26 secret shop knows. Fill it in or delete it. Never let Claude invent it.
3. **Sales comps.** These need real trades from CoStar or broker calls (see Batch 6).

---

## Batch 0: house rules (paste first, once)

```
You are editing the SMU Cox Real Estate Impact Fund's Taylor Commons asset management deck. Other people are editing
this file right now: change only the slides I name, find them by title, and do not move, restyle or delete anything
else unless I ask.

House style for every change:
- Font Garamond. Body text 12pt (never under 10pt in the body). Footer notes and sources 9pt.
- Every content slide uses the "CONTENT" layout: title in the title placeholder; one italic takeaway sentence (with
  numbers) in the subtitle placeholder under it; grey rule; body; a closing line in a dashed rounded box near the
  bottom (12pt, centered, key phrase bold); notes and sources in the footer band next to the SMU logo (9pt, short).
- Colors: charcoal ink 22262B, slate 4A5560, greys 5F6670 / 8C939A / C9CDD1, light fill EEEDEA. Brick red B4442C
  only for Taylor Commons itself and for bad news. No SMU navy, no red corner swoosh, no bright blue.
- Panel headers: 15pt bold with a thin charcoal rule under them.
- Third person only on slides (no we, our, us, you, I). No em dashes. Numbers like $1.43M, ($612K), 0.93x, 74.9%,
  $13.4M to $15.6M. Months like Nov-26.
- Do not invent numbers. Text in [square brackets] is a placeholder for a teammate: leave it in brackets.
```

---

## Batch 1: reorder the deck

```
Reorder the slides to match this list exactly (match by title; titles in quotes are the current or new titles).
Before moving anything, list every slide including hidden ones with its title and whether it is hidden, and tell
me about any slide that is not on this list; do not delete those yet.

1  Cover
2  Agenda
3  Asset Overview
4  Site Visit: Location, Curb Appeal and Units
5  Site Visit: Amenities, Leasing and Marketing
6  Secret Shopping - Update            (rename to "Secret Shopping Update (Nov-26)" in Batch 2)
7  Location Overview
8  Migration Trends
9  Market Overview: Dallas-Fort Worth
10 Submarket Overview: East Fort Worth
11 Then vs. Now: IC Underwriting vs. Today
12 Operations: Where Leasing Broke Down
13 Historical Financials vs. Underwriting   (new or empty; see Batch 5)
14 Rent Comparables
15 Rent Comparables Map
16 Sales Comparables
17 The March 2025 Refinancing Trade
18 HB 21: What Ends and What Stays
19 The 2027 Cliff and Refinancing Gap
20 Valuation and Recovery
21 Renovation Plan
22 What Went Wrong
23 Key Risks and Mitigants                  (new; see Batch 5)
24 How Can the Fund Avoid the Same Mistakes
25 What the Fund Does Now                   (new; see Batch 5; replaces any empty "Final Recommendation")
26 Appendix
27-33 the seven "Secret Shopping - Asset Photos" slides, in this order: Community Signage & Visibility (3 photos),
   Community Signage & Visibility (2 photos), Leasing Office, Model Unit Interior, Deferred Maintenance - Foundation
   Issues, Deferred Maintenance - HVAC Units, Deferred Maintenance - Misc.

If a slide in the list does not exist yet, add a blank slide on the CONTENT layout with that title in its place.
Then show me the final order.
```

Next, update the agenda so its page numbers match:

```
On the "Agenda" slide, update the ten rows (keep the layout) to:
1  Asset Overview | 224 units, built 1979, bought in Jun-23 for $25.9M | 3
2  Site Visit Findings | Hard to find and dated; a model unit now, but deferred maintenance shows | 4 to 6
3  Location & Market | DFW is recovering; east Fort Worth Class C is not | 7 to 10
4  Performance vs. Underwriting | 75% occupied; T12 NOI of $1.43M vs. $2.21M planned | 11 to 13
5  Comparables | Comp set 72.8% occupied; TC asks $970, $56 below the comp average | 14 to 16
6  Capital Structure & 2027 Maturity | 9% bridge, HB 21 tax cliff and a $7.1M refinancing gap | 17 to 20
7  Renovation Plan | A $1.8M upgrade budget vs. 12 renovated units | 21
8  What Went Wrong & Key Risks | Operations missed first; the HFC refinancing made it worse | 22 to 23
9  Lessons & Recommendation | Screens to catch this earlier, and what the fund does now | 24 to 25
10 Appendix | Secret shop photos and supporting detail | 26 to 33
Keep the takeaway as is. Each preview must stay on one line at 12pt; shorten it if it wraps.
```

---

## Batch 2: secret shopping body slide

Context: the inspection report dated Nov-11-25 is the **first** visit. My two Site Visit slides summarize it.
The photo slides come from the **Nov-26** secret shop. The biggest change since then is a staged 1BR model unit,
which didn't exist in Nov-25. Signage, hours, fees and condition haven't improved, and some have gotten worse.

```
Rework the slide titled "Secret Shopping - Update". Keep its structure (two panel headers "Factors of Low Occupancy"
on the left about 60% width, "Current Listing Offerings" on the right, dashed closing box at the bottom), but:

Title: Secret Shopping Update (Nov-26)
Takeaway (italic subtitle placeholder, replacing "Overall Experience"):
A staged model unit is the one fix since the Nov-25 visit; signage, hours, fees and upkeep still cost Taylor Commons tours

Left panel "Factors of Low Occupancy": a numbered list, 12pt, bold lead-in then plain text, each tagged in grey
italics with its status vs. the Nov-25 visit:
1. Hard to find: one small, unbranded "Now Leasing" sign on the I-30 frontage road; the gate sign reads only up
   close (unchanged)
2. Closed when renters shop: Mon to Fri 9 to 6, Sat 10 to 5:30, Sun closed; the gym keeps the same hours
   (unchanged) [add: office was closed at the time of the visit, if true]
3. Higher all-in cost: a $175/mo fee package vs. $135 at The Brittany and $99 at Tides at 5400 erases most of the
   $56 rent discount to comps (unchanged) [confirm fees quoted on Nov-26 visit]
4. Units not ready: 49 of 56 vacant units not ready to lease (box score 11/15/26)
5. Visible deferred maintenance: stair landings on cinder blocks, rusted HVAC condensers, rotted siding (new
   since Nov-25; photos in appendix)
Below the list, one line in charcoal bold: "Fixed since Nov-25: staged 1BR (A2) model unit by the office"

Right panel "Current Listing Offerings": a small table, 12pt, hairline rules:
  Floor plan | Asking rent
  1 bed (avg) | $894
  2 bed / 1 bath | $1,091
  2 bed / 2 bath | $1,178
  Renovated (avg) | $1,202
  Property average | $970
then three short lines under it:
  Concession: 6 weeks free
  Referral bonus: $500 / $750 / $1,000 [confirm still offered Nov-26]
  Online: website and listings current; [number] reviews [confirm]
[Add anything the leasing agent quoted on Nov-26: specials, move-in costs, availability, lead time to a ready unit]

Closing line in the dashed box (replace "Desc"):
**The tour is no longer the problem; getting prospects to the gate and into a ready unit is**

Footer, 9pt:
Notes: Asking rents from the Rent Comparables slide (box score 11/15/26); fee packages from the Nov-25 visit.
Sources: Team secret shop (Nov-26); Taylor Commons inspection report (Nov-25); TexSun box score.
```

Then flag the earlier visit on the two site visit slides:

```
On the slides titled "Site Visit: Location, Curb Appeal and Units" and "Site Visit: Amenities, Leasing and
Marketing", add " (Nov-25)" to the end of each title. On the second one, wherever it says there is no model unit,
add in grey italics: "(model unit added by Nov-26)". Change nothing else.
```

---

## Batch 3: the seven secret shop photo slides (appendix)

```
For each of the seven slides titled "Secret Shopping - Asset Photos":
1. Apply the CONTENT layout from this deck's master (the same layout as "Asset Overview"), so the navy/red bar and
   swoosh disappear. Keep the photos where they are and keep their size.
2. Delete the bright blue "Incomplete - add Desc." box.
3. Put the new title in the title placeholder and the takeaway in the italic subtitle placeholder.
4. Under each photo add an 11pt grey (5F6670) caption, left aligned to the photo's left edge, at most two lines.
5. Put the closing line in the existing dashed box (12pt, centered, key phrase bold) and restyle that box to a thin
   grey dashed rounded rectangle.
6. Add the footer line, 9pt: "Source: Team secret shop (Nov-26)."
Do them one slide at a time and show me each before moving on.
```

The text for each slide, in deck order:

**1. Community Signage & Visibility (3 photos)**
- Title: `Secret Shop: Signage and Visibility`
- Takeaway: `A year after the first visit, the only cue from I-30 is still a small, unbranded "Now Leasing" sign`
- Captions, left to right:
  1. The frontage-road sign: no property name, no logo
  2. The same corner from Bridge St.: the office towers dominate the view
  3. Overlook Terrace curves behind the offices; the entrance can't be seen from the road
- Closing: `Drive-by leasing traffic is close to zero until a **branded monument sign goes on the I-30 frontage road**`

**2. Community Signage & Visibility (2 photos)**
- Title: `Secret Shop: Arrival`
- Takeaway: `The only way in is a cracked side street, and the gate sign reads only once a prospect is already there`
- Captions:
  1. Overlook Terrace at Bridge St.: patched, cracked paving on the single entry route
  2. Gate and monument sign (boxed): legible only from the entry drive; the flags help, but only up close
- Closing: `First impressions start at the gate: **resurfacing the entry and adding a lit sign are low-cost fixes**`

**3. Leasing Office**
- Title: `Secret Shop: Leasing Office`
- Takeaway: `The office is the best-kept building on site, but it closes at 6 on weekdays and all day Sunday`
- Captions:
  1. Posted hours: Mon to Fri 9 to 6, Sat 10 to 5:30, Sun closed [add: "Sorry, we're closed" sign was up at
     ___ on ___]
  2. Office exterior: clean, landscaped and freshly painted
  3. Entry planters and railing: well kept
- Closing: `Most target renters work days: **extended weekday and Sunday hours cost little and add tours**`

**4. Model Unit Interior (1 Bedroom, Floorplan A2)**
- Title: `Secret Shop: Model Unit (1BR, A2)`
- Takeaway: `TexSun added a staged model since the Nov-25 visit, closing the biggest gap in the tour`
- Captions:
  1. Kitchen: new counters and backsplash; original cabinets, white appliances and tile
  2. Living area staged with decor and a printed flyer; lamps make up for weak overhead light
  3. Bedroom staged and lamp-lit; tile floors, blinds
- Closing: `The model shows well, but **it is a classic unit; the renovated finishes the plan promised are not on
  display**`
- [Confirm the model is a classic A2 and not an A2R. If it is renovated, change the closing to "it shows the
  renovated finish, but only 12 units have it".]

**5. Deferred Maintenance: Foundation Issues**
- Title: `Secret Shop: Foundation and Stair Landings`
- Takeaway: `Concrete stair landings have settled and cracked and sit on loose cinder blocks, a resident safety
  issue`
- Captions:
  1. A landing pulling away from the walk at [building #]
  2. Cracked slab edge propped on cinder blocks
  3. The same landing: blocks used as shims under a resident stair
- Closing: `This is a liability, not a cosmetic issue: **ask TexSun for a structural review and a repair cost
  before more units turn**`

**6. Deferred Maintenance: HVAC Units**
- Title: `Secret Shop: HVAC Condensers`
- Takeaway: `Many condensers look original-era: rusted, faded and packed with leaves, and [#] were not running`
- Captions:
  1. Two aging Trane condensers; coil tops full of debris
  2. A newer replacement next to an older unit with a damaged, debris-filled top
  3. A pair of faded units with leaf-clogged fan grilles
- Closing: `Expect more summer failures and **replacement capex the ~$340K left in escrow will not cover**`
- [Ask TexSun for an HVAC age survey. Don't put a cost on the slide without one.]

**7. Deferred Maintenance: Misc.**
- Title: `Secret Shop: Exterior Upkeep`
- Takeaway: `Small repairs are left open: rotted siding, warped closet doors, torn screens and loose utility boxes`
- Captions:
  1. Siding rotted through behind a downspout: water damage at the wall base
  2. Exterior closet door weathered and stained at the base
  3. Torn window screen; loose wiring and a utility box at grade
- Closing: `Each item is cheap; together they **tell prospects the property is not being looked after**`

---

## Batch 4: fixes to teammates' slides (numbers must match the rest of the deck)

```
Make only these edits:

"Rent Comparables Map": replace the unfinished takeaway "All comps are within a" with:
Seventeen 1977 to 1995 garden comps surround Taylor Commons north of I-30, and the set is only 72.8% occupied
Add the slide number and a 9pt footer: "Source: Team rent survey (Nov-26); Google Earth."

"Rent Comparables": change "6-weeks of concessions" to "six weeks of concessions". Add to the footer notes, 9pt:
"Validated by calls to [#] properties and [broker name], [dates]." (the assignment requires direct outreach)

"What Went Wrong":
- Market and Leasing, first item: the 19.8% submarket vacancy and 398 new units conflict with the Submarket
  Overview slide (91.3% occupied, 670 units delivered in 2024, 200 under construction). Replace the item with:
  heading "The comp set is only 72.8% occupied"; text "The 17 nearest comps average 72.8% occupancy and six weeks of
  concessions; east Fort Worth rents fell 4.6% in a year". [If the 19.8% figure has a source, tell me and the
  Submarket slide will be updated instead.]
- Market and Leasing, second item: change "From an 88.8% peak in summer 2025 to 74.8% at Sep-26" to
  "From a 90.2% peak in Jul-25 to 74.9% in Nov-26" so it matches the Operations slide; keep the retention text.
- Market and Leasing, third item: change the text to "The fund's ~$19.2M mark assumes the tax break survives;
  without it value is $13.4M to $15.6M vs. ~$21.5M of debt and pref".
- Footer note (1): add the period so it reads "T12 Oct-25 to Sep-26 (taxes accrued)". The Refinancing slide shows
  0.93x on the T12 to Feb-26, so both slides must state their period.

"How Can the Fund Avoid the Same Mistakes":
- Change "had not shown a completed investment track record from cradle to death on more complex value-add business
  plans" to "had not completed a full value-add cycle, from acquisition to exit, on a comparable asset".
- Delete the footer note "(1) Added interest less property-tax savings and HFC fees..." (there is no (1) on this
  slide); keep the sources line.
- Left-align the text in the middle and right columns and set it to 12pt; keep the table structure.
```

---

## Batch 5: the three missing slides (build each on the CONTENT layout)

### Key Risks and Mitigants (required by the assignment)

```
Build the slide titled "Key Risks and Mitigants".
Takeaway: Five risks decide what is left of Taylor Commons; the fund controls none of them directly but can press TexSun on each
Body: a table across the full width, 12pt, hairline rules, columns: Risk | Exposure | Mitigant | Who acts.
Color only the Exposure figures brick red.
Rows:
1. Loan maturity | $16.5M due Mar-27; new loan sizes to ~$9.4M, a $7.1M gap | Negotiate a further extension or
   forbearance now, not in Feb-27; fund commits no new capital | TexSun, W Financial
2. HB 21 tax cliff | ~$520K/yr of taxes from 1/1/27 unless Fort Worth approves the HFC | Protest the 2027 Tarrant
   value; track the HB 21 lawsuit; price the cost to unwind the HFC | TexSun
3. Leasing | 74.9% occupied; 49 of 56 vacant units not ready | Fund make-ready first; extend office hours; cut
   the $175 fee package; branded I-30 sign | TexSun, property manager
4. Physical condition | Settled stair landings, aging HVAC, open exterior repairs; ~$340K left in capex escrow |
   Structural review and HVAC survey; spend escrow on safety items first | TexSun
5. Sponsor reporting | $2.8M of extra bridge proceeds, $803K of partner loan interest and 2026 tax accruals not
   explained | Written request for monthly reporting and answers before the Mar-27 decision | SMU fund (LP)
Closing line: The fund's lever is information and consent, not capital: **ask for answers before March 2027, not after**
Footer 9pt: "Sources: TexSun T12 and box score; AM model v1_2; Deed of Trust (3/24/25); team secret shop (Nov-26); team analysis."
```

### What the Fund Does Now (closing recommendation)

```
Build the slide titled "What the Fund Does Now".
Takeaway: Every path at the March 2027 maturity leaves the common at $0, so the fund should hold, fund nothing and press for answers

Left panel (60%), header "Three Paths at the Mar-27 Maturity": three stacked cards (light fill EEEDEA, 12pt),
each with a bold path name, two lines of text and a right-aligned outcome for the common in brick red:
- Refinance: a new loan sizes to ~$9.4M against $16.5M due; ~$7.1M of new equity would sit behind the $5.0M pref | Common: $0 unless new money is added
- Extend with W Financial: buys time, but 9% on $16.5M is ~$1.49M a year against ~$0.94M of NOI after taxes | Common: $0, shortfall grows
- Sell or hand back the keys: value of $13.4M to $15.6M is below the $16.5M loan | Common: $0

Right panel (35%), header "Recommendation": numbered list, 12pt:
1. Mark to the base case: carry the property at $13M to $16M, not $19.2M; SMU's $350K stays at $0
2. Fund no rescue capital: decline any capital call or new pref in a refinancing
3. Use the LP seat: written questions to TexSun on proceeds, taxes, pref and the extension rate, answered before Mar-27
4. Take the lessons to IC: the screens on the next slide go into the fund's checklist
Closing line: Taylor Commons will not return the $350K; **its value to the fund now is the discipline it adds to the next deal**
Footer 9pt: "Notes: Loan sizing at 1.25x DSCR, 8.0% debt yield and 65% LTV (team assumptions). Sources: AM model v1_2; team analysis."
```

### Historical Financials vs. Underwriting (build once the model has the actuals vs. IC tab)

```
Build or fill the slide titled "Historical Financials vs. Underwriting".
Takeaway: Year 3 NOI is $1.43M with the tax break and $0.94M without it, 42% of the $2.21M IC plan
Left panel (60%), header "IC Underwriting vs. Actual ($000s)": a table with rows Gross potential rent, Vacancy and
concessions, Bad debt, Other income, Effective gross income, Payroll, Repairs and turnover, Other opex, Property taxes,
Insurance, NOI; columns IC Year 2, Actual Year 2, IC Year 3, Actual Year 3 (T12 to Feb-26), Variance. Brick red for
negative variances. [Numbers come from the AM model's actuals vs. IC tab; do not type them in by hand.]
Known checks: IC Year 2 NOI $1,998K vs. actual $370K; IC Year 3 NOI $2,206K vs. actual $1,429K (taxes ~$28K under the HFC).
Right panel (35%), header "What Drove the Miss": three bullets naming the biggest variance lines from the table.
Closing line: [one sentence: the line items, not the tax break, explain most of the miss / or the reverse; decide once the table is filled]
```

---

## Batch 6: Sales Comparables (needs data first)

The slide is empty, and the assignment requires updated sales and cap rate comps. You need 5 to 8 trades with:
- buildings from 1975 to 1990
- east Fort Worth, Arlington or Haltom City
- sold 2025 to 2026
- property, sale date, units, year built, price, $/unit, cap rate (if known), buyer

Sources are CoStar, a broker call (note the name and date for the footer) or news reports. Once you have them:

```
Fill the slide titled "Sales Comparables" with a table of these trades: [paste rows]. Columns: #, Property, Date, Units,
Year Built, Price ($M), $/Unit, Cap Rate, Buyer. Add a "Weighted average" row and, below it, two brick-red rows:
"Taylor Commons base case: $13.4M to $15.6M, $60K to $70K/unit" and "Fund mark (Mar-26): $19.2M, $85.7K/unit".
Takeaway: [one sentence comparing TC's $60K to $70K/unit base case with the comp range]. Footer: sources and broker
name/date.
```

Until then, the only sourced range is the annual report's $91K to $137K/unit for stabilized trades. Taylor Commons
isn't stabilized, so don't use that range as its value.
