# Slide log: status of every slide

## Current order (agenda order, set Nov-18 by build_agenda.py)

Agenda sections and pages: 1 Asset Overview (3); 2 Site Visit Findings (4 to 5); 3 Location & Market (6 to 9:
Location, Migration, Market, Submarket); 4 Performance vs. Underwriting (10 to 12: Then vs. Now, Operations,
Historical Financials); 5 Comparables (13 to 14: rent, sales); 6 Capital Structure & 2027 Maturity (15 to 18:
Refinancing trade, HB 21, 2027 cliff, Valuation); 7 Renovation Plan (19); 8 What Went Wrong (20 to 21); 9 Lessons
for the Fund (22 to 23, incl. Final Recommendation); 10 Appendix (24 to 25).
`build_agenda.py` reorders slides by title into this order and computes the page ranges, so rerun it after adding,
removing or retitling slides. The table below uses the older numbering; match slides by title.


Slide numbers are as of the v2 deck (25 slides). "Done" slides are on the new master, 12pt, sourced, rendered and
checked; edit them only to fix numbers or apply team feedback, and keep their builder in sync if you use scripts.
"Prior group" slides are leftovers from the earlier template or another deal and should be rebuilt or removed.

| # | Title | Status | Builder | Takeaway (claim) |
|---|---|---|---|---|
| 1 | Cover | Done (master) | build_master.py | none |
| 2 | Agenda | Done | build_agenda.py | Three years into a five-year hold, the common equity is out of the money; this update shows why and what the fund does next |
| 3 | Asset Overview | Done | build_site.py | A 224-unit 1979 garden community in east Fort Worth, bought for $25.9M in Jun-23; SMU holds $350K of common |
| 4 | Location Overview | Done | build_location.py | Eleven minutes east of downtown Fort Worth off I-30, but ringed by older garden competitors and low-rated schools |
| 5 | Then vs. Now: IC Underwriting vs. Today | Done | build_snapshot.py | Three years in, occupancy is 75% vs. 86% at IC and T12 NOI is $1.43M vs. $2.21M underwritten, even with the tax break |
| 6 | Site Visit: Location, Curb Appeal and Units | Done | build_site.py | The property is hard to find and dated on arrival, and the vacant unit toured was dim and not cleaned |
| 7 | Site Visit: Amenities, Leasing and Marketing | Done | build_site.py | A strong leasing agent, but no model unit, short hours and a $175 fee package weaken every tour |
| 8 | Operations: Where Leasing Broke Down | Done | build_operations.py | Occupancy peaked near 90% in Jul-25 and fell to 74.9%, with 49 of 56 vacant units not ready to lease |
| 9 | Migration Trends | Done | build_migration.py | Fort Worth still adds about 20,000 people a year, but Tarrant County now loses domestic movers; growth comes from births and immigration |
| 10 | Market Overview: Dallas-Fort Worth | Done | build_market.py | DFW is absorbing more apartments than it delivers and vacancy is down to 9.4%, but effective rents are still 2.0% lower than a year ago |
| 11 | Submarket: East Fort Worth and the Comp Set | Done | build_market.py | East Fort Worth runs about 90% occupied, but Taylor Commons' older comp set averages 77% and comp rents fell 8% to 16% since IC |
| 12 | Current Rent Comparables vs. Underwriting | Prior group, empty | none | To build: rent by floor plan vs. Brittany, Ryan and others; comp photos in the inspection report; IC comps in the AM model "UW Comps" tab |
| 13 | Current Sales Comparables vs. Underwriting | Prior group, empty | none | To build: needs the team's sales comp outreach; annual report range $91K to $137K/unit, submarket cap ~5.9% |
| 14 | Historical Financials vs. Underwriting | Prior group, empty | none | Planned: bridge from IC Year 3 NOI $2.21M to actual $0.94M (taxes paid) by line: occupancy, rent and concessions, bad debt, other income, payroll and opex, taxes and insurance |
| 15 | The March 2025 Refinancing Trade | Done | build_refi_trade.py | The HFC refinancing added $1.04M of annual interest to save $428K of taxes net of fees, a $612K yearly loss |
| 16 | HB 21: What Ends and What Stays | Done | build_hb21.py | Unless Fort Worth approves the Pecos HFC by YE2026 the exemption ends, while the HFC's lease, fees and rent rules stay |
| 17 | The 2027 Cliff and Refinancing Gap | Done | build_cliff.py | At $0.94M of 2027 NOI a new loan sizes to about $9.4M against $16.5M due in March 2027, a $7.1M gap |
| 18 | Valuation and Recovery | Done | build_valuation.py | Without relief the asset is worth $13.4M to $15.6M, below the $16.5M loan, so SMU's $350K recovers $0 |
| 19 | What Went Wrong? | Prior group, empty | none | Likely fold into an executive summary or a "three things went wrong" slide |
| 20 | HFC Update | Prior group, empty | none | Superseded by slides 15 to 17; remove |
| 21 | Renovation Plan | Prior group, empty | none | To build: $1.8M capex budget (Control Sheet) vs. capex actually spent (T12 capex $207K) and 12 renovated-layout units |
| 22 | How Can the Fund Avoid the Same Mistakes | Prior group, empty | none | Lessons: screens for structure changes after closing, rate-lock expiries, sponsor liquidity, consent rights |
| 23 | Final Recommendation | Prior group, empty | none | "What the fund does now": the team wants this creative but rigorous; draft idea is hold at $0, fund no new capital, and press TexSun on the open questions |
| 24 | Appendix | Divider | none | |
| 25 | Asset Photos | Prior group, empty | none | Fill with remaining photos or drop |

## Build order used for the scripted slides (from the original v1 deck)

`build_master.py` (rebuild master on the v1 file) -> `build_refi_trade.py 11` -> `build_hb21.py 12` ->
`build_cliff.py 13` -> `build_valuation.py 14` -> `build_snapshot.py 2` -> `build_operations.py 3` ->
`build_site.py` -> `build_location.py 3` -> `build_migration.py` -> `build_market.py`. On the latest deck, any single builder can be
rerun on its own; it replaces its slide in place.

## Feedback the user gave while building (keep honoring it)

- Garamond deck; clean, not used by another group: no navy, no red swoosh.
- Header needs something up top: thin charcoal top bar plus running head; title and takeaway shifted up; the rule
  under the takeaway is plain grey (no red segment).
- Notes and sources go in the footer next to the logo, and only a little of them; keep the footer small.
- Body text around 12pt.
- Maps make slides more interesting: use real or illustrative maps where they help (location, migration).
- Lean into the HFC and refinancing story; it is the centerpiece.
