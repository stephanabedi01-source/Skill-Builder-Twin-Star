"""Slides: Asset Overview (rebuilt) + Site Visit 1 and 2 (team site visit and secret shop, Nov-26).

PHOTO RULE: only the team's Nov-26 secret shop photos (img/v26_*.jpg) may be used. No photos from the Nov-25
inspection report, the IC deck, the sponsor, or prior groups.

Usage: python3 build_site.py <in.pptx> <out.pptx>
Drops the old 'Asset Overview' slide, inserts the new one at slide 3 and the two site visit slides at 5 and 6
(after 'Then vs. Now').

Facts: 5600 Cotswold Hills Dr., Fort Worth 76112; built 1979; 224 units (152 1BR / 72 2BR, box score 11/15/2026);
avg 767 SF (Control Sheet); closed 6/15/2023 at $25.9M ($115,625/unit, IC); sponsor TexSun Holdings; Cottonwood
running day-to-day operations (annual report); SMU $350K of $9.875M common; senior $16.5M W Financial bridge ~9%
due Mar-27; Pecos HFC ground lease 3/24/2025. Fee package and referral terms: Nov-25 inspection report text.
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"]
drop_slide_titled(prs, "Asset Overview", layout="5_CONTENT")


# ======================================================================== Asset Overview
s = new_slide(prs, "Asset Overview",
              "A 224-unit 1979 garden community in east Fort Worth, bought for $25.9M in Jun-23; SMU holds $350K "
              "of common")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "Taylor Commons Apartments", "A header")
photo(s, "v26_gate", L0, 1.74, LW, 2.72, "A hero", focus=0.62)
pw = (LW - 0.24) / 3
for k, (key, nm, fc) in enumerate([("v26_office", "office", 0.6), ("v26_model_living", "model", 0.55),
                                   ("v26_building", "building", 0.4)]):
    photo(s, key, L0 + k * (pw + 0.12), 4.58, pw, 1.3, f"A photo {nm}", focus=fc)
caption(s, L0, 5.93, LW, "Entry gate; leasing office; staged model unit; typical two-story building (Nov-26)",
        "A caption")

header(s, R0, g["ROW1"], RW, "Property", "B header")
rows_p = [
    [[("Address", {"b": True})], "5600 Cotswold Hills Dr."],
    [[("Submarket", {"b": True})], "East Fort Worth / Woodhaven"],
    [[("Built", {"b": True})], "1979; two-story garden"],
    [[("Units", {"b": True})], [("224: 152 one-bed, 72 two-bed", {}), ("(1)", {"sup": True})]],
    [[("Average size", {"b": True})], "767 SF"],
    [[("Amenities", {"b": True})], "Pool, fitness center, lounge"],
]
y = 1.70
for k, (a, b) in enumerate(rows_p):
    textbox(s, R0, y, 1.55, 0.26, [a if isinstance(a, list) else [(a, {})]], size=12, name=f"B k{k+1}",
            anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, R0 + 1.55, y, 3.08, 0.26, [b if isinstance(b, list) else [(b, {})]], size=12, name=f"B v{k+1}",
            anchor=MSO_ANCHOR.MIDDLE)
    y += 0.3
    if k < len(rows_p) - 1:
        line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"B rule {k+1}")
header(s, R0, y + 0.14, RW, "Investment", "C header")
rows_i = [
    ("Closed", [("Jun-23 at $25.9M ($116K/unit)", {})]),
    ("Sponsor", [("TexSun Holdings", {})]),
    ("Operations", [("Cottonwood running day to day", {})]),
    ("SMU position", [("$350K of $9.9M common", {"b": True})]),
    ("Senior debt", [("$16.5M bridge, ~9%, due Mar-27", {"c": "accent"})]),
    ("Land", [("Pecos HFC ground lease, Mar-25", {})]),
]
y += 0.56
for k, (a, b) in enumerate(rows_i):
    textbox(s, R0, y, 1.55, 0.26, [[(a, {"b": True})]], size=12, name=f"C k{k+1}", anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, R0 + 1.55, y, 3.08, 0.26, [b], size=12, name=f"C v{k+1}", anchor=MSO_ANCHOR.MIDDLE)
    y += 0.3
    if k < len(rows_i) - 1:
        line(s, R0, y - 0.02, R0 + RW, y - 0.02, "hair", 0.5, name=f"C rule {k+1}")

closing(s, [("Older, basic product that has to win on price and condition, ", {}),
            ("which is what the site visit tested", {"b": True})])
footer_notes(s, "Notes: (1) Box score 11/15/26.",
             "Sources: IC presentation; Deed of Trust and Ground Lease (3/24/25); Fund annual report 2025-26; "
             "team site visit (Nov-26).")
speaker_notes(s, """Presenter: TBD

Taylor Commons is a 224-unit garden-style community on the east side of Fort Worth, in the Woodhaven submarket, just north of I-30. It was built in 1979: two-story buildings, 152 one-bedrooms and 72 two-bedrooms, averaging about 770 square feet, with a pool, a small fitness center and a resident lounge in the leasing office.

The fund closed in June 2023 at $25.9 million, about $116 thousand a unit. TexSun is the sponsor, but the preferred investor, Cottonwood, is now running the property day to day. Our position is $350 thousand of a $9.9 million common equity check.

Two things changed since closing and matter for the rest of the deck: the senior loan is now a $16.5 million bridge at about 9% due in March, and the land is owned by the Pecos HFC under a ground lease.

This is older, basic product. It wins on price and condition, which is what we went to look at on the site visit.

Reference notes:
- Fixes the prior version: the property is not '24 affordable units'; the Regulatory Agreement requires 50% of units for households under 80% AMI and 90% under 140% AMI.
- 171,715 rentable SF per the IC model (the prior slide showed 178,923 SF).""")
place(prs, s, 2)


# ======================================================================== Site visit helper
def scorecard(s, y, items, name):
    for k, (area, grade, text) in enumerate(items):
        gc = "accent" if grade == "Weak" else "ink"
        textbox(s, R0, y, RW - 0.9, 0.26, [[(area, {"b": True})]], size=12, name=f"{name} area {k+1}")
        textbox(s, R1 - 0.9, y, 0.9, 0.26, [[(grade, {"b": True, "c": gc})]], size=12, align=PP_ALIGN.RIGHT,
                name=f"{name} grade {k+1}")
        textbox(s, R0, y + 0.28, RW, 0.78, [[(text, {})]], size=12, name=f"{name} text {k+1}")
        y += 1.12
        if k < len(items) - 1:
            line(s, R0, y - 0.08, R0 + RW, y - 0.08, "hair", 0.75, name=f"{name} rule {k+1}")
    return y


R1 = g["R1"]


def photo_grid(s, items, name):
    pw, ph = (LW - 0.24) / 3, 1.5
    for k, (key, cap, focus) in enumerate(items):
        x = L0 + (k % 3) * (pw + 0.12)
        y = 1.74 + (k // 3) * (ph + 0.5)
        photo(s, key, x, y, pw, ph, f"{name} photo {k+1}", focus=focus)
        textbox(s, x, y + ph + 0.05, pw, 0.4, [[(cap, {"i": True, "c": "grey_txt"})]], size=11,
                name=f"{name} caption {k+1}")


# ======================================================================== Site Visit 1
old1 = drop_slide_titled(prs, "Site Visit: Location, Curb Appeal and Units")
s = new_slide(prs, "Site Visit: Location, Signage and Arrival",
              "Taylor Commons is still hard to find: one unbranded sign on I-30, office towers in front and a gate "
              "seen only up close")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "What a Prospect Sees", "A header")
photo_grid(s, [("v26_sign", "The only sign on the frontage road", 0.5),
               ("v26_corner", "Office towers own the corner", 0.75),
               ("v26_overlook", "Overlook Terrace bends out of view", 0.7),
               ("v26_paving", "Cracked paving on the only way in", 0.7),
               ("v26_gate", "Gate sign reads only up close", 0.65),
               ("v26_office", "Clean, landscaped leasing office", 0.6)], "A")
header(s, R0, g["ROW1"], RW, "Team Assessment", "B header")
yy1 = scorecard(s, 1.74, [
    ("Visibility", "Weak",
     "Behind office towers north of I-30; the frontage-road sign carries no property name"),
    ("Wayfinding", "Weak",
     "Single entry via Overlook Terrace, a curving side street with cracked, patched paving"),
    ("Entry and office", "Fair",
     "Gate, flags and leasing office are clean and landscaped, but signage reads only at the gate"),
], "B")
callout(s, R0, yy1, RW, 0.62, [("Signage is ", {}), ("unchanged", {"b": True}),
                               (" since the Nov-25 visit", {})], "B callout", size=12)
closing(s, [("Visibility costs leads before a prospect sees a unit: ", {}),
            ("leads doubled in 2025 while tours stayed flat", {"b": True})])
footer_notes(s, "Notes: Leads 665 to 1,315 and tours 121 to 123, 26 weeks to Jun-25 vs. 26 weeks to Dec-25.",
             "Sources: Team site visit and secret shop (Nov-26); TexSun box scores.")
speaker_notes(s, """Presenter: TBD

We went back out to the property and secret-shopped it this month.

Getting there is still the first problem. Coming off I-30, the only cue is a small 'Now Leasing' sign on the frontage road with no property name on it. The corner is dominated by the glass office towers, and Overlook Terrace, the only way in, curves behind them, so you cannot see the property from the road. The street itself is cracked and patched.

At the gate it gets better: the gate, the flags and the leasing office are clean and well landscaped. But the Taylor Commons sign only reads once you are already at the gate.

None of the signage has changed since last year's visit. That matters because the funnel shows it: leads roughly doubled in 2025, but tours stayed flat. Prospects find the listing; fewer of them find, or choose to drive to, the property.

Reference notes:
- Leasing funnel (box score sums): leads 665 -> 1,315; tours 121 -> 123; 26 weeks to 6/2/25 vs. 26 weeks to 12/8/25.
- Photos: team secret shop, Nov-26 (only current-year photos are used in this deck).""")
place(prs, s, old1 if old1 is not None else 3)

# ======================================================================== Site Visit 2
old2 = drop_slide_titled(prs, "Site Visit: Amenities, Leasing and Marketing")
s = new_slide(prs, "Site Visit: Leasing Office, Model Unit and Upkeep",
              "A staged model unit now anchors the tour, but short hours and visible deferred maintenance undercut it")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "The Tour", "A header")
photo_grid(s, [("v26_hours", "Closed 6 PM weekdays and all Sunday", 0.45),
               ("v26_model_living", "New since Nov-25: staged 1BR model", 0.6),
               ("v26_model_kitchen", "New counters; white appliances, tile", 0.55),
               ("v26_landing2", "Stair landing propped on cinder blocks", 0.6),
               ("v26_hvac2", "Rusted, debris-filled condensers", 0.55),
               ("v26_siding", "Rotted siding behind a downspout", 0.5)], "A")
header(s, R0, g["ROW1"], RW, "Team Assessment", "B header")
yy = scorecard(s, 1.74, [
    ("Leasing office", "Fair",
     "Clean and inviting, but open only 9 to 6 weekdays and 10 to 5:30 Saturday; closed Sunday"),
    ("Model unit", "Good",
     "Staged 1BR model now shown; classic finishes: new counters, original cabinets, white appliances"),
    ("Exterior upkeep", "Weak",
     "Landings on cinder blocks, rusted HVAC condensers and rotted siding along resident walkways"),
], "B")
callout(s, R0, yy, RW, 0.62, [("Fee package: ", {}), ("$175 a month", {"b": True}),
                              (" vs. $135 at The Brittany and $99 at Tides", {}), ("(1)", {"sup": True})],
        "B callout", size=12)
closing(s, [("The unit now shows well; ", {}),
            ("the hours and the walk to it still cost leases", {"b": True})])
footer_notes(s, "Notes: (1) Trash, W/D, pest, Wi-Fi and liability waiver, as quoted at the Nov-25 visit. More photos "
                "in the appendix.",
             "Sources: Team site visit and secret shop (Nov-26); posted office hours.")
speaker_notes(s, """Presenter: TBD

Inside the gate, the biggest change since last year is a staged model unit. Last year the agent had to show an occupied unit in the far corner; now there is a furnished one-bedroom, the A2 plan, and it shows well. It is a classic unit, though: new counters and backsplash, but the original cabinets, white appliances and tile floors.

The leasing office itself is clean and inviting, but the hours have not changed: 9 to 6 on weekdays, 10 to 5:30 on Saturday, and closed Sunday, which is when people who work full time go apartment hunting. [Add: the office was closed when we arrived, if that is when the 'closed' sign photo was taken.]

The walk to the unit is where it falls apart. Several concrete stair landings have settled and cracked and are propped up on loose cinder blocks. Many of the AC condensers are rusted and packed with leaves, and some siding has rotted through. The landings are a safety and liability issue, not a cosmetic one.

Pricing: the monthly fee package was $175 at last year's visit, against $135 at The Brittany and $99 at Tides at 5400. [Confirm what was quoted this year.]

So the unit sells now; the hours and the condition of the walk still cost leases.

Reference notes:
- Office hours from the posted sign: Mon-Fri 9:00 to 6:00, Sat 10:00 to 5:30, Sun closed.
- Fee package (Nov-25 inspection): Taylor Commons $175 (trash $20, W/D $60, pest $9, Spectrum $70, waiver $16); The Brittany $135; Tides at 5400 $99.
- Full photo set (signage, office, model, foundation, HVAC, misc.) is in the appendix.""")
place(prs, s, old2 if old2 is not None else 4)

prs.save(DST)
print("saved", DST)
