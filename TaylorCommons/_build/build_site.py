"""Slides: Asset Overview (rebuilt) + Site Visit 1 and 2 (from the team inspection, November 11).

Usage: python3 build_site.py <in.pptx> <out.pptx>
Drops the old 'Asset Overview' slide, inserts the new one at slide 3 and the two site visit slides at 5 and 6
(after 'Then vs. Now').

Facts: 5600 Cotswold Hills Dr., Fort Worth 76112; built 1979; 224 units (152 1BR / 72 2BR, box score 11/15/2026);
avg 767 SF (Control Sheet); closed 6/15/2023 at $25.9M ($115,625/unit, IC); sponsor TexSun Holdings; Cottonwood
running day-to-day operations (annual report); SMU $350K of $9.875M common; senior $16.5M W Financial bridge ~9%
due Mar-27; Pecos HFC ground lease 3/24/2025. Site visit text and photos: Taylor_Commons_Inspection.pdf.
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
photo(s, "aerial", L0, 1.74, LW, 2.72, "A aerial", focus=0.8)
pw = (LW - 0.24) / 3
for k, (key, nm) in enumerate([("leasing_office", "office"), ("pool", "pool"), ("exterior3", "building")]):
    photo(s, key, L0 + k * (pw + 0.12), 4.58, pw, 1.3, f"A photo {nm}")
caption(s, L0, 5.93, LW, "Aerial view; leasing office; pool; typical two-story building", "A caption")

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
             "sponsor photos; team site visit.")
speaker_notes(s, """Presenter: TBD

Taylor Commons is a 224-unit garden-style community on the east side of Fort Worth, in the Woodhaven submarket, just north of I-30. It was built in 1979: two-story buildings, 152 one-bedrooms and 72 two-bedrooms, averaging about 770 square feet, with a pool, a small fitness center and a resident lounge in the leasing office.

The fund closed in June 2023 at $25.9 million, about $116 thousand a unit. TexSun is the sponsor, but the preferred investor, Cottonwood, is now running the property day to day. Our position is $350 thousand of a $9.9 million common equity check.

Two things changed since closing and matter for the rest of the deck: the senior loan is now a $16.5 million bridge at about 9% due in March, and the land is owned by the Pecos HFC under a ground lease.

This is older, basic product. It wins on price and condition, which is what we went to look at on the site visit.

Reference notes:
- Fixes the prior version: the property is not '24 affordable units'; the Regulatory Agreement requires 50% of units for households under 80% AMI and 90% under 140% AMI.
- 171,715 rentable SF per the IC model (the prior slide showed 178,923 SF).""")
move_slide(prs, s, 2)


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
s = new_slide(prs, "Site Visit: Location, Curb Appeal and Units",
              "The property is hard to find and dated on arrival, and the vacant unit toured was dim and not cleaned")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "What a Prospect Sees", "A header")
photo_grid(s, [("sign", "The only sign on the frontage road", 0.5),
               ("office_view", "Office towers block the view from I-30", 0.5),
               ("gate", "Gate signage reads only up close", 0.5),
               ("exterior2", "Original 1979 stucco buildings", 0.5),
               ("kitchen", "Updated counters, white appliances", 0.5),
               ("living", "No light fixture in the living room", 0.5)], "A")
header(s, R0, g["ROW1"], RW, "Team Assessment", "B header")
yy1 = scorecard(s, 1.74, [
    ("Location and visibility", "Weak",
     "Tucked behind office buildings off I-30, one small \"Now Leasing\" sign, single entry off Overlook Terrace"),
    ("Curb appeal", "Fair",
     "Clean, landscaped entry, but 1979 stucco, faded striping and repetitive buildings read dated"),
    ("Unit interiors", "Fair",
     "New counters and backsplash, but tile floors, white appliances and no ceiling lights; unit not cleaned"),
], "B")
callout(s, R0, yy1, RW, 0.62, [("The agent ", {}), ("struggled to find it", {"b": True}), (" for her own interview", {})], "B callout", size=12)
closing(s, [("Visibility and curb appeal cost leads ", {}), ("before a prospect ever sees a unit", {"b": True})])
footer_notes(s, "Notes: Team site visit and secret shop, November 11.",
             "Sources: Team site visit and photos.")
speaker_notes(s, """Presenter: TBD

We visited the property and secret-shopped the leasing office on November 11.

Getting there is the first problem. The property sits behind a row of office buildings off I-30, so you cannot see it from the highway. The only directional sign on the frontage road is a small 'Now Leasing' sign with no property name on it, and the property signage at the gate only reads once you are almost there. The leasing agent told us she had trouble finding it for her own interview.

On arrival, the entry and leasing office are clean and landscaped, but the buildings are original 1979 stucco with dark trim, the parking striping is faded, and the buildings all look the same. Next to renovated neighbors like The Brittany, it reads dated.

Inside, the vacant unit we saw had new counters and backsplash, but tile floors, white appliances, and no light fixtures in the living room or bedrooms, so it was dim. It had not been cleaned. A renovated model unit with good light would show much better.

All of this costs leads before anyone sees a unit, which matches the funnel: leads doubled, tours did not.""")
move_slide(prs, s, 4)

# ======================================================================== Site Visit 2
s = new_slide(prs, "Site Visit: Amenities, Leasing and Marketing",
              "A strong leasing agent, but no model unit, short hours and a $175 fee package weaken every tour")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "The Tour", "A header")
photo_grid(s, [("pool", "Pool is clean; deck furniture is worn", 0.5),
               ("fitness", "Small gym, open office hours only", 0.5),
               ("lounge", "Lounge works as a waiting area", 0.5),
               ("trash", "Overflowing trash on the tour route", 0.5),
               ("hours", "Closed Sundays and 6 PM weekdays", 0.4),
               ("website", "Website and listings are current", 0.5)], "A")
header(s, R0, g["ROW1"], RW, "Team Assessment", "B header")
yy = scorecard(s, 1.74, [
    ("Amenities", "Fair", "Clean pool and lounge; small fitness center locked outside office hours"),
    ("Leasing and tour", "Weak",
     "Engaged, knowledgeable agent, but no model unit; tour of an occupied corner unit past overflowing trash"),
    ("Hours and marketing", "Fair",
     "Closes 6 PM weekdays, all Sunday; site and listings current; $500 to $1,000 referral bonuses"),
], "B")
callout(s, R0, yy, RW, 0.62, [("Fee package: ", {}), ("$175 a month", {"b": True}),
                              (" vs. $135 at The Brittany and $99 at Tides", {})], "B callout", size=12)
closing(s, [("The fixes are cheap (a model unit, longer hours, trash pickup), ", {}),
            ("but each competes for the same scarce cash", {"b": True})])
footer_notes(s, "Notes: Team site visit and secret shop, November 11. Fee package is trash, W/D, pest, Wi-Fi and "
                "liability waiver; The Brittany $135, Tides at 5400 $99.",
             "Sources: Team site visit; leasing offices; taylorcommons.com.")
speaker_notes(s, """Presenter: TBD

The amenities are serviceable. The pool is clean and a good size, but the deck furniture is worn. The fitness center is small and only open during leasing office hours, which does not help residents who work all day. The lounge in the leasing office is nice but works more like a waiting room than a place residents use.

The leasing agent was the best part of the visit: engaged, knew the product and the neighborhood, and sold the location well. But the tour hurt her. There is no model unit, so she showed us an occupied unit in the far corner of the property, and the route went past overflowing trash bins and unkept pet areas.

Hours are a real problem for this renter: the office closes at 6 on weekdays and is closed Sundays, and most prospects work full time. The website and the listing sites are current and look good, and there is a referral program paying $500 to $1,000.

One pricing point: the monthly fee package here, trash, washer and dryer, pest, Wi-Fi and the liability waiver, is about $175 a month. The Brittany's is $135 and Tides at 5400 is $99. That is on top of rents that TexSun already says are high on some floor plans.

None of these fixes are expensive. A model unit, longer hours and trash pickup are small dollars. But they compete for the same cash the property needs to turn its 49 vacant units.

Reference notes:
- The Brittany (TexSun, 280 units, 1984) has a model unit next to the leasing office overlooking the pool, with stainless appliances and vinyl plank.
- Fee packages from the inspection: Taylor Commons $175 (trash $20, W/D $60, pest $9, Spectrum $70, waiver $16); The Brittany $135; Tides at 5400 $99.""")
move_slide(prs, s, 5)

prs.save(DST)
print("saved", DST)
