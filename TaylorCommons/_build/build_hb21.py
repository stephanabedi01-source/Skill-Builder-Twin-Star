"""Slide: HB 21: What Ends and What Stays (HFC block, slide 3 of 4).

Usage: python3 build_hb21.py <in.pptx> <out.pptx> [position]

Facts and numbers:
  HB 21 signed 5/28/2025, effective immediately; properties owned by out-of-jurisdiction HFCs need local
    approval by 12/31/2026 or lose the exemption 1/1/2027 (The Real Deal 5/29/25; HB 21 summaries).
  Lawsuit: Texas Workforce Housing Coalition et al., filed 9/9/2025, Cameron County; no ruling found (Multifamily Dive).
  Fort Worth 2025 legislative agenda: supported requiring local approval for foreign HFCs (fortworthtexas.gov).
  Regulatory Agreement (3/24/2025): >=50% of units <80% AMI, >=90% <140% AMI; rents capped at AMI x % x 30% / 12;
    term = while HFC owns the land; termination by HFC + Grantor agreement, Grantor pays HFC costs (Sec. 4).
  Ground lease: 99 years from 3/24/2025, improvements revert to the HFC at expiry (Memorandum, recitals 4-5).
  HFC fees run-rate $55K = (3,646 + 935) x 12  (T12 Housing Authority annual + admin fees, Jun-25 onward).
  2027 taxes ~$520K = pre-HFC T12 property taxes Mar-24..Feb-25 (estimate; 2027 assessment not available).
  2027 NOI no relief = 1,429.3 - 491.9 = 937.4 ($000s); with relief = 1,429.3 (T12 Feb-26).
  Rent caps: 2026 Fort Worth-Arlington HMFA MFI $106,700 (TDHCA); 1BR 1.5 persons (75%), 2BR 3 persons (90%):
    1BR 106,700 x .75 x .80 x .30 / 12 = 1,600; 2BR 106,700 x .90 x .80 x .30 / 12 = 1,921 (estimates).
  Market rents (Apr-26 rent roll, unit-weighted): 1BR $935 (152 units), 2BR $1,130 (72 units).
"""
import sys

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 12

MFI = 106700
CAP_1BR = round(MFI * .75 * .80 * .30 / 12)     # 1,600
CAP_2BR = round(MFI * .90 * .80 * .30 / 12)     # 1,921
MKT_1BR = round((16*872.4375 + 78*863.3846 + 3*1150 + 38*1033.0263 + 1282 + 14*1040.5 + 2*1132.5) / 152)  # 935
MKT_2BR = round((38*1079.8158 + 2*1369 + 28*1157.5357 + 4*1291.5) / 72)                                 # 1,130
assert (CAP_1BR, CAP_2BR, MKT_1BR, MKT_2BR) == (1600, 1921, 935, 1130), (CAP_1BR, CAP_2BR, MKT_1BR, MKT_2BR)

prs = Presentation(SRC)
g = THEME["grid"]
s = new_content_slide(
    prs, "HB 21: What Ends and What Stays",
    "Unless Fort Worth approves the Pecos HFC by YE2026 the exemption ends, while the HFC's lease, fees and rent "
    "rules stay")
add_slide_number(s)

# ------------------------------------------------------------------ left: what changes on 1/1/2027
L0, LW = g["L0"], g["LP1"] - g["L0"]
section_header(s, L0, g["ROW1"], LW, "What Changes on 1/1/2027 Without Relief", "A header")
rows = [
    ["Property tax exemption", "In place; T12 taxes $28K",
     [("Ends; ~$520K a year", {"b": True}), ("(1)", {"sup": True})]],
    ["Rent and income rules", "50% of units <80% AMI, 90% <140%", "No change"],
    ["Pecos HFC ground lease", "99 years; HFC owns the land", "No change"],
    ["HFC fees", "~$55K a year", "No change"],
    ["Exit from the HFC", "Needs Pecos HFC consent; owner pays", "No change"],
]
grid_table(s, L0, 1.74, [2.0, 3.05, 2.3], ["", "Through YE2026", "From 1/1/2027"], rows, "A table",
           size=10.5, row_h=0.38, hair_rules=True, left_cols=(0, 1), accent_cols={(0, 2)})

# timeline to the deadline
section_header(s, L0, 4.30, LW, "Path to the Deadline", "A2 header")
events = [("Mar-25", "HFC and 9.5% bridge", "close"), ("May-25", "HB 21 signed,", "effective at once"),
          ("Sep-25", "Coalition sues;", "no ruling found"), ("Mar-26", "Loan extended", "12 months at ~9%"),
          ("Dec-26", "Local approval", "deadline"), ("Mar-27", "Bridge loan", "matures")]
TY = 5.16
x0, x1 = L0 + 0.55, g["LP1"] - 0.55
xs = [x0 + (x1 - x0) * k / (len(events) - 1) for k in range(len(events))]
line(s, x0, TY, xs[4], TY, "grey_mid", 1.25, name="A2 timeline")
line(s, xs[4], TY, xs[5], TY, "accent", 1.5, dash=MSO_LINE_DASH_STYLE.DASH, name="A2 timeline critical")
for k, ((d, l1, l2), xx) in enumerate(zip(events, xs)):
    crit = k >= 4
    box(s, xx - 0.08, TY - 0.08, 0.16, 0.16, fill="accent" if crit else "ink", line=None, shape=MSO_SHAPE.OVAL,
        name=f"A2 dot {k+1}")
    textbox(s, xx - 0.65, TY - 0.36, 1.3, 0.2, [[(d, {"b": True, "c": "accent" if crit else "ink"})]], size=10.5,
            align=PP_ALIGN.CENTER, name=f"A2 date {k+1}")
    textbox(s, xx - 0.65, TY + 0.15, 1.3, 0.4, [[(l1, {})], [(l2, {"c": "grey_txt", "i": True})]], size=10,
            align=PP_ALIGN.CENTER, name=f"A2 event {k+1}")

# ------------------------------------------------------------------ right: three paths + rent caps
R0, RW = g["RP0"], g["R1"] - g["RP0"]
section_header(s, R0, g["ROW1"], RW, "Three Ways 2027 Can Go", "B header", sup="(2)")
prow = [
    [[("Fort Worth approves", {"b": True})], [("Unlikely; city backed HB 21", {})], "$1.43M"],
    [[("Lawsuit wins", {"b": True})], [("Filed Sep-25; pending", {})], "$1.43M"],
    [[("No relief", {"b": True})], [("Default on 1/1/2027", {})], "$0.94M"],
]
grid_table(s, R0, 1.74, [1.55, 2.05, 1.03], ["Path", "Status", "2027 NOI"], prow, "B table",
           size=10.5, row_h=0.36, hair_rules=True, left_cols=(0, 1), accent_rows=(2,), highlight_rows=(2,))

section_header(s, R0, 3.32, RW, "Rent Caps vs. Market Rents", "C header", sup="(3)")
crow = [["One bedroom (152 units)", f"${CAP_1BR:,}", f"${MKT_1BR:,}", f"${CAP_1BR - MKT_1BR:,}"],
        ["Two bedroom (72 units)", f"${CAP_2BR:,}", f"${MKT_2BR:,}", f"${CAP_2BR - MKT_2BR:,}"]]
yb = grid_table(s, R0, 3.74, [1.9, 0.9, 0.9, 0.93], ["$ / month", "80% cap", "Market", "Room"], crow, "C table",
                size=10.5, row_h=0.32, hair_rules=True)
callout(s, R0, yb + 0.14, RW, 0.6,
        [("Caps sit ", {}), ("$665 to $791", {"b": True}), (" above market, so the cost is fees and compliance, "
                                                           "not rent", {})], "C callout", size=10.5)

# ------------------------------------------------------------------ closing, footer, notes
closing_line(s, [("Base case from 2027: NOI near ", {}), ("$0.94M against $1.49M of interest at 9%", {"b": True}),
                 (", with the loan due in March 2027", {})])
footer_notes(s,
             "Notes: (1) Pre-HFC T12 level (Mar-24 to Feb-25); 2027 assessment not yet out. (2) T12 Feb-26 NOI "
             "($1,429K) less the tax increase; HFC fees sit below NOI. (3) 2026 Fort Worth MFI of $106,700 at 1.5 "
             "persons per bedroom x 80% x 30% / 12 (estimate); market from Apr-26 rent roll.",
             "Sources: Regulatory Agreement, Ground Lease (3/24/25); HB 21; The Real Deal; Multifamily Dive; "
             "City of Fort Worth; TDHCA.")

speaker_notes(s, """Presenter: TBD

HB 21 was signed on May 28, 2025, two months after this deal moved into the Pecos HFC. It says an HFC can only exempt a property outside its home area if the local government signs off. For properties already in that position, the deadline is the end of this year. No approval, no exemption starting January 1, 2027.

On the left, what that does here. The tax exemption ends, and we estimate taxes go back to about $520 thousand a year. Everything else stays: the rent and income rules, the 99-year ground lease to Pecos, about $55 thousand a year of HFC fees, and the fact that getting out of the HFC needs Pecos to agree, at the owner's cost.

On the right, three ways 2027 can go. Fort Worth could approve the Pecos HFC, but the city lobbied for this exact restriction, so we do not count on it. A coalition of owners sued in September 2025 and we have not found a ruling. So our base case is no relief, and NOI drops to about $940 thousand.

One point in the sponsor's favor: the rent caps are not the problem. At 80% of area income a one-bedroom could charge about $1,600, and the market here is about $935. The HFC costs fees and paperwork, not rent.

So the base case for 2027 is about $940 thousand of NOI against almost $1.5 million of interest, with the loan due in March. That is the next slide.

Reference notes:
- HB 21 summary: properties owned by out-of-jurisdiction HFCs must get local approval by 12/31/2026 or lose the exemption on 1/1/2027; LIHTC properties are excluded (The Real Deal, 5/29/2025; HB 21 summaries).
- Lawsuit: Texas Workforce Housing Coalition and Post WB Apartments v. Bexar Appraisal District, Cameron County District Court, filed 9/9/2025 (Multifamily Dive, National Law Review). No ruling found as of Nov-26: confirm with TexSun counsel.
- Fort Worth 2025 state legislative agenda supported requiring local approval for HFCs operating outside their jurisdiction.
- Regulatory Agreement Sec. 1(d) and 2(a): 90% of units for households under 140% AMI, 50% under 80% AMI; rents capped at AMI x 80% (or 140%) x 30% / 12. Sec. 4: term runs while the HFC owns the land; ends if the ground lease ends or by agreement of the HFC and owner, owner pays HFC costs.
- Rent cap math: 2026 MFI $106,700 (TDHCA). 1BR at 1.5 persons = 75% of 4-person; 2BR at 3 persons = 90%. Estimates; confirm with the Novogradac calculator named in the agreement.
- TexSun (Q1-26): no further clarity on the exemption; counsel checking in periodically.""")

move_slide(prs, s, POS)
prs.save(DST)
print("saved", DST, "at position", POS + 1)
