"""Slide: Operations: Where Leasing Broke Down.

Usage: python3 build_operations.py <in.pptx> <out.pptx> [position]

Occupancy (physical):
  Jul-24..Dec-25 monthly averages of weekly box score (TC_Box_Score_Data_4.13.26.xlsx, row 'Taylor Commons');
    weekly peak 90.2% (week of 7/21/2025).
  Jan-26..Apr-26 month-end (Trending_Occupancy_5.xlsx): 75.9%, 74.6%, 73.2%, 74.1%.
  May-26..Oct-26: no sponsor data. Nov-26: 74.9% (AM model v1_1 'Rent Roll' box score 11/15/2026).
  IC stabilized occupancy 95% (IC model vacancy 5%, Growth Assumptions 2024+).
Leasing funnel, 26 weeks 12/9/24..6/2/25 vs. 6/9/25..12/8/25 (box score sums):
  leads 665 -> 1,315; tours 121 -> 123; applications 76 -> 68; net leases 59 -> 35; move-ins 57 -> 31;
  move-outs 55 -> 44.
Make-ready: 56 vacant, 49 not ready (box score 11/15/2026). TexSun Q1-26: turning 1-2 units a month to manage cash.
Trade-out 2026 YTD (Lease_Trade-out_12.xlsx): new leases averaged 159 days vacant; renewals -4% effective.
Retention Oct-25..Apr-26 (Resident_Retention.xlsx): 145 expiring, 58 renewals (40%), 41 early move-outs (28%).
"""
import sys

from tc_common import *  # noqa: F401,F403
from deck_kit import line_ser, cat_ax, val_ax, add_chart

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 4

months = ["Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Jan-25", "Feb-25", "Mar-25", "Apr-25",
          "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26",
          "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26"]
occ = [.721, .743, .753, .780, .817, .871, .885, .879, .869, .837, .838, .880, .895, .885, .880, .856, .804, .797,
       .759, .746, .732, .741, None, None, None, None, None, None, .749]
assert len(months) == len(occ) == 29

prs = Presentation(SRC)
g = THEME["grid"]
s = new_slide(
    prs, "Operations: Where Leasing Broke Down",
    "Occupancy peaked near 90% in Jul-25 and fell to 74.9%, with 49 of 56 vacant units not ready to lease")
add_slide_number(s)

# ------------------------------------------------------------------ left: occupancy trend
L0, LW = g["L0"], g["LP1"] - g["L0"]
header(s, L0, g["ROW1"], LW, "Physical Occupancy, Jul-24 to Nov-26", "A header", sup="(1)")
VMIN, VMAX = 0.65, 1.0
P = Plot(0.75, 1.95, 6.85, 3.3, n=len(months), vmin=VMIN, vmax=VMAX)
vals = [v for v in occ]


def body():
    ser = line_ser(0, "Occupancy", months, vals, "accent", marker="circle", msize=4, mfill="accent", width=2.0)
    return (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{ser}'
            f'<c:marker val="1"/><c:axId val="4001"/><c:axId val="4002"/></c:lineChart>'
            + cat_ax(4001, 4002) + val_ax(4002, 4001, VMIN, VMAX))


add_chart(s, P.x, P.y, P.w, P.h, months, [("Occupancy", vals)], body, "A occupancy chart")
xs = [P.center(k) for k in range(len(months))]
# gridlines and axis labels
for v in (0.70, 0.80, 0.90, 1.00):
    line(s, P.x, P.v(v), P.x + P.w, P.v(v), "hair", 0.5, name=f"A grid {int(v*100)}")
    textbox(s, P.x - 0.55, P.v(v) - 0.1, 0.48, 0.2, [[(f"{v:.0%}", {"c": "grey_txt"})]], size=11,
            align=PP_ALIGN.RIGHT, name=f"A axis {int(v*100)}")
# IC stabilized occupancy claim line
line(s, P.x, P.v(0.95), P.x + P.w, P.v(0.95), "ink", 1.0, dash=MSO_LINE_DASH_STYLE.DASH, name="A IC line")
textbox(s, P.x + 0.05, P.v(0.95) - 0.27, 2.6, 0.22, [[("IC stabilized occupancy 95%", {"b": True})]], size=11,
        name="A IC label")
# data gap May-26 to Oct-26
k0, k1 = months.index("Apr-26"), months.index("Nov-26")
line(s, xs[k0], P.v(occ[k0]), xs[k1], P.v(occ[k1]), "grey_mid", 1.0, dash=MSO_LINE_DASH_STYLE.ROUND_DOT,
     name="A gap line")
textbox(s, xs[k0] + 0.05, P.v(0.741) + 0.08, 1.6, 0.22, [[("No sponsor data", {"c": "grey_txt", "i": True})]],
        size=11, name="A gap label")
# HFC refinancing marker
kh = months.index("Mar-25")
line(s, xs[kh], P.v(0.95) + 0.05, xs[kh], P.y + P.h, "grey_mid", 0.75, dash=MSO_LINE_DASH_STYLE.DASH, name="A HFC line")
textbox(s, xs[kh] + 0.06, P.y + P.h - 0.3, 1.9, 0.22, [[("HFC and bridge close", {"c": "grey_txt", "i": True})]],
        size=11, name="A HFC label")
# labels on key points
kp = months.index("Jul-25")
pill(s, xs[kp] - 0.85, P.v(0.895) - 0.52, 1.7, 0.32, [("90.2% peak, Jul-25", {"b": True})], "A peak pill",
     line_color="ink", size=11)
textbox(s, xs[k1] - 0.6, P.v(0.749) + 0.1, 0.75, 0.22, [[("74.9%", {"b": True, "c": "accent"})]], size=12,
        align=PP_ALIGN.CENTER, name="A last label")
textbox(s, xs[0] - 0.3, P.v(0.721) - 0.3, 0.8, 0.22, [[("72.1%", {"b": True})]], size=12, name="A first label")
# month labels
for k, mth in enumerate(months):
    if (k % 3 == 0 and k != 27) or k == len(months) - 1:
        textbox(s, xs[k] - 0.35, P.y + P.h + 0.06, 0.7, 0.22, [[(mth, {"c": "grey_txt"})]], size=11,
                align=PP_ALIGN.CENTER, name=f"A month {k+1}")

# ------------------------------------------------------------------ right: leasing funnel
R0, RW = g["RP0"], g["R1"] - g["RP0"]
header(s, R0, g["ROW1"], RW, "Leasing Funnel, 26 Weeks", "B header", sup="(2)")
rows = [
    ["Leads", "665", "1,315"],
    ["Tours", "121", "123"],
    ["Applications", "76", "68"],
    ["Net leases", "59", "35"],
    ["Move-ins", "57", "31"],
    ["Move-outs", "55", "44"],
]
yb = grid_table(s, R0, 1.74, [1.95, 1.3, 1.38], ["", "To Jun-25", "To Dec-25"], rows, "B table",
                size=12, row_h=0.33, hair_rules=True, accent_cols={(3, 2), (4, 2)})
header(s, R0, yb + 0.22, RW, "Why Units Sit Empty", "C header", sup="(3)")
items = [
    [[("Make-ready: ", {"b": True}), ("49 of 56 vacant units not ready", {})]],
    [[("Pricing: ", {"b": True}), ("new leases sat 159 days vacant", {})]],
    [[("Retention: ", {"b": True}), ("40% of expiring leases renewed", {})]],
]
numbered_items(s, R0, yb + 0.68, RW, items, [0.3, 0.3, 0.3], name="C item", size=12, num_d=0.24, gap=0.1)

# ------------------------------------------------------------------ closing, footer, notes
closing(s, [("Leads doubled but leases fell: ", {}),
            ("without cash to turn units, the property cannot lease its way back to 90%", {"b": True})])
footer_notes(s,
             "Notes: (1) Monthly average of weekly box scores to Dec-25; month-end Jan-26 to Apr-26; box score "
             "11/15/26. (2) 12/9/24 to 6/2/25 vs. 6/9/25 to 12/8/25. (3) Nov-26 box score; 2026 trade-out; "
             "retention Oct-25 to Apr-26.",
             "Sources: TexSun box scores, trending occupancy, trade-out and retention reports; IC model.")

speaker_notes(s, """Presenter: TBD

Before the refinancing story, the operating story, because the property was struggling on the ground first.

On the left is physical occupancy over the last two and a half years. It climbed from about 72% in mid-2024 to a peak just above 90% in July 2025, right after the HFC and bridge loan closed. Then it rolled over: 80% by December, about 74% through the spring, and 74.9% in last week's box score. The underwriting assumed 95% at stabilization; the property has never been close. We have no sponsor data between April and October of this year, so that stretch is drawn as a dotted line.

On the right, the leasing funnel. In the second half of 2025, leads doubled, to more than 1,300. Tours stayed flat, and net leases fell from 59 to 35. So people were looking, but they were not getting to the door, and when they did, units were not ready.

That is the core problem. 49 of the 56 vacant units are not ready to lease, because the property only has cash to turn one or two a month. New leases this year sat empty for 159 days on average. TexSun also says some floor plans are priced above Brittany and Ryan, which are more renovated. And only 40% of residents with expiring leases renewed over the last seven months.

Leads are not the problem. Cash to turn units is. That ties straight back to the debt service.

Reference notes:
- Weekly peak 90.2% (week of 7/21/2025); annual report cites 88.8% as the summer 2025 peak (monthly basis).
- Trending occupancy 2026 month-end: Jan 75.9%, Feb 74.6%, Mar 73.2%, Apr 74.1%.
- Box score 11/15/2026: 167 of 223 occupied (74.9%), 78.5% leased; 56 vacant (49 not ready, 7 ready).
- Lead-to-tour conversion fell from 18% to 9%; about 55% of tours applied (TexSun says ~60%).
- Trade-out 2026: 16 new leases at $1,037 effective vs. $1,018 prior (+1.8%); 21 renewals at $1,057 vs. $1,101 prior effective (-4%).
- Retention Oct-25 to Apr-26: 145 expiring, 58 renewals, 41 early move-outs, 20 notices to vacate. The annual report's 57% retention uses a different basis.
- Delinquency down to $16K at Mar-26 month end (TexSun); evictions slowed by 30 to 90 day writ delays.""")

place(prs, s, POS)
prs.save(DST)
print("saved", DST)
