"""Slide: Location Overview (replaces the two prior 'Location Overview' slides).

Usage: python3 build_location.py <in.pptx> <out.pptx> [position]

Base images and marker positions come from the prior team's slides (base_v1.pptx slides 4 and 5): the satellite
map (2555x1216, shown at 0.72,1.3 / 11.09x4.9 with crops l .00803 r .07403 t .07014 b .07773) and the DFW map
(1598x876, shown at 0.67,1.3 / 11.98x4.9, crops l .01386 r .01474 t .14142 b .13362). Marker slide coordinates
are converted to image pixels, then to this slide. Drive times and school ratings are the prior team's labels
(GreatSchools scores). Context: inspection report (location section).
"""
import sys

from PIL import Image

from tc_common import *  # noqa: F401,F403

SRC, DST = sys.argv[1], sys.argv[2]
POS = int(sys.argv[3]) if len(sys.argv) > 3 else 3
IMG = os.path.join(HERE, "img")


def to_px(x, y, X0, Y0, W, H, crop, size):
    cl, cr, ct, cb = crop
    return ((cl + (x - X0) / W * (1 - cl - cr)) * size[0], (ct + (y - Y0) / H * (1 - ct - cb)) * size[1])


# ---------------------------------------------------------------- satellite map crop
SAT = Image.open(os.path.join(IMG, "sat_map.png")).convert("RGB")
S_ARGS = (0.72, 1.3, 11.09, 4.9, (0.00803, 0.07403, 0.07014, 0.07773), SAT.size)
px0, py0 = to_px(2.10, 1.30, *S_ARGS)
px1, py1 = to_px(10.55, 6.20, *S_ARGS)
SAT.crop((int(px0), int(py0), int(px1), int(py1))).save(os.path.join(IMG, "sat_crop.jpg"), quality=90)

prs = Presentation(SRC)
g = THEME["grid"]
L0, LW, R0, RW = g["L0"], g["LP1"] - g["L0"], g["RP0"], g["R1"] - g["RP0"]
MX, MY, MW, MH = L0, 1.74, LW, 4.25
for _ in range(2):
    drop_slide_titled(prs, "Location Overview")


def map_xy(x, y):
    """Prior-slide coordinates -> this slide's map frame."""
    px, py = to_px(x, y, *S_ARGS)
    return MX + (px - px0) / (px1 - px0) * MW, MY + (py - py0) / (py1 - py0) * MH


s = new_slide(prs, "Location Overview",
              "Eleven minutes east of downtown Fort Worth off I-30, but ringed by older garden competitors and "
              "low-rated schools")
add_slide_number(s)
header(s, L0, g["ROW1"], LW, "Immediate Surroundings", "A header")
add_image(s, os.path.join(IMG, "sat_crop.jpg"), MX, MY, MW, MH, name="A satellite map")

# pins: (label, x, y, kind)  kind: apt / school / retail
pins = [
    ("1", 6.93, 3.75, "apt"), ("2", 8.03, 3.72, "apt"), ("3", 8.62, 4.12, "apt"), ("4", 5.40, 4.19, "apt"),
    ("5", 6.90, 4.91, "apt"), ("6", 3.39, 4.98, "apt"),
    ("S1", 6.70, 3.50, "school"), ("S2", 5.32, 3.96, "school"), ("S3", 6.55, 4.76, "school"),
    ("S4", 5.08, 4.85, "school"),
    ("R1", 7.59, 3.66, "retail"), ("R2", 10.07, 4.20, "retail"), ("R3", 3.43, 2.65, "retail"),
]
style = {"apt": ("ink", "white", None), "school": ("white", "ink", "ink"), "retail": ("grey_mid", "white", None)}
D = 0.3
for lab, x, y, kind in pins:
    cx, cy = map_xy(x, y)
    fill, tc, ln = style[kind]
    b = box(s, cx - D / 2, cy - D / 2, D, D, fill=fill, line=ln or "white", line_w=1.0, shape=MSO_SHAPE.OVAL,
            name=f"A pin {lab}")
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_tf(tf, [[(lab, {"b": True, "c": tc})]], 10 if len(lab) > 1 else 11, align=PP_ALIGN.CENTER)
# subject
sx, sy = map_xy(6.45, 3.87)
box(s, sx - 0.13, sy - 0.13, 0.26, 0.26, fill="accent_fill", line="white", line_w=1.5, shape=MSO_SHAPE.OVAL,
    name="A subject dot")
lab = box(s, sx - 1.85, sy - 0.62, 1.6, 0.32, fill="accent_fill", line=None, name="A subject label")
lab.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
fill_tf(lab.text_frame, [[("Taylor Commons", {"b": True, "c": "white"})]], 12, align=PP_ALIGN.CENTER)
line(s, sx - 0.25, sy - 0.3, sx - 0.08, sy - 0.08, "accent", 1.25, name="A subject leader")
ix, iy = map_xy(6.68, 3.91)
add_image(s, os.path.join(IMG, "i30.png"), ix + 0.1, iy + 0.06, 0.26, 0.26, name="A I-30 shield")

# DFW inset (top left of the map)
DFW = Image.open(os.path.join(IMG, "dfw_map.png")).convert("RGB")
D_ARGS = (0.67, 1.3, 11.98, 4.9, (0.01386, 0.01474, 0.14142, 0.13362), DFW.size)
dx0, dy0 = to_px(2.3, 2.75, *D_ARGS)
dx1, dy1 = to_px(9.6, 5.45, *D_ARGS)
DFW.crop((int(dx0), int(dy0), int(dx1), int(dy1))).save(os.path.join(IMG, "dfw_crop.jpg"), quality=90)
IW, IH = 2.45, 0.91
IX, IY = MX + MW - IW - 0.12, MY + MH - IH - 0.12
box(s, IX - 0.03, IY - 0.03, IW + 0.06, IH + 0.06, fill="white", line="grey_mid", line_w=0.75,
    shape=MSO_SHAPE.RECTANGLE, name="A inset frame")
add_image(s, os.path.join(IMG, "dfw_crop.jpg"), IX, IY, IW, IH, name="A DFW inset")


def inset_xy(x, y):
    px, py = to_px(x, y, *D_ARGS)
    return IX + (px - dx0) / (dx1 - dx0) * IW, IY + (py - dy0) / (dy1 - dy0) * IH


for nm, x, y, colr in [("Taylor Commons", 4.815, 4.015, "accent_fill"), ("Monticello", 3.845, 4.095, "ink")]:
    cx, cy = inset_xy(x, y)
    box(s, cx - 0.06, cy - 0.06, 0.12, 0.12, fill=colr, line="white", line_w=0.75, shape=MSO_SHAPE.OVAL,
        name=f"A inset {nm}")
cx, cy = inset_xy(3.845, 4.095)
textbox(s, cx - 0.42, cy + 0.09, 0.84, 0.19, [[("Monticello", {"b": True})]], size=10, align=PP_ALIGN.CENTER,
        name="A inset label Monticello", fill="white")

# ---------------------------------------------------------------- right: key
header(s, R0, g["ROW1"], RW, "What Is Nearby", "B header")
y = 1.72


def group(title, rows, marker):
    global y
    textbox(s, R0, y, RW, 0.26, [[(title, {"b": True})]], size=12, name=f"B {title}")
    y += 0.27
    for lab, txt, note in rows:
        fill, tc, ln = style[marker]
        b = box(s, R0, y + 0.01, 0.22, 0.22, fill=fill, line=ln or fill, line_w=0.75, shape=MSO_SHAPE.OVAL,
                name=f"B key {lab}")
        tf = b.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        fill_tf(tf, [[(lab, {"b": True, "c": tc})]], 8 if len(lab) > 1 else 9, align=PP_ALIGN.CENTER)
        textbox(s, R0 + 0.34, y, RW - 1.25, 0.26, [[(txt, {})]], size=11.5, name=f"B item {lab}",
                anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, R1 - 0.95, y, 0.95, 0.26, [[(note, {"c": "grey_txt"})]], size=11.5, align=PP_ALIGN.RIGHT,
                name=f"B note {lab}", anchor=MSO_ANCHOR.MIDDLE)
        y += 0.25
    y += 0.07
    line(s, R0, y - 0.04, R0 + RW, y - 0.04, "hair", 0.75, name=f"B rule {title}")


R1 = g["R1"]
group("Competing apartments", [("1", "Luna Vista and The Crossing", ""), ("2", "The Griffin, Villas at Alameda", ""),
                               ("3", "Steeplechase and The Campbell", ""), ("4", "Quail Ridge", ""),
                               ("5", "Santa Barbara Gardens", ""), ("6", "City View", "")], "apt")
group("Schools (rating of 10)", [("S1", "Intl. Leadership of Texas K-8", "3; 4 min"),
                                 ("S2", "Nolan Catholic High", "n/a; 4 min"),
                                 ("S3", "Eastern Hills High", "2; 5 min"),
                                 ("S4", "Meadowbrook Middle", "4; 7 min")], "school")
group("Retail", [("R1", "Woodcreek Village: Albertsons, Home Depot", ""),
                 ("R2", "Eastchase Market: Target, ALDI, AMC", ""),
                 ("R3", "Beach Belknap: El Rancho, Whataburger", "")], "retail")

closing(s, [("Close to jobs and retail, but ", {}),
            ("competing with older garden communities on nearly every side", {"b": True})])
footer_notes(s, "Notes: Drive times and school ratings (GreatSchools) per the prior team's map; drive time to "
                "downtown Fort Worth is a team estimate.",
             "Sources: Google Maps; GreatSchools; team site visit and inspection report.")
speaker_notes(s, """Presenter: TBD

Taylor Commons sits on the east side of Fort Worth, just north of I-30, about eleven minutes from downtown. The inset shows it next to Monticello, the fund's other Fort Worth deal, which is on the west side of downtown.

Zooming in, you can see what surrounds it. Access is only from Bridge Road through Overlook Terrace, behind a row of office buildings, which is why visibility from the interstate is poor.

The numbered dark pins are competing apartment communities, and there are older garden properties on nearly every side: Luna Vista and The Crossing, The Griffin and Villas at Alameda, Steeplechase and The Campbell, Quail Ridge, Santa Barbara Gardens. Most are the same vintage as Taylor Commons and compete on price.

Retail is close: Woodcreek Village with Albertsons and Home Depot, Eastchase Market with Target, ALDI and an AMC, and Beach Belknap to the northwest. Schools are a weakness: the nearby public schools rate 2 to 4 out of 10.

So the location works for a workforce renter who wants to be near jobs and shopping, but the property has to win against a lot of similar product.

Reference notes:
- Pins placed from the prior team's map markers; drive times and school ratings are from that slide and should be re-checked.
- Inspection: access exclusively via Bridge Road to Overlook Terrace; site tucked behind highway frontage; office developments to the south.""")
move_slide(prs, s, POS)
prs.save(DST)
print("saved", DST, "at position", POS + 1)
