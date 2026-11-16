"""Rebuild the slide master for the Taylor Commons AM deck.

Usage: python3 build_master.py <in.pptx> <out.pptx>

Replaces the shared SMU navy/red template look with a clean charcoal + brick system (Garamond kept):
  * theme colors: navy 354CA1 -> ink 22262B, SMU red CC0035 -> brick B4442C, light blue tints -> warm greys
  * content layouts: removes the corner swoosh image and the navy/red title tab, aligns title and takeaway
    to the 0.5in margin, makes the takeaway an italic grey line, tidies the footer (fund name | page number)
  * logo recolored from navy to charcoal (layouts only; the cover keeps the white logo)
  * TITLE layout: solid charcoal background instead of the gradient photo, short brick rule under the title
  * existing slides: same hex swaps so teammates' slides pick up the new palette; cover text re-laid out
"""
import io
import re
import sys
import zipfile

from PIL import Image

SRC, DST = sys.argv[1], sys.argv[2]

INK, SLATE, BRICK = "22262B", "4A5560", "B4442C"
HEX_SWAP = {"354CA1": INK, "CC0035": BRICK, "AEB9E0": "C9CDD1", "E9EDF7": "F2F1EE", "1F2430": INK}

EMU = 914400
def e(inches):
    return str(int(round(inches * EMU)))


def swap_hex(x):
    for a, b in HEX_SWAP.items():
        x = x.replace(f'val="{a}"', f'val="{b}"')
    return x


def drop_element(x, tag, name_pat):
    """Remove every <tag>...</tag> element whose cNvPr name/descr matches name_pat."""
    out, pos = [], 0
    pat = re.compile(rf"<{tag}>")
    while True:
        m = pat.search(x, pos)
        if not m:
            out.append(x[pos:]); break
        end = x.find(f"</{tag}>", m.start()) + len(f"</{tag}>")
        block = x[m.start():end]
        out.append(x[pos:m.start()])
        if not re.search(name_pat, block[:600]):
            out.append(block)
        pos = end
    return "".join(out)


def set_xfrm_for_ph(x, ph_marker, off=None, ext=None):
    """Change <a:off>/<a:ext> of the <p:sp> containing ph_marker (e.g. 'idx="106"')."""
    i = x.find(ph_marker)
    if i < 0:
        return x
    s = x.rfind("<p:sp>", 0, i)
    t = x.find("</p:sp>", i)
    block = x[s:t]
    if off:
        block = re.sub(r'<a:off x="\d+" y="\d+"/>', f'<a:off x="{e(off[0])}" y="{e(off[1])}"/>', block, count=1)
    if ext:
        block = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', f'<a:ext cx="{e(ext[0])}" cy="{e(ext[1])}"/>', block, count=1)
    return x[:s] + block + x[t:]


def edit_content_layout(x):
    # 1. corner swoosh picture and the navy/red title tab
    x = drop_element(x, "p:pic", r"wedge\.png")
    x = drop_element(x, "p:sp", r'name="Shape 0"|name="Shape 1"')
    # 2. title + takeaway to the 0.5in margin
    x = set_xfrm_for_ph(x, 'idx="106"', off=(0.50, 0.19), ext=(9.10, 0.55))
    tk = 'idx="109"' if 'idx="109"' in x else 'idx="107"'   # takeaway slot differs by layout
    x = set_xfrm_for_ph(x, tk, off=(0.50, 0.70), ext=(12.33, 0.34))
    # takeaway: italic slate 14pt
    i = x.find(tk)
    if i > 0:
        s = x.rfind("<p:sp>", 0, i); t = x.find("</p:sp>", i)
        block = x[s:t]
        block = re.sub(r'(<a:lvl1pPr[^>]*>(?:<a:buNone/>)?)<a:defRPr[^>]*?(?:/>|>.*?</a:defRPr>)',
                       rf'\1<a:defRPr sz="1400" b="0" i="1"><a:solidFill><a:srgbClr val="{SLATE}"/></a:solidFill>'
                       '<a:latin typeface="Garamond"/><a:ea typeface="Garamond"/><a:cs typeface="Garamond"/></a:defRPr>',
                       block, count=1, flags=re.S)
        x = x[:s] + block + x[t:]
    # 3. footer: fund name right-aligned, page number at the right margin
    i = x.find('name="Text 3"')
    if i > 0:
        s = x.rfind("<p:sp>", 0, i); t = x.find("</p:sp>", i)
        block = x[s:t]
        block = re.sub(r'<a:off x="\d+" y="\d+"/><a:ext cx="\d+" cy="\d+"/>',
                       f'<a:off x="{e(10.05)}" y="{e(6.66)}"/><a:ext cx="{e(2.23)}" cy="{e(0.26)}"/>', block, count=1)
        block = block.replace('algn="ctr"><a:buNone/>', 'algn="r"><a:buNone/>')
        block = block.replace("<a:t>|   SMU Cox Real Estate Impact Fund</a:t>",
                              "<a:t>SMU Cox Real Estate Impact Fund   |</a:t>")
        x = x[:s] + block + x[t:]
    x = set_xfrm_for_ph(x, 'type="sldNum"', off=(12.33, 6.66), ext=(0.50, 0.26))
    # footer rule and logo up: the band 6.60 to 7.45 holds notes and sources on each slide
    x = re.sub(r'(name="Shape 2".*?<a:off x="\d+" )y="\d+"', rf'\1y="{e(6.58)}"', x, count=1, flags=re.S)
    x = re.sub(r'(name="Image 1".*?<a:off x="\d+" )y="\d+"', rf'\1y="{e(6.70)}"', x, count=1, flags=re.S)
    x = x.replace("</p:spTree>", header_shapes() + "</p:spTree>", 1)
    return swap_hex(x)


def header_shapes():
    """Header furniture: charcoal top bar, running head at top right, hairline rule under the takeaway with a
    short brick lead segment at the left margin."""
    def sp(id_, name, x, y, w, h, geom, fill=None, ln=None, txt=None):
        f = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
        l = (f'<a:ln w="{int(ln[1]*12700)}"><a:solidFill><a:srgbClr val="{ln[0]}"/></a:solidFill></a:ln>'
             if ln else "<a:ln><a:noFill/></a:ln>")
        body = ('<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' if not txt
                else '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>'
                     '<a:p><a:pPr algn="r"/><a:r><a:rPr lang="en-US" sz="1000" i="1"><a:solidFill><a:srgbClr val="5F6670"/>'
                     '</a:solidFill><a:latin typeface="Garamond"/><a:ea typeface="Garamond"/><a:cs typeface="Garamond"/>'
                     f'</a:rPr><a:t>{txt}</a:t></a:r></a:p></p:txBody>')
        return (f'<p:sp><p:nvSpPr><p:cNvPr id="{id_}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
                f'<p:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
                f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{f}{l}</p:spPr>{body}</p:sp>')
    return (sp(91, "Header top bar", 0, 0, 13.333, 0.08, "rect", fill=INK)
            + sp(92, "Header running head", 8.83, 0.24, 4.0, 0.26, "rect",
                 txt="Taylor Commons  |  Asset Management Update")
            + sp(93, "Header rule", 0.50, 1.12, 12.33, 0, "line", ln=("D5D8DB", 0.75))
            + sp(94, "Header rule lead", 0.50, 1.12, 1.10, 0, "line", ln=(BRICK, 2.25)))


def edit_title_layout(x):
    # solid charcoal background in place of the gradient photo
    x = re.sub(r"<p:bg>.*?</p:bg>",
               f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{INK}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
               x, flags=re.S)
    # title placeholder lower and wider; brick rule under it
    x = set_xfrm_for_ph(x, 'idx="101"', off=(0.75, 2.05), ext=(10.5, 1.45))
    rule = (f'<p:sp><p:nvSpPr><p:cNvPr id="90" name="Cover rule"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{e(0.78)}" y="{e(3.62)}"/><a:ext cx="{e(1.40)}" cy="0"/></a:xfrm>'
            f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="28575">'
            f'<a:solidFill><a:srgbClr val="{BRICK}"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>')
    x = x.replace("</p:spTree>", rule + "</p:spTree>")
    # move the white logo to the bottom-right margin
    x = re.sub(r'(<p:cNvPr id="2" name="Image 0".*?<a:off )x="\d+" y="\d+"', rf'\1x="{e(11.33)}" y="{e(6.55)}"', x,
               flags=re.S)
    x = re.sub(r'(<p:cNvPr id="2" name="Image 0".*?<a:ext )cx="\d+" cy="\d+"', rf'\1cx="{e(1.50)}" cy="{e(0.44)}"', x,
               flags=re.S)
    return swap_hex(x)


def edit_theme(x):
    x = x.replace('<a:clrScheme name="SMU Cox Alts">', '<a:clrScheme name="Taylor Commons AM">')
    m = re.search(r"<a:clrScheme.*?</a:clrScheme>", x, flags=re.S)
    cs = m.group(0)
    for slot, val in {"dk2": INK, "lt2": "F2F1EE", "accent1": INK, "accent2": BRICK, "accent3": "C9CDD1",
                      "accent4": SLATE, "accent5": "8C939A", "accent6": "D5D8DB", "hlink": SLATE,
                      "folHlink": "8C939A"}.items():
        cs = re.sub(rf"<a:{slot}><a:srgbClr val=\"[0-9A-F]+\"/></a:{slot}>",
                    f'<a:{slot}><a:srgbClr val="{val}"/></a:{slot}>', cs)
    return x.replace(m.group(0), cs)


def recolor_logo(data, hexcol):
    im = Image.open(io.BytesIO(data)).convert("RGBA")
    r, g, b = int(hexcol[0:2], 16), int(hexcol[2:4], 16), int(hexcol[4:6], 16)
    px = im.load()
    for yy in range(im.height):
        for xx in range(im.width):
            a = px[xx, yy][3]
            px[xx, yy] = (r, g, b, a)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


def edit_cover_slide(x):
    """Slide 1: title, presenters and fund line re-laid out on the new cover (no overlap)."""
    # title text box (placeholder inherits layout position; drop any slide-level xfrm override)
    i = x.find('type="title"')
    if i > 0:
        s = x.rfind("<p:sp>", 0, i); t = x.find("</p:sp>", i)
        block = re.sub(r"<p:spPr>.*?</p:spPr>", "<p:spPr/>", x[s:t], count=1, flags=re.S)
        x = x[:s] + block + x[t:]
    x = re.sub(r'(name="Text 3".*?<a:off )x="\d+" y="\d+"(/><a:ext )cx="\d+" cy="\d+"',
               rf'\1x="{e(0.78)}" y="{e(3.85)}"\2cx="{e(10.0)}" cy="{e(0.70)}"', x, count=1, flags=re.S)
    x = re.sub(r'(name="Text 4".*?<a:off )x="\d+" y="\d+"(/><a:ext )cx="\d+" cy="\d+"',
               rf'\1x="{e(0.78)}" y="{e(6.62)}"\2cx="{e(8.0)}" cy="{e(0.35)}"', x, count=1, flags=re.S)
    x = x.replace('sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/>', 'sz="1400" b="0"><a:solidFill><a:srgbClr val="FFFFFF"/>')
    x = x.replace("<a:t>SMU Cox Real Estate Impact Fund | October 2026</a:t>", "<a:t>SMU Cox Real Estate Impact Fund | November 2026</a:t>")
    return x


zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    n = item.filename
    if n == "ppt/theme/theme1.xml":
        data = edit_theme(data.decode("utf8")).encode("utf8")
    elif n == "ppt/slideLayouts/slideLayout1.xml":
        data = edit_title_layout(data.decode("utf8")).encode("utf8")
    elif re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n):
        data = edit_content_layout(data.decode("utf8")).encode("utf8")
    elif n == "ppt/slideMasters/slideMaster1.xml":
        data = swap_hex(data.decode("utf8")).encode("utf8")
    elif n in ("ppt/slides/slide1.xml", "ppt/slides/slide17.xml"):
        data = swap_hex(edit_cover_slide(data.decode("utf8"))).encode("utf8")
    elif re.match(r"ppt/slides/slide\d+\.xml$", n):
        data = swap_hex(data.decode("utf8")).encode("utf8")
    elif n == "ppt/media/image4.png":          # navy SMU logo used only by the content layouts
        data = recolor_logo(data, INK)
    zout.writestr(item, data)
zout.close()
print("wrote", DST)
