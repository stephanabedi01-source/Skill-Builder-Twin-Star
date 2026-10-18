"""Render slides of a .pptx to PNG for visual QA (LibreOffice -> PDF -> pdftoppm).

Usage: python3 render_slides.py deck.pptx OUTDIR [--pages 7 9] [--dpi 110] [--sheet]

- OUTDIR should be a NEW folder per render round (e.g. out/r07a, out/r07b). Never delete old renders with rm
  globs; just use a fresh folder each time, which also keeps a before/after trail.
- --pages: first and last slide number (1-based) to rasterize; default all.
- --sheet: also write OUTDIR/sheet.png, a contact sheet of the rendered pages (quick whole-deck look).
- The deck font may not exist in the sandbox (e.g. Garamond); LibreOffice substitutes a wider serif, so text
  that *barely* overflows in the render usually fits in PowerPoint. Anything clearly colliding is real.
"""
import glob
import os
import subprocess
import sys


def main():
    a = sys.argv[1:]
    deck, out = a[0], a[1]
    dpi = a[a.index("--dpi") + 1] if "--dpi" in a else "110"
    os.makedirs(out, exist_ok=True)
    r = subprocess.run(["timeout", "240", "soffice", "--headless", "--convert-to", "pdf", "--outdir", out, deck],
                       capture_output=True, text=True)
    pdf = os.path.join(out, os.path.splitext(os.path.basename(deck))[0] + ".pdf")
    if not os.path.exists(pdf):
        sys.exit(f"PDF not produced: {r.stderr[-500:]}")
    cmd = ["pdftoppm", "-r", dpi, "-png"]
    if "--pages" in a:
        i = a.index("--pages")
        cmd += ["-f", a[i + 1], "-l", a[i + 2]]
    subprocess.run(cmd + [pdf, os.path.join(out, "s")], check=True)
    pngs = sorted(glob.glob(os.path.join(out, "s-*.png")))
    if "--sheet" in a and pngs:
        subprocess.run(["montage", *pngs, "-tile", "2x", "-geometry", "+6+6", os.path.join(out, "sheet.png")])
    print("\n".join(pngs))


if __name__ == "__main__":
    main()
