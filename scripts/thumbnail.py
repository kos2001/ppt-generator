#!/usr/bin/env python3
"""Render a .pptx to slide images and a contact-sheet grid for visual QA.

You cannot see a deck render, so layout bugs (text cutoff, overflow, clashing
colors) slip through. This exports the deck to per-slide images and tiles them
into a single grid image you can open to eyeball every slide at once — the
visual-validation step from Anthropic's official pptx skill.

Pipeline:  .pptx --(PowerPoint COM)--> PNGs --(Pillow)--> grid

Usage:
    python thumbnail.py deck.pptx                 # -> deck.thumbs/ + deck.grid.png
    python thumbnail.py deck.pptx -o out_dir --cols 4 --dpi 120

Requirements:
    - Desktop Microsoft PowerPoint (COM-registered) + `pip install pywin32`
    - Pillow for the grid (already a skill dependency)

PowerPoint exports each slide straight to PNG (Slide.Export), so there is no
external converter to install. Because it drives the signed-in PowerPoint, it
also renders DRM/EDM decks the user can open — no separate plaintext copy.
"""
import argparse
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

MSO_FALSE, MSO_TRUE = 0, -1
PT_TO_IN = 1.0 / 72.0


def pptx_to_pngs(pptx_path, out_dir, dpi):
    """Export each slide to a PNG via PowerPoint COM. Returns sorted PNG paths.

    Drives the installed desktop PowerPoint: PowerPoint opens the deck in the
    user's authorized session and writes one PNG per slide, sized from the
    slide's real dimensions at the requested DPI.
    """
    try:
        import win32com.client
    except ImportError:
        raise SystemExit(
            "pywin32 not found. Install it to render slides via PowerPoint: "
            "`pip install pywin32` (needs desktop PowerPoint installed). "
            "Or open the deck in PowerPoint/Keynote to inspect it manually."
        )
    from pythoncom import com_error  # type: ignore

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        pres = powerpoint.Presentations.Open(
            os.path.abspath(pptx_path), ReadOnly=MSO_TRUE,
            Untitled=MSO_FALSE, WithWindow=MSO_FALSE,
        )
    except com_error as e:
        raise SystemExit(
            "PowerPoint COM could not open the deck (%s). Ensure desktop "
            "PowerPoint is installed and you can open this file." % (e,)
        )
    paths = []
    try:
        ps = pres.PageSetup
        w_px = max(1, int(round(ps.SlideWidth * PT_TO_IN * dpi)))
        h_px = max(1, int(round(ps.SlideHeight * PT_TO_IN * dpi)))
        for i in range(1, pres.Slides.Count + 1):
            # PowerPoint needs a native, absolute path (backslashes) — a
            # relative or forward-slash path makes Slide.Export fail.
            out = os.path.normpath(os.path.abspath(
                os.path.join(out_dir, "slide-%02d.png" % i)))
            pres.Slides(i).Export(out, "PNG", w_px, h_px)
            paths.append(out)
    finally:
        pres.Close()
        powerpoint.Quit()
    return paths


def make_grid(png_paths, grid_path, cols, pad=12, bg=(245, 245, 245)):
    from PIL import Image, ImageDraw
    if not png_paths:
        raise SystemExit("no slide images to tile")
    thumbs = [Image.open(p).convert("RGB") for p in png_paths]
    cw = max(t.width for t in thumbs)
    ch = max(t.height for t in thumbs)
    rows = (len(thumbs) + cols - 1) // cols
    grid = Image.new("RGB",
                     (cols * cw + pad * (cols + 1), rows * ch + pad * (rows + 1)), bg)
    draw = ImageDraw.Draw(grid)
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (cw + pad)
        y = pad + r * (ch + pad)
        grid.paste(t, (x, y))
        draw.text((x + 4, y + 2), str(i + 1), fill=(0, 0, 0))
    grid.save(grid_path)
    return grid_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render a .pptx to a thumbnail grid.")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", help="output dir (default <deck>.thumbs)")
    ap.add_argument("--cols", type=int, default=3, help="grid columns (default 3)")
    ap.add_argument("--dpi", type=int, default=100, help="raster DPI (default 100)")
    args = ap.parse_args(argv)

    base = os.path.splitext(os.path.basename(args.pptx))[0]
    out_dir = args.out or (os.path.splitext(args.pptx)[0] + ".thumbs")
    os.makedirs(out_dir, exist_ok=True)

    pngs = pptx_to_pngs(args.pptx, out_dir, args.dpi)
    grid = os.path.join(os.path.dirname(os.path.abspath(args.pptx)), base + ".grid.png")
    make_grid(pngs, grid, args.cols)
    print("Rendered %d slide image(s) in %s" % (len(pngs), out_dir))
    print("Contact sheet: %s" % grid)
    return 0


if __name__ == "__main__":
    sys.exit(main())
