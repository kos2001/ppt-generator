#!/usr/bin/env python3
"""Render a .pptx to slide images and a contact-sheet grid for visual QA.

You cannot see a deck render, so layout bugs (text cutoff, overflow, clashing
colors) slip through. This converts the deck to per-slide images and tiles them
into a single grid image you can open to eyeball every slide at once — the
visual-validation step from Anthropic's official pptx skill.

Pipeline:  .pptx --(LibreOffice)--> PDF --(pdftoppm | PyMuPDF)--> PNGs --(Pillow)--> grid

Usage:
    python thumbnail.py deck.pptx                 # -> deck.thumbs/ + deck.grid.png
    python thumbnail.py deck.pptx -o out_dir --cols 4 --dpi 120

Requirements:
    - LibreOffice (`soffice`) for the .pptx -> PDF step
    - pdftoppm (poppler) OR PyMuPDF (`pip install pymupdf`) for PDF -> PNG
    - Pillow for the grid (already a skill dependency)
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Common LibreOffice locations beyond PATH (Windows/macOS/Linux).
SOFFICE_CANDIDATES = [
    "soffice", "libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice", "/usr/bin/libreoffice",
]


def find_soffice():
    for cand in SOFFICE_CANDIDATES:
        if os.path.sep in cand or (os.altsep and os.altsep in cand):
            if os.path.exists(cand):
                return cand
        elif shutil.which(cand):
            return shutil.which(cand)
    return None


def pptx_to_pdf(pptx_path, out_dir):
    soffice = find_soffice()
    if not soffice:
        raise SystemExit(
            "LibreOffice not found. Install it to convert .pptx -> PDF "
            "(https://www.libreoffice.org/download), or open the deck in "
            "PowerPoint/Keynote to inspect it manually."
        )
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf):
        raise SystemExit("LibreOffice did not produce a PDF for %s" % pptx_path)
    return pdf


def pdf_to_pngs(pdf_path, out_dir, dpi):
    """Rasterize a PDF to one PNG per page. Returns sorted list of PNG paths."""
    prefix = os.path.join(out_dir, "slide")
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), pdf_path, prefix],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        return sorted(glob.glob(prefix + "*.png"))
    # Fallback: PyMuPDF.
    try:
        import fitz
    except ImportError:
        raise SystemExit(
            "Need pdftoppm (poppler) or PyMuPDF to rasterize the PDF. "
            "Install one: `pip install pymupdf`."
        )
    doc = fitz.open(pdf_path)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    paths = []
    for i, page in enumerate(doc, start=1):
        out = "%s-%02d.png" % (prefix, i)
        page.get_pixmap(matrix=mat).save(out)
        paths.append(out)
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

    pdf = pptx_to_pdf(args.pptx, out_dir)
    pngs = pdf_to_pngs(pdf, out_dir, args.dpi)
    grid = os.path.join(os.path.dirname(os.path.abspath(args.pptx)), base + ".grid.png")
    make_grid(pngs, grid, args.cols)
    print("Rendered %d slide image(s) in %s" % (len(pngs), out_dir))
    print("Contact sheet: %s" % grid)
    return 0


if __name__ == "__main__":
    sys.exit(main())
