#!/usr/bin/env python3
"""Render a .pptx to slide images so the deck can be checked visually.

Converts the deck to PDF with LibreOffice (headless), rasterizes each page,
and assembles a numbered contact sheet (one PNG with every slide in a grid).
Look at the sheet (or individual slides) to catch what the fit report can't:
overlaps, awkward wrapping, empty-looking slides, low contrast.

Usage:
    python render_preview.py deck.pptx                # -> deck_preview/
    python render_preview.py deck.pptx -o out_dir --dpi 80 --cols 3

Requires LibreOffice (`soffice`) plus either PyMuPDF (`pip install pymupdf`)
or Poppler's `pdftoppm` for rasterizing. LibreOffice substitutes fonts that
are not installed, so wrapping can differ slightly from PowerPoint — judge
layout and density, not exact line breaks. Korean/CJK fonts found on the
machine (Malgun Gothic, Apple SD Gothic Neo, Noto Sans KR/CJK, Nanum,
Pretendard, ...) are handed to LibreOffice automatically; if Korean still
renders blank, install one (e.g. `brew install --cask font-noto-sans-kr`) or
pass `--font-dir`.
"""
import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys

sys.dont_write_bytecode = True  # keep the skill directory free of __pycache__
import tempfile

_SOFFICE_CANDIDATES = [
    "soffice", "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
]


def find_soffice():
    for c in _SOFFICE_CANDIDATES:
        path = shutil.which(c) or (c if os.path.isfile(c) else None)
        if path:
            return path
    return None


# Fonts LibreOffice may not pick up on its own (notably on macOS, where it can
# fall back to a font with no Hangul and render Korean as blanks). Matching
# files are linked into the throwaway profile's user/fonts dir, which
# LibreOffice always loads.
_FONT_DIRS = [
    "~/Library/Fonts", "/Library/Fonts", "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental", "~/.fonts", "~/.local/share/fonts",
    "/usr/share/fonts", "/usr/local/share/fonts", r"C:\Windows\Fonts",
]
_FONT_PATTERN = re.compile(
    r"(malgun|nanum|notosans(kr|cjk)|noto ?sans ?(kr|cjk)|sourcehansans|"
    r"pretendard|applesdgothic|spoqa|gulim|batang|dotum|carlito|calibri|"
    r"georgia|verdana|arial|helvetica)", re.I)


def _seed_fonts(profile, extra_dirs):
    dest = os.path.join(profile, "user", "fonts")
    os.makedirs(dest, exist_ok=True)
    n = 0
    for d in list(extra_dirs) + _FONT_DIRS:
        d = os.path.expanduser(d)
        if not os.path.isdir(d):
            continue
        for root, _dirs, files in os.walk(d):
            for f in files:
                if not f.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                if not (d in extra_dirs or _FONT_PATTERN.search(f.replace(" ", ""))):
                    continue
                link = os.path.join(dest, f)
                if not os.path.exists(link):
                    try:
                        os.symlink(os.path.join(root, f), link)
                    except OSError:
                        shutil.copy(os.path.join(root, f), link)
                    n += 1
    return n


def to_pdf(pptx, out_dir, font_dirs=()):
    soffice = find_soffice()
    if not soffice:
        sys.exit("ERROR: LibreOffice not found. Install it (macOS: "
                 "`brew install --cask libreoffice`; Debian/Ubuntu: "
                 "`apt install libreoffice-impress`) and retry.")
    # A private profile dir avoids clashing with a running LibreOffice and
    # with first-run dialogs.
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    _seed_fonts(profile, font_dirs)
    try:
        subprocess.run(
            [soffice, "-env:UserInstallation=file://" + profile, "--headless",
             "--convert-to", "pdf", "--outdir", out_dir, pptx],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=300)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        sys.exit("ERROR: LibreOffice did not produce %s" % pdf)
    return pdf


def rasterize(pdf, out_dir, dpi):
    for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(old)
    try:
        import pymupdf
        doc = pymupdf.open(pdf)
        for i, page in enumerate(doc):
            page.get_pixmap(dpi=dpi).save(os.path.join(out_dir, "slide-%02d.png" % (i + 1)))
    except ImportError:
        if not shutil.which("pdftoppm"):
            sys.exit("ERROR: need PyMuPDF (`pip install pymupdf`) or pdftoppm (Poppler).")
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf,
                        os.path.join(out_dir, "slide")], check=True)
        # pdftoppm numbers slide-1.png or slide-01.png depending on page count;
        # normalize to two digits.
        for f in glob.glob(os.path.join(out_dir, "slide-*.png")):
            n = int(os.path.basename(f)[6:-4])
            os.rename(f, os.path.join(out_dir, "slide-%02d.png" % n))
    return sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))


# Layout check on the rendered pixels. The fit report knows what the renderer
# *meant* to draw; this measures what actually rendered. A slide whose body
# leaves a large empty band at the bottom reads as unfinished (a top-heavy
# layout is the most common visual flaw in generated decks).
EMPTY_BOTTOM_LIMIT = 0.35   # share of slide height
FOOTER_BAND = 0.10          # footer / page number zone, ignored
INK_DELTA = 40              # per-channel difference from the background


def layout_metrics(images):
    from PIL import Image
    out = []
    for i, f in enumerate(images):
        im = Image.open(f).convert("RGB")
        w, h = im.size
        small = im.resize((max(1, w // 4), max(1, h // 4)))
        sw, sh = small.size
        px = small.load()
        # Background = the most common color over the whole slide (a header
        # bar or full-bleed band can dominate the border, never the area).
        bg = max(small.getcolors(sw * sh), key=lambda c: c[0])[1]
        body_bottom = int(sh * (1 - FOOTER_BAND))
        rows_with_ink = [y for y in range(body_bottom)
                         if any(max(abs(a - b) for a, b in zip(px[x, y], bg)) > INK_DELTA
                                for x in range(0, sw, 2))]
        ink_cells = sum(1 for y in range(body_bottom) for x in range(0, sw, 2)
                        if max(abs(a - b) for a, b in zip(px[x, y], bg)) > INK_DELTA)
        last = rows_with_ink[-1] if rows_with_ink else 0
        empty_bottom = (body_bottom - last) / float(sh)
        out.append({"slide": i + 1,
                    "empty_bottom": round(empty_bottom, 3),
                    "ink_ratio": round(ink_cells / float(max(1, body_bottom * (sw // 2))), 3)})
    return out


def contact_sheet(images, out_path, cols):
    from PIL import Image, ImageDraw
    ims = [Image.open(f).convert("RGB") for f in images]
    w, h = ims[0].size
    gap, label_h = 12, 22
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w + (cols + 1) * gap,
                              rows * (h + label_h) + (rows + 1) * gap), "#9AA0A6")
    draw = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        x = gap + (i % cols) * (w + gap)
        y = gap + (i // cols) * (h + label_h + gap)
        draw.text((x, y + 4), "slide %d" % (i + 1), fill="#FFFFFF")
        sheet.paste(im, (x, y + label_h))
    sheet.save(out_path)
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render a .pptx to slide PNGs + contact sheet")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", help="Output dir (default: <deck>_preview/)")
    ap.add_argument("--dpi", type=int, default=60, help="Raster DPI (default 60)")
    ap.add_argument("--cols", type=int, default=3, help="Contact sheet columns")
    ap.add_argument("--font-dir", action="append", default=[],
                    help="Extra directory of fonts to make available (repeatable)")
    args = ap.parse_args(argv)

    pptx = os.path.abspath(args.pptx)
    out = os.path.abspath(args.out or os.path.splitext(pptx)[0] + "_preview")
    os.makedirs(out, exist_ok=True)
    pdf = to_pdf(pptx, out, args.font_dir)
    images = rasterize(pdf, out, args.dpi)
    if not images:
        sys.exit("ERROR: no pages rendered")
    sheet = contact_sheet(images, os.path.join(out, "contact_sheet.png"), args.cols)
    metrics = layout_metrics(images)
    with open(os.path.join(out, "layout_metrics.json"), "w") as f:
        json.dump(metrics, f, indent=1)
    print("Rendered %d slides -> %s" % (len(images), out))
    print("Contact sheet: %s" % sheet)
    for m in metrics:
        if m["empty_bottom"] > EMPTY_BOTTOM_LIMIT:
            print("LAYOUT: slide %d: bottom %d%% of the slide is empty -> grow the "
                  "content, center it, or pick a layout that fills the space "
                  "(intentional on title/section/quote slides)"
                  % (m["slide"], round(m["empty_bottom"] * 100)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
