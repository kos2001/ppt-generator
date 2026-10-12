"""Font-free text measurement for fitting text into fixed boxes.

python-pptx cannot ask PowerPoint how big text will render, and the fonts named
by a template are often not installed where the deck is built. So we estimate:
each character gets an average advance width (in em) by class, calibrated
against real fonts (Arial, Helvetica Neue, Georgia, Verdana, Apple SD Gothic
Neo). Estimates lean slightly wide so text that "fits" here fits on screen.

The renderer uses this to shrink text that would overflow its box, and the
validator uses it to warn about slides that are too dense to fit at a readable
size.
"""
import re
import unicodedata

PT_PER_INCH = 72.0
EMU_PER_PT = 12700

# Average advance width per character class, in em. Measured averages
# (lowercase ~0.45-0.52, uppercase ~0.55-0.63, digits ~0.51-0.58) rounded up
# a little for safety. Hangul uses Malgun Gothic's exact advance: every one
# of its 11,172 syllables is 2048/2048 units wide. PowerPoint's "single" line
# is a flat 1.2 x font size for every font, scaled by spcPct.
_W_SPACE = 0.30   # Calibri 0.23, Arial 0.28, Malgun Gothic 0.35
_W_NARROW = 0.28   # i l j . , : ; ' ! | etc.
_W_LOWER = 0.50
_W_UPPER = 0.62
_W_DIGIT = 0.56
_W_WIDE = 1.00     # Hangul, CJK, fullwidth forms (Malgun Gothic: exactly 1 em)
_W_OTHER = 0.60

_NARROW = set("iljtfIr.,:;'!|()[]`")

# Renderers differ (PowerPoint vs Keynote vs LibreOffice, and substituted
# fonts), so wrap against slightly less than the real width.
WIDTH_SLACK = 0.95

# Relative width of each template font versus the Arial-like baseline.
FONT_SCALE = {
    "calibri": 0.92,
    "helvetica neue": 1.0,
    "helvetica": 1.0,
    "arial": 1.0,
    "georgia": 1.0,
    "verdana": 1.13,
}

_HANGUL_RE = re.compile(r"[ᄀ-ᇿ㄰-㆏가-힯]")


def has_hangul(text):
    return bool(_HANGUL_RE.search(text or ""))


def is_wide(ch):
    return unicodedata.east_asian_width(ch) in ("W", "F")


def char_em(ch):
    if ch == " ":
        return _W_SPACE
    if is_wide(ch):
        return _W_WIDE
    if ch in _NARROW:
        return _W_NARROW
    if ch.isdigit():
        return _W_DIGIT
    if ch.islower():
        return _W_LOWER
    if ch.isupper():
        return _W_UPPER
    return _W_OTHER


def text_width_pt(text, size_pt, font=None, bold=False):
    scale = FONT_SCALE.get((font or "").lower(), 1.0) * (1.06 if bold else 1.0)
    # Wide (CJK) glyphs come from the East Asian font, so the latin font's
    # width scale does not apply to them.
    total = 0.0
    for ch in text:
        w = char_em(ch)
        total += w if w == _W_WIDE else w * scale
    return total * size_pt


def wrap_line_count(text, width_pt, size_pt, font=None, bold=False):
    """Greedy word-wrap: how many lines `text` occupies in `width_pt`."""
    if width_pt <= 0:
        return 1
    lines = 0
    for hard in (text or "").split("\n"):
        lines += _wrap_one(hard, width_pt, size_pt, font, bold)
    return max(lines, 1)


def _wrap_one(text, width_pt, size_pt, font, bold):
    width_pt *= WIDTH_SLACK
    space = text_width_pt(" ", size_pt, font, bold)
    lines, cur = 1, 0.0
    for word in text.split(" "):
        w = text_width_pt(word, size_pt, font, bold)
        if cur == 0.0:
            add = w
        else:
            add = cur + space + w
        if add <= width_pt:
            cur = add
            continue
        if cur > 0.0:
            lines += 1
        # A single word wider than the line breaks across lines.
        while w > width_pt:
            lines += 1
            w -= width_pt
        cur = w
    return lines


def last_line_ratio(text, width_pt, size_pt, font=None, bold=False):
    """(line_count, fraction of the line width used by the last line) for a
    single paragraph — used to spot a title whose last line is one orphaned
    word."""
    width_pt *= WIDTH_SLACK
    space = text_width_pt(" ", size_pt, font, bold)
    lines, cur = 1, 0.0
    for word in (text or "").split(" "):
        w = text_width_pt(word, size_pt, font, bold)
        add = w if cur == 0.0 else cur + space + w
        if add <= width_pt or cur == 0.0:
            cur = add
        else:
            lines += 1
            cur = w
    return lines, (cur / width_pt if width_pt else 0.0)


def paragraphs_height_pt(paras, width_pt, line_spacing=1.1):
    """Height of a list of paragraph dicts {text, size, bold, font, indent_pt,
    space_after}. Line height ~= 1.2 * size * line_spacing (PowerPoint's
    single spacing is roughly 1.2 em)."""
    total = 0.0
    for p in paras:
        size = p["size"]
        avail = width_pt - p.get("indent_pt", 0.0)
        n = wrap_line_count(p.get("text", ""), avail, size, p.get("font"),
                            p.get("bold", False))
        total += n * size * 1.2 * line_spacing + p.get("space_after", 0.0)
    return total


def fit_scale(paras, width_pt, height_pt, line_spacing=1.1, min_scale=0.6):
    """Largest scale in [min_scale, 1] (by ~4% steps) at which `paras` fit.

    Returns (scale, fits). When even min_scale overflows, returns
    (min_scale, False) so callers can warn.
    """
    scale = 1.0
    while True:
        scaled = [dict(p, size=p["size"] * scale,
                       space_after=p.get("space_after", 0.0) * scale,
                       indent_pt=p.get("indent_pt", 0.0))
                  for p in paras]
        if paragraphs_height_pt(scaled, width_pt, line_spacing) <= height_pt:
            return scale, True
        if scale <= min_scale + 1e-9:
            return min_scale, False
        scale = max(min_scale, scale - 0.04)
