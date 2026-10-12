#!/usr/bin/env python3
"""Audit any .pptx for slide design flaws that can be detected from geometry.

Works on decks this skill generated (Mode A) and on decks the user brought
(Mode B, before and after editing). Findings are tagged with the SlideAudit
taxonomy category they belong to (Zhang et al., UIST '25 — see
references/review-taxonomy.md), so the automated pass and the visual review
speak the same language:

  Content Overflow or Truncation  text taller than its box, shapes off-slide
  Occluded Content                text boxes whose text areas overlap
  Misaligned Elements             left edges that *almost* line up
  Inappropriate Font Sizing       text too small to read
  Excessive Text Density          too many words on one slide
  Inconsistent Text Styling       too many typefaces; Hangul with no a:ea font
  Insufficient Contrast           explicit text color vs. what is behind it
  Low-Quality / Distorted Images  low effective DPI, stretched aspect ratio
  Accessibility                   pictures without alt text

The checks are estimates (text width comes from textfit.py, not a real layout
engine) and skip what they cannot know (theme colors, inherited sizes are
assumed 18pt). Treat findings as leads, then confirm on the rendered slide.

Usage:
    python audit_pptx.py deck.pptx              # human-readable report
    python audit_pptx.py deck.pptx --json       # machine-readable
    python audit_pptx.py deck.pptx --strict     # exit 1 if any error
"""
import argparse
import io
import json
import os
import re
import sys

sys.dont_write_bytecode = True  # keep the skill directory free of __pycache__

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import textfit  # noqa: E402

EMU_PER_IN = 914400
DEFAULT_PT = 18.0           # size assumed when a run inherits its size
MIN_PT = 10.0               # smallest readable body text on a projected slide
MIN_PT_FOOTER = 7.0         # footer band may go smaller, but not tiny
FOOTER_BAND = 0.10          # bottom share of the slide treated as footer
MAX_WORDS = 90              # words (Hangul: ~2 syllables per word) per slide
MAX_TYPEFACES = 3           # distinct Latin typefaces in one deck
NEAR_ALIGN = (0.03, 0.15)   # inches: edges this close but not equal = misaligned
OVERFLOW_TOL = 1.12         # estimated text height vs. box height
MIN_DPI = 96                # effective resolution of a picture on a 13.33" slide
ASPECT_TOL = 0.03           # stretched more than 3% = distorted

SEVERITY_ORDER = {"error": 0, "warn": 1, "info": 2}


def _in(emu):
    return emu / float(EMU_PER_IN)


class Finding(dict):
    def __init__(self, slide, severity, category, message, shape=None):
        super().__init__(slide=slide, severity=severity, category=category,
                         message=message, shape=shape)


# --------------------------------------------------------------------------- #
# Geometry helpers
# --------------------------------------------------------------------------- #
def _iter_shapes(shapes, ox=0, oy=0, sx=1.0, sy=1.0, depth=0):
    """Yield (shape, (left, top, width, height) in slide EMU, z-order index),
    flattening groups through their child coordinate transform."""
    for shp in shapes:
        if shp.shape_type is not None and shp.shape_type == 6:  # GROUP
            xfrm = shp._element.grpSpPr.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
            try:
                off, ext = xfrm[0], xfrm[1]
                ch_off, ch_ext = xfrm[2], xfrm[3]
                gx, gy = int(off.get("x")), int(off.get("y"))
                gw, gh = int(ext.get("cx")), int(ext.get("cy"))
                cx, cy = int(ch_off.get("x")), int(ch_off.get("y"))
                cw, ch = int(ch_ext.get("cx")) or 1, int(ch_ext.get("cy")) or 1
                nsx, nsy = sx * gw / cw, sy * gh / ch
                nox, noy = ox + sx * gx - nsx * cx, oy + sy * gy - nsy * cy
            except (TypeError, IndexError, AttributeError):
                nox, noy, nsx, nsy = ox, oy, sx, sy
            yield from _iter_shapes(shp.shapes, nox, noy, nsx, nsy, depth + 1)
            continue
        if shp.left is None or shp.width is None:
            continue
        rect = (ox + sx * shp.left, oy + sy * shp.top, sx * shp.width, sy * shp.height)
        yield shp, rect


def _intersect(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return (w, h) if w > 0 and h > 0 else (0, 0)


def _contains_point(r, x, y):
    return r[0] <= x <= r[0] + r[2] and r[1] <= y <= r[1] + r[3]


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #
def _run_pt(run, para):
    if run.font.size is not None:
        return run.font.size.pt
    try:
        dflt = para._p.pPr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
        if dflt is not None and dflt.get("sz"):
            return int(dflt.get("sz")) / 100.0
    except AttributeError:
        pass
    return DEFAULT_PT


def _font_scale(tf):
    body = tf._txBody.bodyPr
    na = body.find("{http://schemas.openxmlformats.org/drawingml/2006/main}normAutofit")
    if na is not None and na.get("fontScale"):
        return int(na.get("fontScale")) / 100000.0
    return 1.0


def _insets(tf):
    def g(v, d):
        return d if v is None else v
    return (g(tf.margin_left, 91440), g(tf.margin_right, 91440),
            g(tf.margin_top, 45720), g(tf.margin_bottom, 45720))


def _paragraph_measure(tf, scale):
    paras = []
    for p in tf.paragraphs:
        runs = list(p.runs)
        text = "".join(r.text for r in runs)
        size = max([_run_pt(r, p) for r in runs] or [DEFAULT_PT]) * scale
        font = next((r.font.name for r in runs if r.font.name), None)
        bold = any(r.font.bold for r in runs)
        indent = 0.0
        if p._p.pPr is not None and p._p.pPr.get("marL"):
            indent = int(p._p.pPr.get("marL")) / 12700.0
        spacing = 1.0
        try:
            if p.line_spacing is not None and isinstance(p.line_spacing, float):
                spacing = p.line_spacing
        except (ValueError, TypeError):
            pass
        after = p.space_after.pt if p.space_after is not None else 0.0
        paras.append({"text": text, "size": size, "font": font, "bold": bold,
                      "indent_pt": indent, "space_after": after,
                      "_spacing": spacing})
    return paras


def _text_height_pt(paras, width_pt):
    total = 0.0
    for p in paras:
        total += textfit.paragraphs_height_pt([p], width_pt, p.get("_spacing", 1.0))
    return total


def _text_rect(shape, rect, tf, est_h_emu):
    """Where the text actually sits inside the shape (anchor-aware)."""
    left, right, top, bottom = _insets(tf)
    x, y, w, h = rect
    inner_h = max(0, h - top - bottom)
    anchor = str(tf.vertical_anchor) if tf.vertical_anchor is not None else "TOP"
    th = min(est_h_emu, inner_h) if est_h_emu <= inner_h else est_h_emu
    if "MIDDLE" in anchor:
        ty = y + top + (inner_h - th) / 2
    elif "BOTTOM" in anchor:
        ty = y + h - bottom - th
    else:
        ty = y + top
    return (x + left, ty, max(0, w - left - right), th)


# --------------------------------------------------------------------------- #
# Color helpers
# --------------------------------------------------------------------------- #
def _luminance(hexstr):
    c = [int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(a, b):
    la, lb = _luminance(a), _luminance(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _solid_fill(shape):
    try:
        f = shape.fill
        if f.type == 1:  # MSO_FILL.SOLID
            return str(f.fore_color.rgb)
    except (AttributeError, TypeError, ValueError, KeyError):
        pass
    return None


def _slide_background(slide):
    for src in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        try:
            f = src.background.fill
            if f.type == 1:
                return str(f.fore_color.rgb)
        except (AttributeError, TypeError, ValueError, KeyError):
            continue
    return "FFFFFF"


def _run_color(run):
    try:
        c = run.font.color
        if c and c.type is not None and c.type == 1:  # RGB
            return str(c.rgb)
    except (AttributeError, TypeError, ValueError):
        pass
    return None


# --------------------------------------------------------------------------- #
# Audit
# --------------------------------------------------------------------------- #
def audit(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(path)
    sw, sh = prs.slide_width, prs.slide_height
    findings = []
    typefaces = {}
    hangul_without_ea = 0

    for sidx, slide in enumerate(prs.slides, start=1):
        items = list(_iter_shapes(slide.shapes))
        bg = _slide_background(slide)
        text_boxes = []   # (shape, text rect, label)
        words = 0
        lefts = []

        for z, (shp, rect) in enumerate(items):
            label = "%s" % shp.name
            x, y, w, h = rect
            # -- off-slide ------------------------------------------------
            tol = 0.05 * EMU_PER_IN
            if x < -tol or y < -tol or x + w > sw + tol or y + h > sh + tol:
                findings.append(Finding(sidx, "error", "Content Overflow or Truncation",
                                        "shape extends past the slide edge", label))

            # -- pictures ---------------------------------------------------
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                _audit_picture(shp, rect, sidx, label, findings)
                continue

            if not getattr(shp, "has_text_frame", False) or not shp.has_text_frame:
                continue
            tf = shp.text_frame
            text = tf.text.strip()
            if not text:
                continue
            # A lone glyph or two of punctuation (a giant quote mark, a bullet
            # ornament) is decoration: it may overhang its box by design.
            decorative = len(text) <= 2 and not any(c.isalnum() for c in text)

            scale = _font_scale(tf)
            paras = _paragraph_measure(tf, scale)
            l_in, r_in, t_in, b_in = _insets(tf)
            width_pt = max(1.0, (w - l_in - r_in) / 12700.0)
            box_h_pt = max(1.0, (h - t_in - b_in) / 12700.0)
            if tf.word_wrap is False:
                widest = max(textfit.text_width_pt(p["text"], p["size"], p["font"], p["bold"])
                             for p in paras)
                est_h = sum(p["size"] * 1.2 * p["_spacing"] for p in paras)
                if widest > width_pt * 1.1 and w > 0:
                    findings.append(Finding(sidx, "warn", "Content Overflow or Truncation",
                                            "unwrapped text is ~%d%% wider than its box: \"%s\""
                                            % (round(widest / width_pt * 100), text[:40]),
                                            label))
            else:
                est_h = _text_height_pt(paras, width_pt)
            est_h_emu = est_h * 12700
            if decorative:
                continue
            if est_h > box_h_pt * OVERFLOW_TOL + 6:
                below = y + t_in + est_h_emu > sh
                findings.append(Finding(
                    sidx, "error" if below else "warn", "Content Overflow or Truncation",
                    "text needs ~%d%% of its box height%s: \"%s\""
                    % (round(est_h / box_h_pt * 100),
                       " and runs off the slide" if below else "", text[:40]), label))
            trect = _text_rect(shp, rect, tf, est_h_emu)
            trect = _ink_width(trect, tf, paras, width_pt)
            text_boxes.append((shp, trect, label, text))
            # Only main-column blocks (wide, left-aligned) are compared for
            # alignment; card and diagram labels align to their own containers.
            if w > 0.3 * sw and _left_aligned(tf):
                lefts.append((rect[0] + l_in, label))

            # -- runs: size, typeface, color, hangul ------------------------
            in_footer = y >= sh * (1 - FOOTER_BAND)
            own_fill = _solid_fill(shp)
            behind = own_fill or _color_behind(items[:z], rect) or bg
            for p in tf.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    pt = _run_pt(r, p) * scale
                    if pt < (MIN_PT_FOOTER if in_footer else MIN_PT):
                        findings.append(Finding(
                            sidx, "warn", "Inappropriate Font Sizing",
                            "%.0fpt text is too small to read: \"%s\"" % (pt, r.text[:30]),
                            label))
                    if r.font.name:
                        typefaces[r.font.name] = typefaces.get(r.font.name, 0) + 1
                    if textfit.has_hangul(r.text):
                        rpr = r._r.find(
                            "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
                        ea = None if rpr is None else rpr.find(
                            "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
                        if ea is None or not ea.get("typeface"):
                            hangul_without_ea += 1
                    color = _run_color(r)
                    if color and behind:
                        ratio = contrast(color, behind)
                        large = pt >= 24 or (pt >= 18.5 and r.font.bold)
                        need = 3.0 if large else 4.5
                        if ratio < need:
                            findings.append(Finding(
                                sidx, "warn", "Insufficient Contrast",
                                "%s on %s is %.1f:1 (needs %.1f:1): \"%s\""
                                % (color, behind, ratio, need, r.text[:30]), label))
            words += _word_count(text)

        # -- occlusion: overlapping text areas ------------------------------
        for i in range(len(text_boxes)):
            for j in range(i + 1, len(text_boxes)):
                a, b = text_boxes[i][1], text_boxes[j][1]
                iw, ih = _intersect(a, b)
                small = min(a[2] * a[3], b[2] * b[3]) or 1
                if iw > 0.05 * EMU_PER_IN and ih > 0.05 * EMU_PER_IN and iw * ih > 0.1 * small:
                    findings.append(Finding(
                        sidx, "warn", "Occluded Content",
                        "text of '%s' overlaps text of '%s'"
                        % (text_boxes[i][2], text_boxes[j][2]), text_boxes[j][2]))

        # -- near-miss alignment of left edges -----------------------------
        reported = set()
        for i in range(len(lefts)):
            for j in range(i + 1, len(lefts)):
                d = abs(lefts[i][0] - lefts[j][0]) / EMU_PER_IN
                key = tuple(sorted((round(lefts[i][0] / 9144), round(lefts[j][0] / 9144))))
                if NEAR_ALIGN[0] < d < NEAR_ALIGN[1] and key not in reported:
                    reported.add(key)
                    findings.append(Finding(
                        sidx, "info", "Misaligned Elements",
                        "left edges of '%s' and '%s' differ by %.2f\" — align them"
                        % (lefts[i][1], lefts[j][1], d)))

        if words > MAX_WORDS:
            findings.append(Finding(sidx, "warn", "Excessive Text Density",
                                    "~%d words on one slide (aim for <= %d; move "
                                    "detail to notes or split)" % (words, MAX_WORDS)))

    # -- deck-level -------------------------------------------------------
    if len(typefaces) > MAX_TYPEFACES:
        top = sorted(typefaces, key=typefaces.get, reverse=True)
        findings.append(Finding(0, "warn", "Inconsistent or Distracting Text Styling",
                                "%d typefaces in one deck (%s) — use one or two"
                                % (len(typefaces), ", ".join(top[:6]))))
    if hangul_without_ea:
        theme_font = _theme_hangul_font(prs)
        if theme_font:
            findings.append(Finding(0, "info", "Inconsistent or Distracting Text Styling",
                                    "%d Korean text run(s) rely on the theme's Hangul font "
                                    "(%s); set a:ea per run to pin it" % (hangul_without_ea,
                                                                          theme_font)))
        else:
            findings.append(Finding(0, "warn", "Inconsistent or Distracting Text Styling",
                                    "%d Korean text run(s) have no East Asian font (a:ea) "
                                    "and the theme names none: Hangul falls back to "
                                    "whatever the viewer's PC picks" % hangul_without_ea))
    findings.sort(key=lambda f: (f["slide"], SEVERITY_ORDER[f["severity"]]))
    return findings


def _left_aligned(tf):
    for p in tf.paragraphs:
        if p.alignment is not None and "LEFT" not in str(p.alignment):
            return False
    return True


def _ink_width(trect, tf, paras, width_pt):
    """Shrink a one-line text rect to the width the text actually covers,
    positioned by its alignment — a wide footer box with short text must not
    count as overlapping the page number beside it."""
    if len(paras) != 1:
        return trect
    p = paras[0]
    ink = textfit.text_width_pt(p["text"], p["size"], p["font"], p["bold"])
    if ink >= width_pt:
        return trect
    ink_emu = ink * 12700
    x, y, w, h = trect
    align = str(tf.paragraphs[0].alignment) if tf.paragraphs[0].alignment is not None else "LEFT"
    if "CENTER" in align:
        x = x + (w - ink_emu) / 2
    elif "RIGHT" in align:
        x = x + w - ink_emu
    return (x, y, ink_emu, h)


def _theme_hangul_font(prs):
    """The Hangul typeface the deck's theme provides (minor/body font), if any."""
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    try:
        root = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    except Exception:  # noqa: BLE001
        return None
    minor = root.find(".//" + ns + "minorFont")
    if minor is None:
        return None
    ea = minor.find(ns + "ea")
    if ea is not None and ea.get("typeface"):
        return ea.get("typeface")
    for f in minor.findall(ns + "font"):
        if f.get("script") == "Hang" and f.get("typeface"):
            return f.get("typeface")
    return None


def _word_count(text):
    latin = len(re.findall(r"[A-Za-z0-9]+", text))
    hangul = len(re.findall(r"[가-힯]", text))
    return latin + (hangul + 1) // 2


def _color_behind(below, rect):
    """Solid fill of the topmost shape under the center of `rect`."""
    cx, cy = rect[0] + rect[2] / 2, rect[1] + rect[3] / 2
    for shp, r in reversed(below):
        if _contains_point(r, cx, cy):
            fill = _solid_fill(shp)
            if fill:
                return fill
    return None


def _audit_picture(shp, rect, sidx, label, findings):
    descr = shp._element.nvPicPr.cNvPr.get("descr") if hasattr(shp._element, "nvPicPr") else None
    # python-pptx and many tools store the file name as alt text; that tells a
    # screen-reader user nothing, so treat it as missing.
    if not descr or re.search(r"\.(png|jpe?g|gif|bmp|tiff?|svg|emf|wmf)$", descr.strip(), re.I):
        findings.append(Finding(sidx, "info", "Accessibility",
                                "picture has no meaningful alt text%s"
                                % (" (only the file name)" if descr else ""), label))
    try:
        from PIL import Image
        with Image.open(io.BytesIO(shp.image.blob)) as im:
            pw, ph = im.size
    except Exception:  # noqa: BLE001 - unreadable/vector image: skip
        return
    vis_w = pw * (1 - (shp.crop_left or 0) - (shp.crop_right or 0))
    vis_h = ph * (1 - (shp.crop_top or 0) - (shp.crop_bottom or 0))
    w_in, h_in = _in(rect[2]), _in(rect[3])
    if w_in <= 0 or h_in <= 0 or vis_w <= 0 or vis_h <= 0:
        return
    dpi = min(vis_w / w_in, vis_h / h_in)
    if dpi < MIN_DPI:
        findings.append(Finding(sidx, "warn", "Low-Quality or Distorted Images",
                                "picture is ~%d dpi at its displayed size (needs ~%d+)"
                                % (dpi, MIN_DPI), label))
    stretch = (vis_w / vis_h) / (w_in / h_in)
    if abs(stretch - 1) > ASPECT_TOL:
        findings.append(Finding(sidx, "warn", "Improper Image Scaling or Size",
                                "picture is stretched %d%% %s"
                                % (round(abs(stretch - 1) * 100),
                                   "horizontally" if stretch < 1 else "vertically"), label))


def to_text(findings, path):
    if not findings:
        return "OK: %s — no geometry-detectable design flaws found" % path
    lines = []
    for f in findings:
        where = "deck" if f["slide"] == 0 else "slide %d" % f["slide"]
        shape = " [%s]" % f["shape"] if f.get("shape") else ""
        lines.append("%s: %s: %s%s: %s" % (f["severity"].upper(), where,
                                           f["category"], shape, f["message"]))
    counts = {s: sum(1 for f in findings if f["severity"] == s)
              for s in ("error", "warn", "info")}
    lines.append("\n%d error(s), %d warning(s), %d note(s)"
                 % (counts["error"], counts["warn"], counts["info"]))
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Audit a .pptx for design flaws (SlideAudit taxonomy)")
    ap.add_argument("pptx")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    ap.add_argument("--strict", action="store_true", help="exit 1 if any error is found")
    args = ap.parse_args(argv)
    findings = audit(args.pptx)
    print(json.dumps(findings, ensure_ascii=False, indent=1) if args.json
          else to_text(findings, args.pptx))
    if args.strict and any(f["severity"] == "error" for f in findings):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
