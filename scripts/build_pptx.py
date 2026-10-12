#!/usr/bin/env python3
"""Render a presentation spec (JSON) into an editable .pptx file.

The spec describes WHAT goes on each slide (content + a layout name); the chosen
template describes HOW it looks (palette + fonts). This separation is the whole
point: the same content can be re-rendered in any template, and a deck is built
deterministically from a single JSON file rather than hand-placed shape by shape.

Usage:
    python build_pptx.py spec.json -o out.pptx
    python build_pptx.py spec.json            # writes <title>.pptx next to spec
    echo '{...}' | python build_pptx.py - -o out.pptx   # spec from stdin

The spec schema and layout catalog are documented in references/spec-schema.md
and references/layouts.md. validate_spec.py checks a spec before rendering.

After saving, the renderer prints a fit report: any text box it had to shrink
to fit, and any that still overflows at the minimum size. Treat OVERFLOW lines
as bugs in the spec (split the slide or cut words).
"""
import argparse
import copy
import sys

sys.dont_write_bytecode = True  # keep the skill directory free of __pycache__

import json
import math
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                             XL_LABEL_POSITION, XL_TICK_LABEL_POSITION)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from templates import get_theme, DEFAULT_THEME, THEMES, DEFAULT_EA_FONT  # noqa: E402
from flowchart_layout import layer_graph, normalize_edges  # noqa: E402
import textfit  # noqa: E402
import icons  # noqa: E402

# 16:9 widescreen canvas (13.333" x 7.5"). Most modern decks use this.
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.7)
BOTTOM = Inches(0.75)          # keep body content above the footer band

# Per-build state: the East Asian font for the deck, the slide being drawn, and
# the fit report collected while drawing. Reset by build().
_CTX = {"ea_font": DEFAULT_EA_FONT, "slide": 0, "fit": [], "bounds": [],
        "title_style": "topic"}

TITLE_STYLES = ("topic", "assertion")


# --------------------------------------------------------------------------- #
# Low-level drawing helpers
# --------------------------------------------------------------------------- #
def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _pt(emu):
    return float(emu) / textfit.EMU_PER_PT


def _blank_slide(prs):
    # slide_layouts[6] is the built-in blank layout: no placeholders, full
    # control. We draw every element ourselves so templates look identical
    # regardless of the underlying PowerPoint theme.
    return prs.slides.add_slide(prs.slide_layouts[6])


def fill_background(slide, theme):
    if theme["background"].upper() == "FFFFFF":
        return  # white is the default; skip the extra shape
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(theme["background"])
    shape.line.fill.background()
    shape.shadow.inherit = False
    # Send to back so all later content draws on top.
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)
    return shape


def add_rect(slide, left, top, width, height, fill_hex=None, line_hex=None,
             shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, int(left), int(top), int(width), int(height))
    if fill_hex is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _rgb(fill_hex)
    if line_hex is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = _rgb(line_hex)
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp


def _set_ea_font(font_obj_rpr, text):
    """Give a run (or default run props) an explicit East Asian typeface.

    python-pptx only writes <a:latin>; Hangul/CJK glyphs come from <a:ea>, and
    without it PowerPoint picks a fallback that depends on the viewer's OS
    language. Tagging Korean runs ko-KR also lets PowerPoint apply Korean
    proofing and line-breaking rules.
    """
    rPr = font_obj_rpr
    for tag in ("a:ea", "a:cs"):
        old = rPr.find(qn(tag))
        if old is not None:
            rPr.remove(old)
    ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", _CTX["ea_font"])
    # Schema order: ... latin, ea, cs, sym, hlinkClick ... — move ea right
    # after latin when present.
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addnext(ea)
    if textfit.has_hangul(text):
        rPr.set("lang", "ko-KR")
        rPr.set("altLang", "en-US")


def style_run(run, *, font, size, color, bold=False, italic=False):
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    # Hangul has no true italic: PowerPoint slants only the Latin letters and
    # digits in a mixed line, which looks broken. Keep Korean runs upright.
    f.italic = italic and not textfit.has_hangul(run.text)
    f.color.rgb = _rgb(color)
    _set_ea_font(run._r.get_or_add_rPr(), run.text)


def _korean_wrap(p, text):
    # Break Korean lines between words (어절), not mid-word. ECMA-376 eaLnBrk
    # defaults to true ("East Asian words may break anywhere"), which splits
    # words like 로드|맵 across lines.
    # The run-level lang="ko-KR" is what makes PowerPoint break Korean at
    # spaces; latinLnBrk="1" would re-enable breaks anywhere, so pin it off.
    if textfit.has_hangul(text):
        pPr = p._p.get_or_add_pPr()
        pPr.set("eaLnBrk", "0")
        pPr.set("latinLnBrk", "0")


# Bullet geometry per level: (left margin, hanging indent) in inches.
_BULLET_GEOM = [(0.34, 0.30), (0.74, 0.26), (1.10, 0.24)]
_BULLET_CHARS = ["•", "–", "◦"]


# Numbered lists need room for "10." — a wider hanging indent at level 0.
_NUMBER_GEOM = (0.48, 0.44)


def _set_bullet(p, char, color_hex, level, number=None):
    """Real PowerPoint bullet (a:buChar) with a hanging indent, so wrapped
    lines align under the text instead of under the glyph, and the bullet
    stays a bullet when the user edits the slide.

    With `number`, the paragraph is an auto-numbered item (a:buAutoNum,
    "1." style) starting at that number. Every item carries its own startAt,
    so numbering survives sub-bullets in between and continuation slides,
    in PowerPoint, Keynote and LibreOffice alike."""
    marl, hang = _NUMBER_GEOM if number is not None else _BULLET_GEOM[min(level, 2)]
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(marl))))
    pPr.set("indent", str(-int(Inches(hang))))
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone",
                "a:buAutoNum"):
        old = pPr.find(qn(tag))
        if old is not None:
            pPr.remove(old)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", color_hex)
    etree.SubElement(pPr, qn("a:buSzPct")).set("val", "100000")
    if number is not None:
        num = etree.SubElement(pPr, qn("a:buAutoNum"))
        num.set("type", "arabicPeriod")
        num.set("startAt", str(number))
        return
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)


def _check_orphan(text, width, size, font):
    """Flag a wrapped title whose last line is a lone short word."""
    lines, ratio = textfit.last_line_ratio(text, _pt(width), size, font, bold=True)
    if lines >= 2 and ratio < 0.2:
        _CTX["fit"].append((_CTX["slide"], "orphan", 1.0, True, text))


SPARSE_FILL = 0.35   # bullets filling less than this share of the body


def _check_sparse(paras, theme, width, height):
    """Flag a bullets slide whose text fills only a sliver of the body area —
    the slide reads as empty; another layout would carry it better."""
    measure = [dict(p, font=theme["body_font"],
                    indent_pt=Inches(_BULLET_GEOM[p.get("level", 0)][0]) / 12700)
               for p in paras]
    need = textfit.paragraphs_height_pt(measure, _pt(width), 1.1)
    if paras and need / _pt(height) < SPARSE_FILL:
        _CTX["fit"].append((_CTX["slide"], "sparse", need / _pt(height), True,
                            paras[0].get("text", "")))


def _fit_record(kind, scale, fits, text):
    if scale < 0.999 or not fits:
        snippet = re.sub(r"\s+", " ", text)[:48]
        _CTX["fit"].append((_CTX["slide"], kind, scale, fits, snippet))


def add_text(slide, text, left, top, width, height, *, font, size,
             color, bold=False, italic=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.05, wrap=True,
             fit=True, min_scale=0.7):
    """Add a text box. `text` may be a string or a list of paragraph dicts.

    A paragraph dict supports: {text, level, bold, italic, size, color,
    space_after, bullet, bullet_color}.

    With `fit` (the default) the text is measured and, if it would overflow
    the box, every size is scaled down uniformly (not below `min_scale`).
    Text that still does not fit is reported in the fit report.
    """
    box = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    paragraphs = text if isinstance(text, list) else [{"text": text}]
    paragraphs = [{"text": p} if isinstance(p, str) else p for p in paragraphs]

    scale = 1.0
    if fit and wrap and any(p.get("text") for p in paragraphs):
        measure = [{
            "text": p.get("text", ""),
            "size": p.get("size", size),
            "bold": p.get("bold", bold),
            "font": font,
            "indent_pt": (Inches(_NUMBER_GEOM[0] if p.get("number") is not None
                                 else _BULLET_GEOM[min(p.get("level", 0), 2)][0]) / 12700
                          if p.get("bullet") else 0.0),
            "space_after": p.get("space_after") or 0.0,
        } for p in paragraphs]
        scale, fits = textfit.fit_scale(measure, _pt(width), _pt(height),
                                        line_spacing, min_scale)
        _fit_record("text", scale, fits,
                    " / ".join(p.get("text", "") for p in paragraphs))

    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.level = min(para.get("level", 0), 2)
        if para.get("space_after") is not None:
            p.space_after = Pt(para["space_after"] * scale)
        run = p.add_run()
        run.text = para.get("text", "")
        style_run(run, font=font, size=round(para.get("size", size) * scale, 1),
                  color=para.get("color", color), bold=para.get("bold", bold),
                  italic=para.get("italic", italic))
        if para.get("bullet"):
            _set_bullet(p, para["bullet"], para.get("bullet_color", color),
                        p.level, para.get("number"))
        _korean_wrap(p, run.text)
    return box


def _title_lines_and_size(text, width, size, font):
    """Fit a title: shrink to one line down to 80% size; otherwise wrap to two
    lines (down to 70%). Returns (size_pt, lines).

    Assertion titles are full-sentence takeaways that are expected to run to
    two lines, so they wrap at full size instead of shrinking, which keeps
    every headline in the deck the same size. The one exception: if wrapping
    would strand a single short word on line two, shrinking up to 10% to stay
    on one line reads better."""
    w = _pt(width)
    one_line = (size, size * 0.9, size * 0.8)
    if _CTX["title_style"] == "assertion":
        lines, ratio = textfit.last_line_ratio(text, w, size, font, bold=True)
        one_line = (size, size * 0.9) if (lines == 2 and ratio < 0.2) else (size,)
    for s in one_line:
        if textfit.wrap_line_count(text, w, s, font, bold=True) == 1:
            return s, 1
    for s in (size, size * 0.9, size * 0.8, size * 0.7):
        if textfit.wrap_line_count(text, w, s, font, bold=True) <= 2:
            if s < size * 0.999:
                _fit_record("title", s / size, True, text)
            return s, 2
    _fit_record("title", 0.7, False, text)
    return size * 0.7, 3


def _slide_title(slide, theme, title, *, eyebrow=None, subtitle=None):
    """Content-slide header. Returns the y (EMU) where the body may start, so
    layouts never collide with a long two-line title or a subtitle.

    The visual treatment depends on the theme's `header_style`:

    - "accent" (default): a small accent block to the left of a colored title
      on the slide background. Light, editorial.
    - "bar": a full-width header band filled with the primary color and a title
      in `on_primary` (e.g. a black header bar with white text). Formal,
      report/document style. No divider line is drawn below it.
    """
    if theme.get("header_style") == "bar":
        y = _header_bar(slide, theme, title, eyebrow=eyebrow)
    else:
        y = _header_accent(slide, theme, title, eyebrow=eyebrow)
    if subtitle:
        # Intentionally larger than body text so it reads as a lead-in, not a
        # caption. Relative to the theme's body size, overridable via subtitle_pt.
        sub_pt = theme.get("subtitle_pt",
                           min(theme["body_pt"] + 4, theme["heading_pt"] - 6))
        sub_w = SLIDE_W - 2 * MARGIN
        n = min(textfit.wrap_line_count(subtitle, _pt(sub_w), sub_pt,
                                        theme["body_font"]), 2)
        sub_h = Pt(sub_pt * 1.25 * n)
        add_text(slide, subtitle, MARGIN, y + Inches(0.08), sub_w, sub_h,
                 font=theme["body_font"], size=sub_pt,
                 color=theme["secondary"], italic=True, line_spacing=1.05)
        y = y + Inches(0.08) + sub_h
    return Emu(int(y + Inches(0.3)))


def _header_accent(slide, theme, title, *, eyebrow=None):
    top = Inches(0.5)
    if eyebrow:
        add_text(slide, eyebrow.upper(), MARGIN + Inches(0.32), Inches(0.42),
                 SLIDE_W - 2 * MARGIN, Inches(0.3),
                 font=theme["body_font"], size=12, color=theme["secondary"],
                 bold=True)
        top = Inches(0.72)
    width = SLIDE_W - 2 * MARGIN - Inches(0.32)
    size, lines = _title_lines_and_size(title, width, theme["heading_pt"],
                                        theme["heading_font"])
    line_h = Pt(size * 1.22)
    height = max(Inches(0.62), line_h * lines)
    add_rect(slide, MARGIN, top + Inches(0.12), Inches(0.16),
             min(Inches(0.55), height - Inches(0.05)), fill_hex=theme["accent"])
    add_text(slide, title, MARGIN + Inches(0.32), top, width, height,
             font=theme["heading_font"], size=size,
             color=theme["primary"], bold=True, line_spacing=1.0, fit=False)
    _check_orphan(title, width, size, theme["heading_font"])
    return top + height


def _header_bar(slide, theme, title, *, eyebrow=None):
    # Reserve room at the right for the top-right classification marker (drawn
    # globally per slide by _classification_marker) so a long title never
    # collides with it.
    label_w = Inches(3.0) if theme.get("header_label") else 0
    title_w = SLIDE_W - 2 * MARGIN - label_w
    size, lines = _title_lines_and_size(title, title_w, theme["heading_pt"],
                                        theme["heading_font"])
    title_h = Pt(size * 1.22) * lines
    eyebrow_h = Inches(0.3) if eyebrow else 0
    bar_h = max(Inches(1.0), title_h + eyebrow_h + Inches(0.3))
    add_rect(slide, 0, 0, SLIDE_W, bar_h, fill_hex=theme["primary"])
    # Optional thin rule directly under the bar (a brand accent line). Themes
    # that don't set `header_rule` (e.g. report) get no divider.
    rule = theme.get("header_rule")
    if rule:
        add_rect(slide, 0, bar_h, SLIDE_W, Inches(0.05), fill_hex=rule)
    # Keep the eyebrow + title stacked WITHIN the bar — on a dark bar a title
    # that spills below it would be white text on the white body (invisible).
    if eyebrow:
        add_text(slide, eyebrow.upper(), MARGIN, Inches(0.12),
                 title_w, Inches(0.28),
                 font=theme["body_font"], size=13,
                 color=theme["on_primary"], bold=True)
    add_text(slide, title, MARGIN, eyebrow_h, title_w, bar_h - eyebrow_h,
             font=theme["heading_font"], size=size,
             color=theme["on_primary"], bold=True, line_spacing=1.0,
             anchor=MSO_ANCHOR.MIDDLE, fit=False)
    _check_orphan(title, title_w, size, theme["heading_font"])
    return bar_h + Inches(0.1)


# Layouts whose top-right corner sits on a LIGHT background, so the marker must
# use a dark color to stay legible. Every other layout has a dark area there
# (header bar or full-bleed primary fill) and uses the light on_primary color.
_LIGHT_TOP_LAYOUTS = {"title", "quote", "image"}
_BLEED_LAYOUTS = {"section", "closing"}


def _classification_marker(slide, theme, layout_name, bg_hex=None):
    """Draw the classification label (e.g. "Confidential") in the top-right
    corner of the slide, on every slide, in a color that contrasts with that
    slide's background. Enabled by the theme's `header_label`."""
    label = theme.get("header_label")
    if not label:
        return
    if bg_hex:
        # An explicit per-slide background (e.g. a dark cover) overrides the
        # layout-based guess: pick light/dark text by the background's luminance.
        r, g, b = int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16)
        lum = 0.299 * r + 0.587 * g + 0.114 * b
        color = theme["on_primary"] if lum < 128 else theme["primary"]
    elif layout_name in _BLEED_LAYOUTS:
        color = theme.get("on_bleed", theme["on_primary"])
    else:
        light_top = layout_name in _LIGHT_TOP_LAYOUTS
        # In bar-style themes the quote slide now carries the dark header bar,
        # so its top is dark — the marker must switch to the light color.
        if light_top and layout_name == "quote" and theme.get("header_style") == "bar":
            light_top = False
        color = theme["primary"] if light_top else theme["on_primary"]
    add_text(slide, label, SLIDE_W - MARGIN - Inches(3.0), Inches(0.2),
             Inches(3.0), Inches(0.6), font=theme["body_font"], size=18,
             color=color, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP,
             fit=False)


def _footer(slide, theme, slide_no=None, total=None, footer_text=None):
    parts = []
    if footer_text:
        parts.append(footer_text)
    txt = "   ·   ".join(parts) if parts else ""
    if txt:
        add_text(slide, txt, MARGIN, SLIDE_H - Inches(0.5),
                 SLIDE_W - 2 * MARGIN - Inches(1.0), Inches(0.3),
                 font=theme["body_font"], size=10, color=theme["text_muted"])
    if slide_no is not None:
        # Bottom-right page number, forced to "current / total" (e.g. "3 / 15")
        # when the total is known, so the reader always sees their position.
        num = "%d / %d" % (slide_no, total) if total else str(slide_no)
        add_text(slide, num, SLIDE_W - MARGIN - Inches(1.6),
                 SLIDE_H - Inches(0.62), Inches(1.6), Inches(0.42),
                 font=theme["body_font"], size=16, color=theme["text_muted"],
                 align=PP_ALIGN.RIGHT, fit=False)


def _bullet_paragraphs(theme, items, numbered=False, start=1):
    """Turn a list of bullet items into paragraph dicts with real bullets.

    Items may be plain strings or {text, level, bold}. Levels indent, use a
    lighter glyph, and step the size down. With `numbered`, top-level items
    are numbered from `start` (sub-items keep dash bullets).
    """
    paras = []
    n = start
    for item in items:
        if isinstance(item, str):
            item = {"text": item}
        level = min(item.get("level", 0), 2)
        paras.append({
            "text": item.get("text", ""),
            "level": level,
            "bold": item.get("bold", False),
            "size": item.get("size", theme["body_pt"] - 2 * level),
            "color": item.get("color", theme["text"] if level == 0 else theme["text_muted"]),
            "space_after": item.get("space_after", 10 if level == 0 else 6),
            "bullet": _BULLET_CHARS[level],
            "bullet_color": theme["accent"] if level == 0 else theme["text_muted"],
        })
        if numbered and level == 0:
            paras[-1]["number"] = n
            n += 1
    return paras


def _body_height(top):
    return SLIDE_H - top - BOTTOM


# --------------------------------------------------------------------------- #
# Slide layouts — one function per layout name in the spec
# --------------------------------------------------------------------------- #
def layout_title(slide, theme, s):
    # Optional per-slide background override (e.g. a dark cover): fill the whole
    # slide with the given hex and switch the text to light colors so it stays
    # legible. Without it, the title slide renders on the theme background.
    bg = s.get("background")
    if bg:
        add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_hex=bg)
        title_color = s.get("color", "FFFFFF")
        eyebrow_color = sub_color = meta_color = title_color
    else:
        fill_background(slide, theme)
        title_color = theme["primary"]
        eyebrow_color = theme["secondary"]
        sub_color = meta_color = theme["text_muted"]
    # Big accent block on the left edge for visual weight.
    add_rect(slide, 0, 0, Inches(0.4), SLIDE_H, fill_hex=theme["accent"])
    add_text(slide, s.get("eyebrow", "").upper(), Inches(1.1), Inches(2.1),
             SLIDE_W - Inches(2.0), Inches(0.4),
             font=theme["body_font"], size=14, color=eyebrow_color, bold=True)
    add_text(slide, s.get("title", ""), Inches(1.1), Inches(2.5),
             SLIDE_W - Inches(2.0), Inches(2.0),
             font=theme["heading_font"], size=theme["title_pt"] + 8,
             color=title_color, bold=True, line_spacing=1.0,
             anchor=MSO_ANCHOR.BOTTOM)
    _check_orphan(s.get("title", ""), SLIDE_W - Inches(2.0),
                  theme["title_pt"] + 8, theme["heading_font"])
    if s.get("subtitle"):
        add_text(slide, s["subtitle"], Inches(1.1), Inches(4.65),
                 SLIDE_W - Inches(2.0), Inches(1.0),
                 font=theme["body_font"], size=theme["body_pt"] + 4,
                 color=sub_color, line_spacing=1.1)
    meta = "   |   ".join(x for x in [s.get("author"), s.get("date")] if x)
    if meta:
        add_text(slide, meta, Inches(1.1), SLIDE_H - Inches(1.0),
                 SLIDE_W - Inches(2.0), Inches(0.4),
                 font=theme["body_font"], size=13, color=meta_color)


def layout_section(slide, theme, s):
    # Full-bleed divider slide.
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_hex=theme.get("bleed", theme["primary"]))
    if s.get("number"):
        add_text(slide, str(s["number"]), Inches(1.1), Inches(1.6),
                 Inches(4.0), Inches(1.5),
                 font=theme["heading_font"], size=90, color=theme["accent"],
                 bold=True, fit=False)
    add_text(slide, s.get("title", ""), Inches(1.1), Inches(3.4),
             SLIDE_W - Inches(2.2), Inches(1.6),
             font=theme["heading_font"], size=theme["title_pt"],
             color=theme.get("on_bleed", theme["on_primary"]), bold=True, line_spacing=1.0)
    if s.get("subtitle"):
        add_text(slide, s["subtitle"], Inches(1.1), Inches(5.1),
                 SLIDE_W - Inches(2.2), Inches(1.2),
                 font=theme["body_font"], size=theme["body_pt"] + 2,
                 color=theme.get("on_bleed", theme["on_primary"]), line_spacing=1.15)
    add_rect(slide, Inches(1.1), Inches(3.2), Inches(2.2), Inches(0.08),
             fill_hex=theme["accent"])


def layout_bullets(slide, theme, s):
    fill_background(slide, theme)
    body_top = _slide_title(slide, theme, s.get("title", ""),
                            eyebrow=s.get("eyebrow"), subtitle=s.get("subtitle"))
    if s.get("lead"):
        add_text(slide, s["lead"], MARGIN, body_top,
                 SLIDE_W - 2 * MARGIN, Inches(0.7),
                 font=theme["body_font"], size=theme["body_pt"] + 2,
                 color=theme["secondary"], italic=True, line_spacing=1.1)
        body_top = Emu(body_top + Inches(0.8))
    bullets = s.get("bullets", [])
    paras = _bullet_paragraphs(theme, bullets, numbered=s.get("numbered", False),
                               start=s.get("start_at", 1))
    _check_sparse(paras, theme, SLIDE_W - 2 * MARGIN, _body_height(body_top))
    add_text(slide, paras, MARGIN, body_top,
             SLIDE_W - 2 * MARGIN, _body_height(body_top),
             font=theme["body_font"], size=theme["body_pt"], color=theme["text"],
             line_spacing=1.1, anchor=MSO_ANCHOR.TOP)


def layout_content(slide, theme, s):
    # Free-form paragraphs (no bullets). `body` is a string or list of strings.
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    body = s.get("body", "")
    if isinstance(body, str):
        body = [body]
    paras = [{"text": p, "size": theme["body_pt"], "color": theme["text"],
              "space_after": 12} for p in body]
    add_text(slide, paras, MARGIN, top,
             SLIDE_W - 2 * MARGIN, _body_height(top),
             font=theme["body_font"], size=theme["body_pt"],
             color=theme["text"], line_spacing=1.25)


def _column_block(slide, theme, col, left, top, width, height):
    """Render one column of a two-up layout. `col` has {heading, bullets|body}."""
    y = top
    if col.get("heading"):
        add_text(slide, col["heading"], left, y, width, Inches(0.5),
                 font=theme["heading_font"], size=theme["body_pt"] + 4,
                 color=theme["secondary"], bold=True)
        y = Emu(y + Inches(0.6))
    if col.get("bullets"):
        add_text(slide, _bullet_paragraphs(theme, col["bullets"],
                                           numbered=col.get("numbered", False)),
                 left, y, width, Emu(top + height - y),
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], line_spacing=1.1)
    elif col.get("body"):
        body = col["body"] if isinstance(col["body"], list) else [col["body"]]
        paras = [{"text": p, "space_after": 10} for p in body]
        add_text(slide, paras, left, y, width, Emu(top + height - y),
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], line_spacing=1.2)


def layout_two_column(slide, theme, s):
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    gap = Inches(0.6)
    col_w = (SLIDE_W - 2 * MARGIN - gap) / 2
    top = max(top, Inches(1.9))
    height = _body_height(top)
    _column_block(slide, theme, s.get("left", {}), MARGIN, top, col_w, height)
    _column_block(slide, theme, s.get("right", {}),
                  Emu(MARGIN + col_w + gap), top, col_w, height)


def _bullets_height(theme, items, width):
    paras = _bullet_paragraphs(theme, items)
    measure = [dict(p, font=theme["body_font"],
                    indent_pt=Inches(_BULLET_GEOM[p["level"]][0]) / 12700)
               for p in paras]
    return Pt(textfit.paragraphs_height_pt(measure, _pt(width), 1.1))


def layout_comparison(slide, theme, s):
    # Two contrasting cards side by side (e.g. Before/After, A/B, Pros/Cons).
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    gap = Inches(0.5)
    col_w = (SLIDE_W - 2 * MARGIN - gap) / 2
    top = max(top, Inches(1.9))
    avail = _body_height(top)
    cards = [s.get("left", {}), s.get("right", {})]
    head_fills = [theme["primary"], theme["accent"]]
    head_text = [theme["on_primary"], theme.get("on_accent", theme["on_primary"])]
    text_w = Emu(col_w - Inches(0.6))
    # Size both cards to the taller content instead of always filling the
    # slide, so short lists don't float in a big empty panel.
    need = max(_bullets_height(theme, c.get("bullets", []), text_w) for c in cards)
    height = int(min(avail, max(Inches(2.6), need + Inches(1.35))))
    for i, card in enumerate(cards):
        left = Emu(MARGIN + i * (col_w + gap))
        add_rect(slide, left, top, col_w, height, fill_hex=theme["surface"])
        add_rect(slide, left, top, col_w, Inches(0.7), fill_hex=head_fills[i])
        add_text(slide, card.get("heading", ""), Emu(left + Inches(0.3)), top,
                 text_w, Inches(0.7),
                 font=theme["heading_font"], size=theme["body_pt"] + 4,
                 color=head_text[i], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, _bullet_paragraphs(theme, card.get("bullets", []),
                                           numbered=card.get("numbered", False)),
                 Emu(left + Inches(0.3)), Emu(top + Inches(0.95)),
                 text_w, Emu(height - Inches(1.2)),
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], line_spacing=1.1)


def layout_metrics(slide, theme, s):
    # Up to 4 KPI cards in a row, centered vertically in the body area.
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    metrics = s.get("metrics", [])[:4]
    n = max(len(metrics), 1)
    gap = Inches(0.4)
    total_w = SLIDE_W - 2 * MARGIN
    card_w = (total_w - gap * (n - 1)) / n
    card_h = Inches(2.9)
    top = max(top, Emu(int(top + (_body_height(top) - card_h) / 2)))
    for i, m in enumerate(metrics):
        left = Emu(MARGIN + i * (card_w + gap))
        add_rect(slide, left, top, card_w, card_h, fill_hex=theme["surface"])
        add_rect(slide, left, top, card_w, Inches(0.12), fill_hex=theme["accent"])
        inner_l, inner_w = Emu(left + Inches(0.2)), Emu(card_w - Inches(0.4))
        # A KPI value must stay on one line: scale it to the card width.
        value = str(m.get("value", ""))
        vw = textfit.text_width_pt(value, 48, theme["heading_font"], bold=True)
        vscale = min(1.0, _pt(inner_w) * textfit.WIDTH_SLACK / max(vw, 1.0))
        _fit_record("metric", vscale, vscale >= 0.5, value)
        add_text(slide, value, inner_l, Emu(top + Inches(0.45)),
                 inner_w, Inches(1.1),
                 font=theme["heading_font"], size=round(48 * max(vscale, 0.5), 1),
                 color=theme["primary"], bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE, wrap=False, fit=False)
        add_text(slide, m.get("label", ""), inner_l, Emu(top + Inches(1.65)),
                 inner_w, Inches(0.5),
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], bold=True, align=PP_ALIGN.CENTER)
        if m.get("sublabel"):
            add_text(slide, m["sublabel"], inner_l,
                     Emu(top + Inches(2.15)), inner_w,
                     Inches(0.6), font=theme["body_font"], size=13,
                     color=theme["text_muted"], align=PP_ALIGN.CENTER,
                     line_spacing=1.0)


def layout_quote(slide, theme, s):
    fill_background(slide, theme)
    # Bar-style themes (e.g. samsung/report) carry a header bar on every content
    # slide; draw it here too so the quote slide matches the rest of the deck
    # instead of floating on a bare background. The quote then sits below the bar.
    if theme.get("header_style") == "bar":
        _header_bar(slide, theme, s.get("title", ""))
        qm_top, q_top, attr_top = Inches(1.6), Inches(2.8), Inches(5.9)
    else:
        qm_top, q_top, attr_top = Inches(0.9), Inches(2.3), Inches(5.6)
    add_text(slide, "“", Inches(0.8), qm_top, Inches(3.0), Inches(2.0),
             font=theme["heading_font"], size=160, color=theme["accent"],
             bold=True, fit=False)
    add_text(slide, s.get("quote", ""), Inches(1.6), q_top,
             SLIDE_W - Inches(3.2), Inches(2.8),
             font=theme["heading_font"], size=theme["heading_pt"] + 2,
             color=theme["primary"], italic=True, line_spacing=1.2,
             anchor=MSO_ANCHOR.MIDDLE)
    if s.get("attribution"):
        add_text(slide, "— " + s["attribution"], Inches(1.6), attr_top,
                 SLIDE_W - Inches(3.2), Inches(0.6),
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text_muted"], bold=True)


def _image_dims(path):
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size  # (w, h) in px
    except Exception:
        return None


def _place_image(slide, path, box_left, box_top, box_w, box_h, cover=False,
                 valign="middle"):
    """Place an image inside the given box, centered. `cover` fills the box
    and crops the overflow (for full-bleed); otherwise it is fitted (contain).
    The image is never stretched."""
    dims = _image_dims(path)
    if not dims:
        # Missing/unreadable image: draw a labeled placeholder instead of crashing.
        add_rect(slide, box_left, box_top, box_w, box_h, fill_hex="DDDDDD")
        add_text(slide, "[image not found:\n%s]" % os.path.basename(path),
                 box_left, box_top, box_w, box_h, font="Arial", size=12,
                 color="888888", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return
    iw, ih = dims
    if cover:
        pic = slide.shapes.add_picture(path, int(box_left), int(box_top),
                                       width=int(box_w), height=int(box_h))
        box_ratio, img_ratio = box_w / box_h, iw / ih
        if img_ratio > box_ratio:      # too wide: crop left/right
            excess = 1 - box_ratio / img_ratio
            pic.crop_left = pic.crop_right = excess / 2
        else:                          # too tall: crop top/bottom
            excess = 1 - img_ratio / box_ratio
            pic.crop_top = pic.crop_bottom = excess / 2
        return
    scale = min(box_w / iw, box_h / ih)
    w = int(iw * scale)
    h = int(ih * scale)
    left = int(box_left + (box_w - w) / 2)
    top = int(box_top + ((box_h - h) / 2 if valign == "middle" else 0))
    slide.shapes.add_picture(path, left, top, width=w, height=h)


def layout_image(slide, theme, s):
    """Image with optional text. `position`: right (default), left, or full."""
    fill_background(slide, theme)
    pos = s.get("position", "right")
    img = s.get("image")
    if pos == "full":
        if img:
            _place_image(slide, img, 0, 0, SLIDE_W, SLIDE_H, cover=True)
        if s.get("title"):
            add_rect(slide, 0, SLIDE_H - Inches(1.3), SLIDE_W, Inches(1.3),
                     fill_hex=theme["primary"])
            add_text(slide, s["title"], MARGIN, SLIDE_H - Inches(1.15),
                     SLIDE_W - 2 * MARGIN, Inches(1.0),
                     font=theme["heading_font"], size=theme["heading_pt"],
                     color=theme["on_primary"], bold=True,
                     anchor=MSO_ANCHOR.MIDDLE)
        return
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    half = (SLIDE_W - 2 * MARGIN - Inches(0.5)) / 2
    top = max(top, Inches(1.9))
    height = _body_height(top)
    if pos == "left":
        img_left, txt_left = MARGIN, Emu(MARGIN + half + Inches(0.5))
    else:
        txt_left, img_left = MARGIN, Emu(MARGIN + half + Inches(0.5))
    if img:
        # Top-aligned so the picture lines up with the text column beside it.
        _place_image(slide, img, img_left, top, half, height, valign="top")
    bullets = s.get("bullets")
    if bullets:
        add_text(slide, _bullet_paragraphs(theme, bullets), txt_left, top,
                 half, height, font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], line_spacing=1.15)
    elif s.get("body"):
        body = s["body"] if isinstance(s["body"], list) else [s["body"]]
        paras = [{"text": p, "space_after": 10} for p in body]
        add_text(slide, paras, txt_left, top, half, height,
                 font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], line_spacing=1.25)


def _column_widths(header, rows, ncols, total_w, size, font):
    """Share the table width by content: each column gets room in proportion
    to its widest cell (capped so one long column can't starve the others)."""
    need = []
    for c in range(ncols):
        cells = [header[c] if c < len(header) else ""] + \
                [r[c] if c < len(r) else "" for r in rows]
        widest = max(textfit.text_width_pt(str(x), size, font) for x in cells)
        need.append(min(max(widest, 40.0), _pt(total_w) * 0.55))
    total_need = sum(need)
    return [int(total_w * n / total_need) for n in need]


def layout_table(slide, theme, s):
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    tbl_spec = s.get("table", {})
    header = tbl_spec.get("header", [])
    rows = tbl_spec.get("rows", [])
    ncols = max(len(header), max((len(r) for r in rows), default=1))
    nrows = len(rows) + (1 if header else 0)
    if nrows == 0 or ncols == 0:
        return
    top = max(top, Inches(1.9))
    width = SLIDE_W - 2 * MARGIN
    avail = _body_height(top)
    size = theme["body_pt"] - 2
    widths = _column_widths(header, rows, ncols, width, size, theme["body_font"])
    all_rows = ([header] if header else []) + rows
    # Shrink the font until every row (at its wrapped line count) fits.
    pad = Inches(0.16)
    while True:
        heights = []
        for r in all_rows:
            lines = max(textfit.wrap_line_count(
                str(r[c]) if c < len(r) else "", _pt(widths[c] - Inches(0.24)),
                size, theme["body_font"]) for c in range(ncols))
            heights.append(int(Pt(size * 1.25 * lines) + pad))
        if sum(heights) <= avail or size <= 11:
            break
        size -= 1
    if size < theme["body_pt"] - 2:
        _fit_record("table", size / (theme["body_pt"] - 2.0),
                    sum(heights) <= avail, "table")
    gframe = slide.shapes.add_table(nrows, ncols, MARGIN, top, width,
                                    sum(heights))
    table = gframe.table
    for c, w in enumerate(widths):
        table.columns[c].width = w
    for r, h in enumerate(heights):
        table.rows[r].height = h
    r0 = 0
    if header:
        for c in range(ncols):
            cell = table.cell(0, c)
            cell.text = str(header[c]) if c < len(header) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(theme["primary"])
            _style_cell(cell, theme, theme["on_primary"], size, bold=True)
        r0 = 1
    for ri, row in enumerate(rows):
        for c in range(ncols):
            cell = table.cell(ri + r0, c)
            cell.text = str(row[c]) if c < len(row) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(
                theme["background"] if ri % 2 == 0 else theme["surface"])
            _style_cell(cell, theme, theme["text"], size)


def _style_cell(cell, theme, color_hex, size, bold=False):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.12)
    cell.margin_right = Inches(0.12)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.04)
    for p in cell.text_frame.paragraphs:
        _korean_wrap(p, "".join(r.text for r in p.runs))
        for run in p.runs:
            style_run(run, font=theme["body_font"], size=size,
                      color=color_hex, bold=bold)


_CHART_TYPES = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
    "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
}
_ROUND_CHARTS = {"pie", "doughnut"}


def _chart_palette(theme):
    """Series colors: the theme's own colors first, so charts match the deck
    instead of Office's default blue/orange."""
    return theme.get("chart_colors") or [
        theme["primary"], theme["accent"], theme["secondary"],
        theme["text_muted"], theme["surface"]]


def _fill(fmt, hexstr):
    fmt.fill.solid()
    fmt.fill.fore_color.rgb = _rgb(hexstr)


def _style_chart(chart, theme, spec, kind, n_series, n_points):
    palette = _chart_palette(theme)
    # One text style for the whole chart (axes, legend, labels), in the deck's
    # fonts and a color that reads on the slide background.
    chart.font.name = theme["body_font"]
    chart.font.size = Pt(14)
    chart.font.color.rgb = _rgb(theme["text_muted"])
    _set_ea_font(chart.font._rPr, "")

    plot = chart.plots[0]
    highlight = spec.get("highlight")
    if kind in _ROUND_CHARTS:
        plot.vary_by_categories = True
        for i, point in enumerate(plot.series[0].points):
            _fill(point.format, palette[i % len(palette)])
            point.format.line.color.rgb = _rgb(theme["background"])
    else:
        for i, ser in enumerate(plot.series):
            color = palette[i % len(palette)]
            if kind == "line":
                ser.format.line.color.rgb = _rgb(color)
                ser.format.line.width = Pt(3)
                ser.smooth = False
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = _rgb(color)
                ser.marker.format.line.color.rgb = _rgb(color)
            else:
                _fill(ser.format, color)
        # Single-series emphasis: grey every bar except the highlighted
        # category, which gets the accent color (Knaflic-style "focus").
        if highlight is not None and n_series == 1 and kind != "line":
            idx = (spec.get("categories", []).index(highlight)
                   if isinstance(highlight, str) else int(highlight))
            ser = plot.series[0]
            for i in range(n_points):
                _fill(ser.points[i].format,
                      theme["accent"] if i == idx else theme.get("muted_fill", "BFC5CC"))
        if kind in ("bar", "column", "stacked_bar", "stacked_column"):
            plot.gap_width = 60
            if kind.startswith("stacked"):
                plot.overlap = 100

        # Quiet axes: faint gridlines, no axis line on the value axis.
        va, ca = chart.value_axis, chart.category_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = _rgb(theme.get("grid", theme["surface"]))
        va.major_gridlines.format.line.width = Pt(0.75)
        va.format.line.fill.background()
        ca.format.line.color.rgb = _rgb(theme["text_muted"])
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        if kind in ("bar", "stacked_bar"):
            # Horizontal bars plot bottom-up; reverse so the first category
            # reads at the top, in spec order.
            ca.reverse_order = True
        if spec.get("number_format"):
            va.tick_labels.number_format = spec["number_format"]
            va.tick_labels.number_format_is_linked = False

    # Direct labels beat a legend: on by default for single-series and round
    # charts, where they don't collide.
    labels = spec.get("data_labels")
    if labels is None:
        labels = (n_series == 1 and n_points <= 12) or kind in _ROUND_CHARTS
    plot.has_data_labels = bool(labels)
    if labels:
        dl = plot.data_labels
        dl.font.size = Pt(13)
        dl.font.bold = True
        dl.font.color.rgb = _rgb(theme["on_primary"] if kind in _ROUND_CHARTS
                                 or kind.startswith("stacked") else theme["text"])
        if spec.get("number_format"):
            dl.number_format = spec["number_format"]
            dl.number_format_is_linked = False
        if kind in ("bar", "column"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        elif kind == "line":
            dl.position = XL_LABEL_POSITION.ABOVE
        if kind in _ROUND_CHARTS and spec.get("percent", True):
            dl.show_percentage = True
            dl.show_value = False
            dl.number_format = "0%"
            dl.number_format_is_linked = False


def layout_chart(slide, theme, s):
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    spec = s.get("chart", {})
    kind = spec.get("type", "column")
    if kind not in _CHART_TYPES:
        print("WARN: unknown chart type '%s' -> using 'column'" % kind,
              file=sys.stderr)
        kind = "column"
    data = CategoryChartData()
    data.categories = spec.get("categories", [])
    series = spec.get("series", [])
    if series and isinstance(series[0], dict):
        for ser in series:
            data.add_series(ser.get("name", ""), tuple(ser.get("values", [])))
        n_series = len(series)
        n_points = max(len(ser.get("values", [])) for ser in series)
    else:
        # Bare list of numbers => single unnamed series.
        data.add_series(spec.get("series_name", "Series 1"), tuple(series))
        n_series, n_points = 1, len(series)
    top = max(top, Inches(1.8))
    gframe = slide.shapes.add_chart(
        _CHART_TYPES[kind], MARGIN, top, SLIDE_W - 2 * MARGIN, _body_height(top),
        data)
    chart = gframe.chart
    chart.has_title = False
    if n_series > 1 or kind in _ROUND_CHARTS:
        chart.has_legend = True
        chart.legend.position = (XL_LEGEND_POSITION.RIGHT if kind in _ROUND_CHARTS
                                 else XL_LEGEND_POSITION.TOP)
        chart.legend.include_in_layout = False
    else:
        chart.has_legend = False
    _style_chart(chart, theme, spec, kind, n_series, n_points)


def layout_closing(slide, theme, s):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_hex=theme.get("bleed", theme["primary"]))
    add_rect(slide, 0, SLIDE_H - Inches(0.4), SLIDE_W, Inches(0.4),
             fill_hex=theme["accent"])
    add_text(slide, s.get("title", "Thank You"), MARGIN, Inches(2.6),
             SLIDE_W - 2 * MARGIN, Inches(1.6),
             font=theme["heading_font"], size=theme["title_pt"] + 6,
             color=theme.get("on_bleed", theme["on_primary"]), bold=True, align=PP_ALIGN.CENTER)
    if s.get("subtitle"):
        add_text(slide, s["subtitle"], MARGIN, Inches(4.3),
                 SLIDE_W - 2 * MARGIN, Inches(1.2),
                 font=theme["body_font"], size=theme["body_pt"] + 2,
                 color=theme.get("on_bleed", theme["on_primary"]), align=PP_ALIGN.CENTER,
                 line_spacing=1.2)


# --------------------------------------------------------------------------- #
# Diagrams — drawn as NATIVE PowerPoint shapes (autoshapes, connectors,
# freeforms), never flattened images, so every box, arrow, and label stays
# fully editable in PowerPoint/Keynote after the deck is generated.
# --------------------------------------------------------------------------- #
def _autoshape(slide, shape_type, left, top, w, h, fill, line=None, line_w=1.25):
    sp = slide.shapes.add_shape(shape_type, int(left), int(top), int(w), int(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _rgb(fill)
    if line:
        sp.line.color.rgb = _rgb(line)
        sp.line.width = Pt(line_w)
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def _shape_label(shape, theme, title, desc=None, color=None,
                 title_size=14, desc_size=11, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for attr in ("margin_left", "margin_right"):
        setattr(tf, attr, Inches(0.08))
    for attr in ("margin_top", "margin_bottom"):
        setattr(tf, attr, Inches(0.03))
    # Fit the label to the shape (minus margins) before writing it.
    measure = [{"text": title, "size": title_size, "bold": True,
                "font": theme["heading_font"]}]
    if desc:
        measure.append({"text": desc, "size": desc_size,
                        "font": theme["body_font"]})
    scale, fits = textfit.fit_scale(measure, _pt(shape.width - Inches(0.16)),
                                    _pt(shape.height - Inches(0.06)), 1.0, 0.65)
    _fit_record("diagram", scale, fits, title)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = title
    style_run(r, font=theme["heading_font"], size=round(title_size * scale, 1),
              color=color or theme["on_primary"], bold=True)
    _korean_wrap(p, title)
    if desc:
        p2 = tf.add_paragraph()
        p2.alignment = align
        p2.space_before = Pt(2)
        r2 = p2.add_run()
        r2.text = desc
        style_run(r2, font=theme["body_font"], size=round(desc_size * scale, 1),
                  color=color or theme["on_primary"])
        _korean_wrap(p2, desc)
    return shape


def _connect(slide, x1, y1, x2, y2, color, width=1.75, arrow=True, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                   int(x1), int(y1), int(x2), int(y2))
    c.line.color.rgb = _rgb(color)
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    if dash:
        # OOXML: a:prstDash must precede a:tailEnd in the line element, and the
        # color set above already added a:solidFill before this — so appending
        # here keeps the schema's child order valid.
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if arrow:
        ln.append(ln.makeelement(
            qn("a:tailEnd"), {"type": "triangle", "w": "lg", "len": "lg"}))
    c.shadow.inherit = False
    return c


def _freeform_poly(slide, pts, fill, line=None):
    ff = slide.shapes.build_freeform(float(pts[0][0]), float(pts[0][1]), scale=1.0)
    ff.add_line_segments([(float(x), float(y)) for x, y in pts[1:]], close=True)
    shp = ff.convert_to_shape()
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    if line:
        shp.line.color.rgb = _rgb(line)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _diag_area(top=None):
    """The drawing rectangle below the slide header: (left, top, width, height)."""
    top = max(top or 0, Inches(2.0))
    return MARGIN, top, SLIDE_W - 2 * MARGIN, SLIDE_H - top - Inches(0.7)


def _band_colors(theme, n):
    # Alternate primary/secondary so adjacent shapes read as distinct.
    pair = [theme["primary"], theme["secondary"]]
    return [pair[i % 2] for i in range(n)]


def _diag_process(slide, theme, nodes, top=None):
    left, top, w, h = _diag_area(top)
    n = max(len(nodes), 1)
    arrow_w = Inches(0.5)
    box_w = (w - arrow_w * (n - 1)) / n
    box_h = Inches(1.9)
    box_top = top + (h - box_h) / 2
    colors = _band_colors(theme, n)
    for i, node in enumerate(nodes):
        x = left + i * (box_w + arrow_w)
        sp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, box_top, box_w, box_h,
                        colors[i])
        _shape_label(sp, theme, node.get("title", ""), node.get("desc"),
                     title_size=17, desc_size=13)
        if i < n - 1:
            ax = x + box_w
            ay = box_top + box_h / 2
            _autoshape(slide, MSO_SHAPE.RIGHT_ARROW, ax + Inches(0.05),
                       ay - Inches(0.18), arrow_w - Inches(0.1), Inches(0.36),
                       theme["accent"])


def _diag_cycle(slide, theme, nodes, top=None):
    left, top, w, h = _diag_area(top)
    n = max(len(nodes), 1)
    cx, cy = left + w / 2, top + h / 2
    box_w, box_h = Inches(2.5), Inches(1.1)
    radius = min(w, h) / 2 - box_h / 2
    colors = _band_colors(theme, n)
    centers = []
    for i in range(n):
        ang = -math.pi / 2 + i * 2 * math.pi / n  # start at top, clockwise
        # Stretch horizontally: the slide is wide, so an ellipse uses the
        # space better than a circle and keeps boxes from overlapping.
        px = cx + radius * 1.6 * math.cos(ang)
        py = cy + radius * math.sin(ang)
        centers.append((px, py))
    # Arrows first (so boxes sit on top), connecting node i -> i+1 around the
    # ring. Draw each from the edge of one box to the edge of the next (where the
    # center-to-center line crosses each box's rectangle) so the arrow spans the
    # gap and touches both boxes, instead of floating as a stub in the middle.
    hw, hh = box_w / 2, box_h / 2

    def _edge(px, py, ux, uy):
        tx = hw / abs(ux) if ux else float("inf")
        ty = hh / abs(uy) if uy else float("inf")
        t = min(tx, ty)
        return px + ux * t, py + uy * t

    for i in range(n):
        x1, y1 = centers[i]
        x2, y2 = centers[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        d = math.hypot(dx, dy) or 1
        ux, uy = dx / d, dy / d
        sx, sy = _edge(x1, y1, ux, uy)
        ex, ey = _edge(x2, y2, -ux, -uy)
        _connect(slide, sx, sy, ex, ey, theme["accent"], width=2.0)
    for i, node in enumerate(nodes):
        px, py = centers[i]
        sp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                        px - box_w / 2, py - box_h / 2, box_w, box_h, colors[i])
        _shape_label(sp, theme, node.get("title", ""), node.get("desc"),
                     title_size=15, desc_size=12)


def _diag_hierarchy(slide, theme, nodes, top=None):
    # Two levels: nodes[0] is the root, the rest are its children.
    left, top, w, h = _diag_area(top)
    if not nodes:
        return
    root, children = nodes[0], nodes[1:]
    root_w, root_h = Inches(3.2), Inches(1.1)
    root_x = left + (w - root_w) / 2
    rsp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, root_x, top, root_w, root_h,
                     theme["primary"])
    _shape_label(rsp, theme, root.get("title", ""), root.get("desc"),
                 title_size=20, desc_size=14)
    if not children:
        return
    m = len(children)
    gap = Inches(0.4)
    child_w = min((w - gap * (m - 1)) / m, Inches(3.2))
    total = child_w * m + gap * (m - 1)
    cstart = left + (w - total) / 2
    child_h = Inches(1.3)
    child_top = top + Inches(2.3)
    rcx, rcy = root_x + root_w / 2, top + root_h
    # Elbow connectors: one trunk down from the root, a bus across, and a drop
    # into each child — reads as an org chart rather than a fan of diagonals.
    bus_y = rcy + (child_top - rcy) / 2
    first_cx = cstart + child_w / 2
    last_cx = cstart + (m - 1) * (child_w + gap) + child_w / 2
    _connect(slide, rcx, rcy, rcx, bus_y, theme["secondary"], width=1.5, arrow=False)
    if m > 1:
        _connect(slide, first_cx, bus_y, last_cx, bus_y, theme["secondary"],
                 width=1.5, arrow=False)
    for i, child in enumerate(children):
        cx = cstart + i * (child_w + gap)
        ccx = cx + child_w / 2
        _connect(slide, ccx, bus_y, ccx, child_top, theme["secondary"],
                 width=1.5, arrow=False)
        csp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, child_top,
                         child_w, child_h, theme["surface"], line=theme["secondary"])
        _shape_label(csp, theme, child.get("title", ""), child.get("desc"),
                     color=theme["text"], title_size=17, desc_size=13)


def _diag_pyramid(slide, theme, nodes, top=None, funnel=False):
    left, top, w, h = _diag_area(top)
    m = max(len(nodes), 1)
    base_hw = (w * 0.72) / 2
    cx = left + w / 2
    level_h = h / m
    colors = _band_colors(theme, m)

    def hw(depth):  # half-width at fractional depth 0..m (0 = apex)
        frac = depth / m
        return base_hw * (frac if not funnel else (1 - frac))

    for i, node in enumerate(nodes):
        y_top = top + i * level_h
        y_bot = top + (i + 1) * level_h
        hwt, hwb = hw(i), hw(i + 1)
        if funnel:
            # Keep the funnel's last band from collapsing to a point: label
            # space matters more than geometric purity.
            hwb = max(hwb, base_hw * 0.18)
            hwt = max(hwt, base_hw * 0.18)
        pts = [(cx - hwt, y_top), (cx + hwt, y_top),
               (cx + hwb, y_bot), (cx - hwb, y_bot)]
        _freeform_poly(slide, pts, colors[i])
        # Label centered on the band (its own textbox — band polygon stays
        # clean). Narrow bands put the label beside the shape instead.
        lbl = node.get("title", "")
        if node.get("desc"):
            lbl += "  —  " + node["desc"]
        band_w = 2 * min(hwt, hwb) if (hwt and hwb) else 2 * max(hwt, hwb) * 0.5
        lbl_w = textfit.text_width_pt(lbl, 15, theme["heading_font"], bold=True)
        lbl_size = 15 if lbl_w <= _pt(band_w) * 0.85 else (
            12 if lbl_w * 12 / 15 <= _pt(band_w) * 0.85 else None)
        if lbl_size:
            add_text(slide, lbl, left, y_top, w, level_h,
                     font=theme["heading_font"], size=lbl_size,
                     color=theme["on_primary"], bold=True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        else:
            side_l = cx + max(hwt, hwb) + Inches(0.25)
            add_text(slide, lbl, side_l, y_top, left + w - side_l, level_h,
                     font=theme["heading_font"], size=15, color=theme["text"],
                     bold=True, anchor=MSO_ANCHOR.MIDDLE)


def _diag_funnel(slide, theme, nodes, top=None):
    _diag_pyramid(slide, theme, nodes, top=top, funnel=True)


def _diag_timeline(slide, theme, nodes, top=None):
    left, top, w, h = _diag_area(top)
    n = max(len(nodes), 1)
    line_y = top + h / 2
    pad = Inches(0.4)
    box_w = min(Inches(2.6), (w - pad) / max(1, (n + 1) // 2) - Inches(0.2))
    x0, x1 = left + box_w / 2, left + w - box_w / 2
    # The spine.
    _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, line_y - Inches(0.03),
               w, Inches(0.06), theme["secondary"])
    step = (x1 - x0) / (n - 1) if n > 1 else 0
    colors = _band_colors(theme, n)
    box_h = Inches(1.3)
    for i, node in enumerate(nodes):
        mx = (left + w / 2) if n == 1 else x0 + i * step
        # Marker dot on the spine.
        r = Inches(0.13)
        _autoshape(slide, MSO_SHAPE.OVAL, mx - r, line_y - r, 2 * r, 2 * r,
                   colors[i])
        above = (i % 2 == 0)
        by = (line_y - Inches(0.35) - box_h) if above else (line_y + Inches(0.35))
        bx = max(left, min(mx - box_w / 2, left + w - box_w))
        sp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, box_w,
                        box_h, theme["surface"], line=theme["secondary"])
        _shape_label(sp, theme, node.get("title", ""), node.get("desc"),
                     color=theme["text"], title_size=15, desc_size=12)
        # Stem from marker to the label box.
        sy = (by + box_h) if above else by
        _connect(slide, mx, line_y, mx, sy, theme["secondary"], width=1.25,
                 arrow=False)


def _diag_flowchart(slide, theme, spec):
    """Auto-laid-out flowchart from nodes + edges, drawn as native shapes.

    The "Mermaid-like" capability that stays editable: a layered (Sugiyama-style)
    layout positions an arbitrary directed graph — branches, merges, and
    loop-backs — then draws real rounded-rectangle boxes and arrow connectors.
    Unlike a Mermaid image, every box stays movable/re-typable.

    spec = {nodes: [{id, title, desc}],
            edges: [[from_id, to_id] | {from, to, label}],
            direction: "LR" (default) | "TD"}
    Layering and back-edge detection are shared (flowchart_layout.layer_graph);
    feedback edges are drawn as muted return paths so cycles stay compact.
    """
    nodes = spec.get("nodes", [])
    if not nodes:
        return
    td = str(spec.get("direction", "LR")).upper() == "TD"
    ids, nodemap = [], {}
    for i, nd in enumerate(nodes):
        nid = str(nd.get("id", i))
        ids.append(nid)
        nodemap[nid] = nd
    edges = [(a, b, l) for (a, b, l) in normalize_edges(spec.get("edges", []))
             if a in nodemap and b in nodemap]
    rank, layers, feedback = layer_graph(ids, edges)
    n_layers = max(layers) + 1
    max_in_layer = max(len(v) for v in layers.values())

    left, top, w, h = _diag_area()
    gap_x, gap_y = Inches(0.55), Inches(0.4)
    if td:  # layers stack top->down; siblings spread across each row
        box_w = min(Inches(2.6), (w - gap_x * (max_in_layer - 1)) / max_in_layer)
        box_h = min(Inches(1.05), (h - gap_y * (n_layers - 1)) / n_layers)
    else:   # layers march left->right; siblings stack within each column
        box_w = min(Inches(2.3), (w - gap_x * (n_layers - 1)) / n_layers)
        box_h = min(Inches(1.2), (h - gap_y * (max_in_layer - 1)) / max_in_layer)

    rects = {}
    for lvl in range(n_layers):
        group = layers.get(lvl, [])
        m = len(group)
        if td:
            row_w = m * box_w + (m - 1) * gap_x
            x0 = left + (w - row_w) / 2
            y = top + lvl * (box_h + gap_y)
            for j, n in enumerate(group):
                rects[n] = (x0 + j * (box_w + gap_x), y)
        else:
            col_h = m * box_h + (m - 1) * gap_y
            y0 = top + (h - col_h) / 2
            x = left + lvl * (box_w + gap_x)
            for j, n in enumerate(group):
                rects[n] = (x, y0 + j * (box_h + gap_y))

    def _mid(n):
        x, y = rects[n]
        return x + box_w / 2, y + box_h / 2

    # Connectors first so boxes sit on top.
    clearance = Inches(0.22)
    for a, b, label in edges:
        if a == b:
            continue  # self-loop: skip (rare); the box itself implies it
        ax, ay = rects[a]
        bx, by = rects[b]
        if (a, b) not in feedback:
            # Forward step: a straight solid arrow between adjacent box edges.
            color = theme["accent"]
            if td:
                p1 = (ax + box_w / 2, ay + box_h)   # bottom-mid
                p2 = (bx + box_w / 2, by)           # top-mid
            else:
                p1 = (ax + box_w, ay + box_h / 2)   # right-mid
                p2 = (bx, by + box_h / 2)           # left-mid
            _connect(slide, p1[0], p1[1], p2[0], p2[1], color, width=1.75)
            lx, ly = (p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2
        else:
            # Feedback / loop-back: route a DASHED orthogonal path AROUND the
            # boxes (over the top for LR, out the left for TD), with the
            # arrowhead only on the final leg landing on the target box. A
            # straight diagonal here would overlap the forward arrows and read
            # as pointing at the wrong box; routing clear of them and dashing it
            # makes the return direction unmistakable.
            col = theme["secondary"]
            if not td:  # LR flow -> out the right of the source, up, over, down
                yline = top - clearance        # above every box, below the header
                x_out = min(ax + box_w + clearance, left + w)  # right of the column
                xb, ya = bx + box_w / 2, ay + box_h / 2
                _connect(slide, ax + box_w, ya, x_out, ya, col, 1.5, arrow=False, dash=True)
                _connect(slide, x_out, ya, x_out, yline, col, 1.5, arrow=False, dash=True)
                _connect(slide, x_out, yline, xb, yline, col, 1.5, arrow=False, dash=True)
                _connect(slide, xb, yline, xb, by, col, 1.5, arrow=True, dash=True)
                lx, ly = (x_out + xb) / 2, yline - Inches(0.02)
            else:       # TD flow -> route out the left side
                xline = left - clearance if left - clearance > Inches(0.1) else Inches(0.1)
                ya, yb = ay + box_h / 2, by + box_h / 2
                _connect(slide, ax, ya, xline, ya, col, 1.5, arrow=False, dash=True)
                _connect(slide, xline, ya, xline, yb, col, 1.5, arrow=False, dash=True)
                _connect(slide, xline, yb, bx, yb, col, 1.5, arrow=True, dash=True)
                lx, ly = xline, (ya + yb) / 2
        if label:
            add_text(slide, str(label), lx - Inches(0.7), ly - Inches(0.22),
                     Inches(1.4), Inches(0.3), font=theme["body_font"], size=10,
                     color=theme["text_muted"], align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)

    colors = _band_colors(theme, n_layers)
    for n in ids:
        x, y = rects[n]
        sp = _autoshape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, box_w, box_h,
                        colors[rank[n]])
        nd = nodemap[n]
        _shape_label(sp, theme, nd.get("title", n), nd.get("desc"),
                     title_size=13, desc_size=10)


_DIAGRAMS = {
    "process": _diag_process,
    "cycle": _diag_cycle,
    "hierarchy": _diag_hierarchy,
    "pyramid": _diag_pyramid,
    "funnel": _diag_funnel,
    "timeline": _diag_timeline,
    "flowchart": _diag_flowchart,
}


def layout_diagram(slide, theme, s):
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    spec = s.get("diagram", {})
    dtype = spec.get("type", "process")
    nodes = spec.get("nodes", [])
    fn = _DIAGRAMS.get(dtype, _diag_process)
    if dtype not in _DIAGRAMS:
        print("WARN: unknown diagram type '%s' -> using 'process'" % dtype,
              file=sys.stderr)
    # flowchart needs the edges too, so it takes the whole diagram spec; the
    # other types are positional patterns that only need the node list.
    if dtype == "flowchart":
        _diag_flowchart(slide, theme, spec)
    else:
        fn(slide, theme, nodes, top=top)


# --------------------------------------------------------------------------- #
# Structured layouts: agenda, big number, cards, 2x2 matrix
# --------------------------------------------------------------------------- #
def _item(x):
    """Normalize a string or {title, desc} item."""
    return {"title": x, "desc": None} if isinstance(x, str) else \
        {"title": x.get("title", ""), "desc": x.get("desc")}


def layout_agenda(slide, theme, s):
    """Numbered agenda rows with number badges. `current` (1-based) marks the
    part being presented and dims the rest — reuse the same agenda slide as
    a progress tracker between sections."""
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    items = [_item(x) for x in s.get("items", [])][:8]
    n = max(len(items), 1)
    current = s.get("current")
    two_cols = n > 5
    cols = 2 if two_cols else 1
    per_col = (n + cols - 1) // cols
    gap_x = Inches(0.6)
    col_w = (SLIDE_W - 2 * MARGIN - gap_x * (cols - 1)) / cols
    avail = _body_height(top)
    row_h = min(Inches(1.05), avail / per_col)
    badge = min(Inches(0.62), row_h - Inches(0.2))
    for i, it in enumerate(items):
        c, r = divmod(i, per_col)
        x = Emu(int(MARGIN + c * (col_w + gap_x)))
        y = Emu(int(top + r * row_h))
        active = current is None or current == i + 1
        fill = theme["accent"] if active else theme["surface"]
        num_color = theme.get("on_accent", theme["on_primary"]) if active else theme["text_muted"]
        b = _autoshape(slide, MSO_SHAPE.OVAL, x, y + (row_h - badge) / 2,
                       badge, badge, fill)
        _shape_label(b, theme, "%d" % (i + 1), color=num_color, title_size=20)
        text_l = Emu(int(x + badge + Inches(0.3)))
        paras = [{"text": it["title"], "size": theme["body_pt"] + 4, "bold": True,
                  "color": theme["text"] if active else theme["text_muted"]}]
        if it["desc"]:
            paras.append({"text": it["desc"], "size": theme["body_pt"] - 2,
                          "color": theme["text_muted"]})
        add_text(slide, paras, text_l, y, Emu(int(col_w - badge - Inches(0.3))),
                 row_h, font=theme["body_font"], size=theme["body_pt"],
                 color=theme["text"], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)


def layout_big_number(slide, theme, s):
    """One hero number with its label and a one-line takeaway band — for the
    single figure the audience must remember."""
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    value = str(s.get("value", ""))
    avail_w = SLIDE_W - 2 * MARGIN
    size = 130.0
    vw = textfit.text_width_pt(value, size, theme["heading_font"], bold=True)
    size = max(60.0, min(size, size * _pt(avail_w) * textfit.WIDTH_SLACK / max(vw, 1.0)))
    band_h = Inches(1.0) if s.get("takeaway") else 0
    body_h = _body_height(top) - band_h
    value_h = Pt(size * 1.15)
    label_h = Inches(0.6) if s.get("label") else 0
    block = value_h + label_h
    y = Emu(int(top + max(0, (body_h - block) / 2)))
    add_text(slide, value, MARGIN, y, avail_w, value_h,
             font=theme["heading_font"], size=round(size, 1), color=theme["accent"],
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.0, wrap=False, fit=False)
    if s.get("label"):
        add_text(slide, s["label"], MARGIN, Emu(int(y + value_h)), avail_w, label_h,
                 font=theme["body_font"], size=theme["body_pt"] + 6,
                 color=theme["text"], bold=True, align=PP_ALIGN.CENTER)
    if s.get("takeaway"):
        band_top = Emu(int(SLIDE_H - BOTTOM - band_h + Inches(0.1)))
        add_rect(slide, MARGIN, band_top, avail_w, band_h - Inches(0.2),
                 fill_hex=theme["surface"])
        add_rect(slide, MARGIN, band_top, Inches(0.12), band_h - Inches(0.2),
                 fill_hex=theme["accent"])
        add_text(slide, s["takeaway"], Emu(int(MARGIN + Inches(0.4))), band_top,
                 Emu(int(avail_w - Inches(0.6))), band_h - Inches(0.2),
                 font=theme["body_font"], size=theme["body_pt"] + 2,
                 color=theme["text"], anchor=MSO_ANCHOR.MIDDLE)


def layout_cards(slide, theme, s):
    """2–4 parallel cards in a row: options, pillars, features. Each card has
    a heading, an optional big `value`, and bullets or a short `body`."""
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    cards = s.get("cards", [])[:4]
    n = max(len(cards), 1)
    gap = Inches(0.4)
    card_w = (SLIDE_W - 2 * MARGIN - gap * (n - 1)) / n
    top = max(top, Inches(1.9))
    avail = _body_height(top)
    pad = Inches(0.3)
    text_w = Emu(int(card_w - 2 * pad))
    has_value = any(c.get("value") for c in cards)

    def body_paras(c):
        if c.get("bullets"):
            return _bullet_paragraphs(theme, c["bullets"])
        body = c.get("body") or []
        body = body if isinstance(body, list) else [body]
        return [{"text": t, "space_after": 6} for t in body]

    head_h = Inches(0.95)
    value_h = Inches(0.95) if has_value else 0
    need = 0
    for c in cards:
        m = [dict(p, font=theme["body_font"], size=p.get("size", theme["body_pt"]),
                  indent_pt=(Inches(_BULLET_GEOM[p.get("level", 0)][0]) / 12700
                             if p.get("bullet") else 0)) for p in body_paras(c)]
        need = max(need, textfit.paragraphs_height_pt(m, _pt(text_w), 1.1))
    height = int(min(avail, max(Inches(2.6), head_h + value_h + Pt(need) + Inches(0.5))))
    for i, c in enumerate(cards):
        left = Emu(int(MARGIN + i * (card_w + gap)))
        add_rect(slide, left, top, card_w, height, fill_hex=theme["surface"])
        add_rect(slide, left, top, card_w, Inches(0.1), fill_hex=theme["accent"])
        title_l, title_w = Emu(int(left + pad)), text_w
        if c.get("icon"):
            isz = Inches(0.5)
            if icons.add_icon(slide, c["icon"], left + pad, top + Inches(0.33), isz,
                              theme["accent"]) is not None:
                title_l = Emu(int(left + pad + isz + Inches(0.15)))
                title_w = Emu(int(text_w - isz - Inches(0.15)))
        add_text(slide, c.get("title", ""), title_l, Emu(int(top + Inches(0.3))),
                 title_w, Inches(0.6), font=theme["heading_font"],
                 size=theme["body_pt"] + 4, color=theme["primary"], bold=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y = top + head_h
        if has_value:
            if c.get("value"):
                add_text(slide, str(c["value"]), Emu(int(left + pad)), Emu(int(y)),
                         text_w, value_h, font=theme["heading_font"], size=36,
                         color=theme["accent"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            y += value_h
        add_text(slide, body_paras(c), Emu(int(left + pad)), Emu(int(y)), text_w,
                 Emu(int(top + height - y - Inches(0.25))), font=theme["body_font"],
                 size=theme["body_pt"], color=theme["text"], line_spacing=1.1)


def layout_matrix(slide, theme, s):
    """2x2 matrix. `quadrants` in reading order: top-left, top-right,
    bottom-left, bottom-right; each {title, items[]}. Optional axis labels
    `x_axis` / `y_axis` = {"low": "...", "high": "..."} and `highlight`
    (1–4) for the quadrant the slide argues for."""
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    quads = (s.get("quadrants", []) + [{}] * 4)[:4]
    xa, ya = s.get("x_axis") or {}, s.get("y_axis") or {}
    axis_w = Inches(0.75) if ya else 0
    axis_h = Inches(0.45) if xa else 0
    gap = Inches(0.18)
    left0 = Emu(int(MARGIN + axis_w))
    grid_w = SLIDE_W - MARGIN - left0
    grid_h = _body_height(top) - axis_h
    qw = (grid_w - gap) / 2
    qh = (grid_h - gap) / 2
    hl = s.get("highlight")
    for i, q in enumerate(quads):
        r, c = divmod(i, 2)
        x = Emu(int(left0 + c * (qw + gap)))
        y = Emu(int(top + r * (qh + gap)))
        on = hl == i + 1
        fill = theme["primary"] if on else theme["surface"]
        head = theme["on_primary"] if on else theme["primary"]
        body = theme["on_primary"] if on else theme["text"]
        add_rect(slide, x, y, qw, qh, fill_hex=fill)
        paras = [{"text": q.get("title", ""), "size": theme["body_pt"] + 2,
                  "bold": True, "color": head, "space_after": 6}]
        paras += [dict(p, color=body, bullet_color=theme["accent"])
                  for p in _bullet_paragraphs(theme, q.get("items", []))]
        add_text(slide, paras, Emu(int(x + Inches(0.25))), Emu(int(y + Inches(0.18))),
                 Emu(int(qw - Inches(0.5))), Emu(int(qh - Inches(0.3))),
                 font=theme["body_font"], size=theme["body_pt"], color=body,
                 line_spacing=1.05)
    muted = theme["text_muted"]
    if xa:
        y = Emu(int(top + grid_h + Inches(0.08)))
        add_text(slide, "← " + xa.get("low", ""), left0, y, qw, Inches(0.35),
                 font=theme["body_font"], size=13, color=muted)
        add_text(slide, xa.get("high", "") + " →", Emu(int(left0 + qw + gap)), y, qw,
                 Inches(0.35), font=theme["body_font"], size=13, color=muted,
                 align=PP_ALIGN.RIGHT)
    if ya:
        # Unrotated, word-wrapped in the left gutter: rotated text boxes
        # render inconsistently across PowerPoint/Keynote/LibreOffice.
        for label, yy in (("↑ " + ya.get("high", ""), top),
                          (ya.get("low", "") + " ↓", top + qh + gap)):
            add_text(slide, label, MARGIN, Emu(int(yy)),
                     Emu(int(axis_w - Inches(0.12))), qh,
                     font=theme["body_font"], size=13, color=muted,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def layout_icon_grid(slide, theme, s):
    """3–6 items, each an icon in a soft circle with a short title and
    description: capabilities, benefits, principles. An unknown icon falls
    back to the item's number so the slide still reads."""
    fill_background(slide, theme)
    top = _slide_title(slide, theme, s.get("title", ""), eyebrow=s.get("eyebrow"),
                       subtitle=s.get("subtitle"))
    items = s.get("items", [])[:6]
    n = max(len(items), 1)
    cols = n if n <= 4 else 3
    rows = (n + cols - 1) // cols
    gap_x, gap_y = Inches(0.45), Inches(0.35)
    top = max(top, Inches(1.9))
    avail = _body_height(top)
    cell_w = (SLIDE_W - 2 * MARGIN - gap_x * (cols - 1)) / cols
    cell_h = min((avail - gap_y * (rows - 1)) / rows, Inches(2.6))
    block_h = cell_h * rows + gap_y * (rows - 1)
    top = Emu(int(top + max(0, (avail - block_h) / 2)))
    circle = min(Inches(1.0), cell_h * 0.42)
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        # Center a short last row.
        in_row = min(cols, n - r * cols)
        x0 = MARGIN + (cols - in_row) * (cell_w + gap_x) / 2
        x = Emu(int(x0 + c * (cell_w + gap_x)))
        y = Emu(int(top + r * (cell_h + gap_y)))
        cx = Emu(int(x + (cell_w - circle) / 2))
        _autoshape(slide, MSO_SHAPE.OVAL, cx, y, circle, circle, theme["surface"])
        isz = circle * 0.55
        drawn = icons.add_icon(slide, it.get("icon"), cx + (circle - isz) / 2,
                               y + (circle - isz) / 2, isz, theme["accent"])
        if drawn is None:
            badge = _autoshape(slide, MSO_SHAPE.OVAL, cx, y, circle, circle,
                               theme["accent"])
            _shape_label(badge, theme, "%d" % (i + 1),
                         color=theme.get("on_accent", theme["on_primary"]),
                         title_size=22)
        paras = [{"text": it.get("title", ""), "size": theme["body_pt"] + 2,
                  "bold": True, "color": theme["text"], "space_after": 4}]
        if it.get("desc"):
            paras.append({"text": it["desc"], "size": theme["body_pt"] - 3,
                          "color": theme["text_muted"]})
        ty = Emu(int(y + circle + Inches(0.15)))
        add_text(slide, paras, x, ty, cell_w, Emu(int(y + cell_h - ty)),
                 font=theme["body_font"], size=theme["body_pt"], color=theme["text"],
                 align=PP_ALIGN.CENTER, line_spacing=1.05)



LAYOUTS = {
    "title": layout_title,
    "diagram": layout_diagram,
    "section": layout_section,
    "bullets": layout_bullets,
    "content": layout_content,
    "two_column": layout_two_column,
    "comparison": layout_comparison,
    "metrics": layout_metrics,
    "quote": layout_quote,
    "image": layout_image,
    "table": layout_table,
    "chart": layout_chart,
    "closing": layout_closing,
    "agenda": layout_agenda,
    "big_number": layout_big_number,
    "cards": layout_cards,
    "matrix": layout_matrix,
    "icon_grid": layout_icon_grid,
}


# --------------------------------------------------------------------------- #
# Deck-level setup
# --------------------------------------------------------------------------- #
def _set_theme_fonts(prs, theme):
    """Write the template fonts into the deck's theme (major = headings,
    minor = body; latin + East Asian). Text the user adds later in PowerPoint
    then matches the generated slides instead of falling back to Calibri."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    for tag, latin in (("a:majorFont", theme["heading_font"]),
                       ("a:minorFont", theme["body_font"])):
        node = root.find(".//" + qn(tag))
        if node is None:
            continue
        node.find(qn("a:latin")).set("typeface", latin)
        ea = node.find(qn("a:ea"))
        if ea is not None:
            ea.set("typeface", _CTX["ea_font"])
        for f in node.findall(qn("a:font")):
            if f.get("script") == "Hang":
                f.set("typeface", _CTX["ea_font"])
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                standalone=True)


def _add_notes(slide, notes):
    if not notes:
        return
    text = "\n".join(notes) if isinstance(notes, list) else str(notes)
    slide.notes_slide.notes_text_frame.text = text


def _check_bounds(slide, index):
    """Report shapes that extend past the slide edge (a classic generator bug
    that every renderer shows differently)."""
    tol = Inches(0.02)
    for shp in slide.shapes:
        if shp.left is None:
            continue
        if (shp.left < -tol or shp.top < -tol
                or shp.left + shp.width > SLIDE_W + tol
                or shp.top + shp.height > SLIDE_H + tol):
            _CTX["bounds"].append((index, shp.name))


def fit_report():
    """Lines describing text that was shrunk a lot or still overflows, and
    shapes past the slide edge, from the most recent build()."""
    lines = []
    for slide_no, parts in _CTX.get("splits", []):
        lines.append("SPLIT: slide %d: too long for one slide at a readable size "
                     "-> continued over %d slides (set \"split\": false to "
                     "keep it on one)" % (slide_no, parts))
    for slide_no, name in _CTX["bounds"]:
        lines.append("WARN: slide %s: shape '%s' extends past the slide edge"
                     % (slide_no, name))
    for slide_no, kind, scale, fits, snippet in _CTX["fit"]:
        if kind == "sparse":
            lines.append("SPARSE: slide %s: bullets fill only %d%% of the slide body "
                         "-> add substance, or use metrics / cards / diagram / "
                         "big_number / agenda" % (slide_no, round(scale * 100)))
        elif kind == "orphan":
            lines.append("ORPHAN: slide %s: title wraps leaving one short word "
                         "on its last line: \"%s\" -> shorten or rephrase"
                         % (slide_no, snippet))
        elif not fits:
            lines.append("OVERFLOW: slide %s: %s does not fit even at %d%% size: "
                         "\"%s\" -> split the slide or cut words"
                         % (slide_no, kind, round(scale * 100), snippet))
        elif scale < 0.85:
            lines.append("FIT: slide %s: %s shrunk to %d%%: \"%s\""
                         % (slide_no, kind, round(scale * 100), snippet))
    return lines



# --------------------------------------------------------------------------- #
# Continuation slides — content that cannot fit at a readable size is split
# at item boundaries onto "(cont.)" slides instead of being shrunk further.
# --------------------------------------------------------------------------- #
SPLIT_LAYOUTS = ("bullets", "table")
SPLIT_MIN_SCALE = 0.85     # below this the text is too small to read: split


def _probe_fits(theme, s):
    """Render one slide off-screen and report whether everything fits at
    >= SPLIT_MIN_SCALE. Leaves the real fit report untouched."""
    saved = (_CTX["fit"], _CTX["bounds"], _CTX["slide"])
    _CTX["fit"], _CTX["bounds"], _CTX["slide"] = [], [], "probe"
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        LAYOUTS[s.get("layout", "bullets")](_blank_slide(prs), theme, s)
        return all(fits and scale >= SPLIT_MIN_SCALE
                   for _n, kind, scale, fits, _t in _CTX["fit"]
                   if kind in ("text", "table"))
    finally:
        _CTX["fit"], _CTX["bounds"], _CTX["slide"] = saved


def _split_units(s):
    """Indivisible pieces of a splittable slide: a top-level bullet together
    with its sub-bullets, or one table row."""
    if s.get("layout", "bullets") == "table":
        return [[r] for r in s.get("table", {}).get("rows", [])]
    units = []
    for item in s.get("bullets", []):
        level = 0 if isinstance(item, str) else item.get("level", 0)
        if level == 0 or not units:
            units.append([item])
        else:
            units[-1].append(item)
    return units


def _with_units(s, units, part, start_at):
    part_s = copy.deepcopy(s)
    flat = [x for u in units for x in u]
    if s.get("layout", "bullets") == "table":
        part_s["table"]["rows"] = flat          # header repeats on every part
    else:
        part_s["bullets"] = flat
        part_s["start_at"] = start_at
    if part > 0:
        title = s.get("title", "")
        part_s["title"] = title + (" (계속)" if textfit.has_hangul(title) else " (cont.)")
        for k in ("subtitle", "lead", "notes"):
            part_s.pop(k, None)
    return part_s


def _split_slide(theme, s):
    """Return the list of slides `s` renders as: [s] when it fits, otherwise
    balanced continuation parts, each holding at least two units when
    possible."""
    units = _split_units(s)
    if len(units) < 2 or _probe_fits(theme, s):
        return [s]
    # Greedy pass: how many parts are needed?
    chunks, cur = [], []
    for u in units:
        trial = cur + [u]
        if cur and not _probe_fits(theme, _with_units(s, trial, len(chunks), 1)):
            chunks.append(cur)
            cur = [u]
        else:
            cur = trial
    chunks.append(cur)
    # Balanced pass: same number of parts, even sizes, if they all fit.
    k = len(chunks)
    size, extra = divmod(len(units), k)
    balanced, i = [], 0
    for j in range(k):
        n = size + (1 if j < extra else 0)
        balanced.append(units[i:i + n])
        i += n
    parts = balanced
    start = s.get("start_at", 1)
    out = []
    for j, chunk in enumerate(parts):
        part_s = _with_units(s, chunk, j, start)
        if not _probe_fits(theme, part_s):
            parts = None
            break
        out.append(part_s)
        start += len(chunk)
    if parts is None:
        out, start = [], s.get("start_at", 1)
        for j, chunk in enumerate(chunks):
            out.append(_with_units(s, chunk, j, start))
            start += len(chunk)
    return out


def _expand_slides(theme, slides, deck_style):
    """[(label, slide)] with continuation parts inserted. Labels refer to the
    spec's slide numbers ("4", "4 (cont. 1)") so reports point at the JSON."""
    out, splits = [], []
    for i, s in enumerate(slides):
        parts = [s]
        _CTX["title_style"] = s.get("title_style", deck_style)
        if s.get("layout", "bullets") in SPLIT_LAYOUTS and s.get("split", True):
            parts = _split_slide(theme, s)
        if len(parts) > 1:
            splits.append((i + 1, len(parts)))
        for j, p in enumerate(parts):
            out.append(("%d" % (i + 1) if j == 0 else "%d (cont. %d)" % (i + 1, j), p))
    return out, splits


def build(spec, out_path, report=True):
    """Render `spec` to `out_path` (a path or a writable binary file object).
    With `report`, print the fit report to stderr afterwards."""
    theme_name = spec.get("template", DEFAULT_THEME)
    theme = copy.deepcopy(get_theme(theme_name))  # we may inject per-deck overrides
    if theme_name.strip().lower() not in THEMES:
        print("WARN: unknown template '%s' -> using '%s'"
              % (theme_name, DEFAULT_THEME), file=sys.stderr)
    # Per-deck override of the header classification label ("" disables it).
    if "header_label" in spec:
        theme["header_label"] = spec["header_label"]
    # Per-deck font overrides (e.g. a brand font, or "Pretendard" for Korean).
    if spec.get("font"):
        theme["heading_font"] = theme["body_font"] = spec["font"]
    _CTX["ea_font"] = spec.get("ea_font") or theme.get("ea_font", DEFAULT_EA_FONT)
    deck_style = spec.get("title_style", "topic")
    if deck_style not in TITLE_STYLES:
        print("WARN: unknown title_style '%s' -> using 'topic'" % deck_style,
              file=sys.stderr)
        deck_style = "topic"
    _CTX["fit"] = []
    _CTX["bounds"] = []
    theme.setdefault("bleed", theme["primary"])
    theme.setdefault("on_bleed", theme["on_primary"])

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    _set_theme_fonts(prs, theme)

    footer_default = spec.get("footer")
    show_numbers = spec.get("slide_numbers", True)
    if spec.get("auto_split", True):
        slides, _CTX["splits"] = _expand_slides(theme, spec.get("slides", []),
                                                deck_style)
    else:
        slides = [("%d" % (i + 1), s) for i, s in enumerate(spec.get("slides", []))]
        _CTX["splits"] = []

    for i, (label, s) in enumerate(slides):
        _CTX["slide"] = label
        # A slide may override the deck's title style (e.g. an agenda slide
        # keeps a "topic" label in an "assertion" deck).
        _CTX["title_style"] = s.get("title_style", deck_style)
        name = s.get("layout", "bullets")
        fn = LAYOUTS.get(name)
        slide = _blank_slide(prs)
        if fn is None:
            print("WARN: unknown layout '%s' on slide %s -> using 'bullets'"
                  % (name, label), file=sys.stderr)
            fn = layout_bullets
        fn(slide, theme, s)
        # Footers/numbers only on interior content slides, not full-bleed ones.
        if name not in ("title", "section", "closing", "image"):
            # Themes can opt out of footer text (e.g. samsung) so only the
            # page number remains on the right.
            ftext = s.get("footer", footer_default)
            if not theme.get("show_footer_text", True):
                ftext = None
            _footer(slide, theme,
                    slide_no=(i + 1) if show_numbers else None,
                    total=len(slides),
                    footer_text=ftext)
        # Classification marker (e.g. Confidential) on every slide's top-right.
        _classification_marker(slide, theme, name, bg_hex=s.get("background"))
        _add_notes(slide, s.get("notes"))
        _check_bounds(slide, label)

    prs.save(out_path)
    if report:
        lines = fit_report()
        for line in lines:
            print(line, file=sys.stderr)
        if not lines:
            print("Fit: all text fits at (or near) full size.", file=sys.stderr)
    return len(slides)


def _default_out(spec, spec_path):
    title = spec.get("title") or "presentation"
    slug = re.sub(r"[^\w.-]+", "_", title).strip("_") or "presentation"
    base = os.path.dirname(os.path.abspath(spec_path)) if spec_path and spec_path != "-" else os.getcwd()
    return os.path.join(base, slug + ".pptx")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render a presentation spec JSON into a .pptx")
    ap.add_argument("spec", help="Path to spec JSON, or '-' for stdin")
    ap.add_argument("-o", "--output", help="Output .pptx path")
    args = ap.parse_args(argv)

    if args.spec == "-":
        spec = json.load(sys.stdin)
    else:
        with open(args.spec, encoding="utf-8") as f:
            spec = json.load(f)

    out = args.output or _default_out(spec, args.spec)
    n = build(spec, out)
    print("Wrote %d slides -> %s" % (n, out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
