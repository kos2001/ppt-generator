#!/usr/bin/env python3
"""Validate a presentation spec before rendering.

Three passes, all printed as line-oriented feedback:

1. Structure — wrong layout name, missing required fields, malformed
   chart/table/diagram data. These are ERRORs.
2. Density & rhythm lint — too many bullets, bullets that wrap past two
   lines, three same-layout slides in a row, oversized tables, low-contrast
   theme colors. These are WARNs: fix them unless there's a reason not to.
3. Dry-run render (when python-pptx is installed) — the real renderer lays
   the deck out in memory and reports text it had to shrink (WARN) or that
   still overflows its box at the minimum size (ERROR).

Exits non-zero if any errors are found; warnings alone do not fail unless
--strict is given.

Usage:
    python validate_spec.py spec.json [--strict] [--no-render]
"""
import argparse
import io
import json
import os
import re
import sys

sys.dont_write_bytecode = True  # keep the skill directory free of __pycache__

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import textfit  # noqa: E402
import icons  # noqa: E402

VALID_LAYOUTS = {
    "title", "section", "bullets", "content", "two_column", "comparison",
    "metrics", "quote", "image", "table", "chart", "closing", "diagram",
    "agenda", "big_number", "cards", "matrix", "icon_grid",
}

VALID_DIAGRAMS = {"process", "cycle", "hierarchy", "pyramid", "funnel", "timeline"}
VALID_CHARTS = {"column", "bar", "stacked_column", "stacked_bar", "line",
                "pie", "doughnut", "area"}

# layout -> (required keys, at-least-one-of groups)
REQUIRED = {
    "title": (["title"], []),
    "section": (["title"], []),
    "bullets": (["title"], [["bullets"]]),
    "content": (["title"], [["body"]]),
    "two_column": (["title"], [["left", "right"]]),
    "comparison": (["title"], [["left", "right"]]),
    "metrics": (["title"], [["metrics"]]),
    "quote": (["quote"], []),
    "image": ([], [["image", "title"]]),
    "table": (["title"], [["table"]]),
    "chart": (["title"], [["chart"]]),
    "closing": (["title"], []),
    "diagram": (["title"], [["diagram"]]),
    "agenda": (["title"], [["items"]]),
    "big_number": (["title", "value"], []),
    "cards": (["title"], [["cards"]]),
    "matrix": (["title"], [["quadrants"]]),
    "icon_grid": (["title"], [["items"]]),
}

# Density budgets. Heuristics, not laws: they keep slides scannable at the
# 18pt body floor. Lines are measured, so they work for Korean and English.
MAX_BULLETS = 6
MAX_BULLET_LINES = 2          # a top-level bullet wrapping to 3+ lines is prose
BODY_WIDTH_PT = 862           # content width of a full-width text slide
COLUMN_WIDTH_PT = 410         # one column of two_column / comparison
MAX_TABLE_ROWS, MAX_TABLE_COLS = 8, 6
MAX_METRIC_CHARS = 8
MAX_DIAGRAM_NODES = 6
MAX_SAME_LAYOUT_RUN = 2       # three in a row is a warning

TITLE_STYLES = ("topic", "assertion")
# Slides whose title is always a label, whatever the deck's title_style.
LABEL_LAYOUTS = {"title", "section", "closing", "quote"}
MIN_ASSERTION_WORDS = 4       # "분기별 자동 해결률" is a label, not a claim
# A title that ends like a finished sentence: a period, or a Korean finite
# verb ending (했다/한다/됩니다 ...).
_SENTENCE_END = re.compile(
    r"([.。]|(했|한|된|됐|였|이|있|없|졌|겠)다|(습|합|됩|입)니다)\s*$")


def _lint_title(tag, title, style, warnings):
    if not title:
        return
    if _SENTENCE_END.search(title):
        if style == "topic":
            warnings.append("%s: title reads like a sentence: \"%s\" — topic "
                            "titles are noun phrases (or set \"title_style\": "
                            "\"assertion\")" % (tag, title))
        elif title.rstrip().endswith((".", "。")):
            warnings.append("%s: drop the trailing period from the title: \"%s\""
                            % (tag, title))
    elif style == "assertion" and len(title.split()) < MIN_ASSERTION_WORDS:
        warnings.append("%s: title \"%s\" is a label, not a takeaway — state "
                        "the slide's conclusion (subject + result), or set "
                        "\"title_style\": \"topic\" on this slide"
                        % (tag, title))


def _bullet_text(item):
    return item if isinstance(item, str) else item.get("text", "")


def _bullet_level(item):
    return 0 if isinstance(item, str) else item.get("level", 0)


def _lint_bullets(tag, items, width_pt, warnings):
    top = [b for b in items if _bullet_level(b) == 0]
    if len(top) > MAX_BULLETS:
        warnings.append("%s: %d top-level bullets (aim for <= %d) — split the "
                        "slide or move detail to notes" % (tag, len(top), MAX_BULLETS))
    for b in top:
        text = _bullet_text(b)
        lines = textfit.wrap_line_count(text, width_pt - 25, 18, "calibri")
        if lines > MAX_BULLET_LINES:
            warnings.append("%s: bullet wraps to ~%d lines — cut to a fragment: "
                            "\"%s\"" % (tag, lines, text[:50]))


def _luminance(hexstr):
    c = [int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(a, b):
    la, lb = _luminance(a), _luminance(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def lint_theme(theme, warnings):
    """WCAG 2.x contrast for the text/fill pairs the renderer actually draws.
    4.5:1 for body-size text, 3:1 for large/bold headings."""
    on_accent = theme.get("on_accent", theme["on_primary"])
    pairs = [
        ("text on background", theme["text"], theme["background"], 4.5),
        ("muted text on background", theme["text_muted"], theme["background"], 4.5),
        ("muted text on cards", theme["text_muted"], theme["surface"], 4.5),
        ("body text on cards", theme["text"], theme["surface"], 4.5),
        ("title on background", theme["primary"], theme["background"], 3.0),
        ("text on primary fills", theme["on_primary"], theme["primary"], 4.5),
        ("heading on accent fills", on_accent, theme["accent"], 3.0),
        ("text on full-bleed slides", theme.get("on_bleed", theme["on_primary"]),
         theme.get("bleed", theme["primary"]), 4.5),
    ]
    for name, fg, bg, need in pairs:
        ratio = contrast(fg, bg)
        if ratio < need:
            warnings.append("theme: %s is %.1f:1 (needs %.1f:1) — %s on %s"
                            % (name, ratio, need, fg, bg))


def validate(spec):
    errors, warnings = [], []

    if not isinstance(spec, dict):
        return ["spec must be a JSON object"], []
    if "slides" not in spec or not isinstance(spec["slides"], list):
        errors.append("spec must have a 'slides' array")
        return errors, warnings
    if not spec["slides"]:
        errors.append("'slides' is empty — add at least one slide")

    deck_style = spec.get("title_style", "topic")
    if deck_style not in TITLE_STYLES:
        errors.append("unknown title_style '%s' (valid: %s)"
                      % (deck_style, ", ".join(TITLE_STYLES)))
        deck_style = "topic"

    run_layout, run_len = None, 0
    for i, s in enumerate(spec["slides"]):
        tag = "slide %d" % (i + 1)
        if not isinstance(s, dict):
            errors.append("%s: must be an object" % tag)
            continue
        layout = s.get("layout", "bullets")
        if layout not in VALID_LAYOUTS:
            errors.append("%s: unknown layout '%s' (valid: %s)"
                          % (tag, layout, ", ".join(sorted(VALID_LAYOUTS))))
            continue
        style = s.get("title_style", deck_style)
        if style not in TITLE_STYLES:
            errors.append("%s: unknown title_style '%s'" % (tag, style))
        elif layout not in LABEL_LAYOUTS:
            _lint_title(tag, s.get("title", ""), style, warnings)
        req_keys, one_of_groups = REQUIRED[layout]
        for k in req_keys:
            if not s.get(k):
                errors.append("%s (%s): missing required field '%s'" % (tag, layout, k))
        for group in one_of_groups:
            if not any(s.get(k) for k in group):
                errors.append("%s (%s): needs at least one of %s"
                              % (tag, layout, group))

        # Rhythm: the same content layout three times running reads as a wall.
        if layout == run_layout:
            run_len += 1
        else:
            run_layout, run_len = layout, 1
        if run_len == MAX_SAME_LAYOUT_RUN + 1 and layout not in ("section",):
            alts = [x for x in ("metrics", "chart", "diagram", "comparison", "quote")
                    if x != layout]
            warnings.append("%s: third '%s' slide in a row — switch one to %s"
                            % (tag, layout, ", ".join(alts)))

        if layout == "chart":
            ch = s.get("chart", {})
            ctype = ch.get("type", "column")
            if ctype not in VALID_CHARTS:
                errors.append("%s: unknown chart type '%s' (valid: %s)"
                              % (tag, ctype, ", ".join(sorted(VALID_CHARTS))))
            cats = ch.get("categories", [])
            if not cats:
                warnings.append("%s: chart has no 'categories'" % tag)
            series = ch.get("series")
            if not series:
                errors.append("%s: chart has no 'series' data" % tag)
            else:
                named = isinstance(series[0], dict)
                for ser in (series if named else [{"values": series}]):
                    vals = ser.get("values", [])
                    if cats and len(vals) != len(cats):
                        errors.append("%s: series '%s' has %d values for %d categories"
                                      % (tag, ser.get("name", ""), len(vals), len(cats)))
                    if any(not isinstance(v, (int, float)) or isinstance(v, bool)
                           for v in vals):
                        errors.append("%s: series '%s' has non-numeric values "
                                      "(put units in number_format, not the data)"
                                      % (tag, ser.get("name", "")))
                if named and len(series) > 4:
                    warnings.append("%s: %d series — more than ~4 is hard to read"
                                    % (tag, len(series)))
            hl = ch.get("highlight")
            if hl is not None and isinstance(hl, str) and hl not in cats:
                errors.append("%s: highlight '%s' is not one of the categories"
                              % (tag, hl))
        if layout == "table":
            tb = s.get("table", {})
            rows = tb.get("rows", [])
            if not rows:
                errors.append("%s: table has no 'rows'" % tag)
            ncols = max([len(tb.get("header", []))] + [len(r) for r in rows])
            if len(rows) > MAX_TABLE_ROWS or ncols > MAX_TABLE_COLS:
                warnings.append("%s: table is %d rows x %d cols (aim for <= %dx%d) "
                                "— split across slides" % (tag, len(rows), ncols,
                                                           MAX_TABLE_ROWS, MAX_TABLE_COLS))
        if layout == "diagram":
            dg = s.get("diagram", {})
            dt = dg.get("type", "process")
            if dt not in VALID_DIAGRAMS:
                errors.append("%s: unknown diagram type '%s' (valid: %s)"
                              % (tag, dt, ", ".join(sorted(VALID_DIAGRAMS))))
            nodes = dg.get("nodes", [])
            if not nodes:
                errors.append("%s: diagram has no 'nodes'" % tag)
            limit = MAX_DIAGRAM_NODES + (1 if dt == "hierarchy" else 0)
            if len(nodes) > limit:
                warnings.append("%s: %d diagram nodes — more than %d gets cramped"
                                % (tag, len(nodes), limit))
        if layout == "metrics":
            ms = s.get("metrics", [])
            if len(ms) > 4:
                warnings.append("%s: more than 4 metrics — only the first 4 render" % tag)
            for m in ms[:4]:
                if len(str(m.get("value", ""))) > MAX_METRIC_CHARS:
                    warnings.append("%s: metric value '%s' is long — use a short "
                                    "number + unit, move detail to sublabel"
                                    % (tag, m.get("value")))
        if layout == "bullets":
            _lint_bullets(tag, s.get("bullets", []), BODY_WIDTH_PT, warnings)
        if layout == "agenda":
            items = s.get("items", [])
            if len(items) > 8:
                warnings.append("%s: %d agenda items — only the first 8 render"
                                % (tag, len(items)))
            cur = s.get("current")
            if cur is not None and not (isinstance(cur, int) and 1 <= cur <= len(items)):
                errors.append("%s: current must be 1..%d" % (tag, len(items)))
        if layout == "big_number" and len(str(s.get("value", ""))) > 10:
            warnings.append("%s: big_number value '%s' is long — a hero number "
                            "should read at a glance" % (tag, s.get("value")))
        if layout == "cards":
            n = len(s.get("cards", []))
            if not 2 <= n <= 4:
                warnings.append("%s: %d cards — use 2–4 (only the first 4 render)"
                                % (tag, n))
        if layout == "icon_grid":
            n = len(s.get("items", []))
            if not 3 <= n <= 6:
                warnings.append("%s: %d icon_grid items — use 3–6 (only the first "
                                "6 render)" % (tag, n))
        icon_refs = [it.get("icon") for it in s.get("items", [])
                     if layout == "icon_grid" and isinstance(it, dict)]
        icon_refs += [c.get("icon") for c in s.get("cards", []) if isinstance(c, dict)]
        for ref in icon_refs:
            if ref and icons.resolve(ref) is None:
                hint = icons.suggest(ref)
                warnings.append("%s: unknown icon '%s'%s — a number badge renders "
                                "instead (see references/icons.md)"
                                % (tag, ref, " (did you mean %s?)" % ", ".join(hint)
                                   if hint else ""))
        if layout == "matrix":
            if len(s.get("quadrants", [])) != 4:
                errors.append("%s: matrix needs exactly 4 quadrants (TL, TR, BL, BR)"
                              % tag)
            hl = s.get("highlight")
            if hl is not None and hl not in (1, 2, 3, 4):
                errors.append("%s: matrix highlight must be 1–4" % tag)
        if layout in ("two_column", "comparison"):
            for side in ("left", "right"):
                col = s.get(side) or {}
                _lint_bullets("%s %s" % (tag, side), col.get("bullets", []),
                              COLUMN_WIDTH_PT, warnings)
        if layout == "image" and s.get("image"):
            if not os.path.exists(s["image"]):
                warnings.append("%s: image path not found: %s (placeholder will render)"
                                % (tag, s["image"]))

    from templates import THEMES, get_theme
    name = spec.get("template", "corporate")
    if name.strip().lower() not in THEMES:
        warnings.append("unknown template '%s' — will fall back to default" % name)
    lint_theme(get_theme(name), warnings)
    return errors, warnings


def dry_run(spec):
    """Lay the deck out in memory with the real renderer; return
    (errors, warnings) from its fit report. Skipped if python-pptx is missing."""
    try:
        import build_pptx
    except ImportError:
        return [], ["python-pptx not installed — skipped the dry-run fit check"]
    errors, warnings = [], []
    build_pptx.build(spec, io.BytesIO(), report=False)
    for line in build_pptx.fit_report():
        (errors if line.startswith("OVERFLOW") else warnings).append(line)
    return errors, warnings


def main(argv=None):
    ap = argparse.ArgumentParser(description="Validate a presentation spec")
    ap.add_argument("spec")
    ap.add_argument("--strict", action="store_true", help="treat warnings as errors")
    ap.add_argument("--no-render", action="store_true", help="skip the dry-run fit check")
    args = ap.parse_args(argv)
    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    errors, warnings = validate(spec)
    if not errors and not args.no_render:
        e2, w2 = dry_run(spec)
        errors += e2
        warnings += w2
    for w in warnings:
        print("WARN: " + w if not w.startswith(("WARN", "FIT")) else w)
    for e in errors:
        print("ERROR: " + e)
    notes = sum(1 for s in spec.get("slides", []) if isinstance(s, dict) and s.get("notes"))
    if errors or (args.strict and warnings):
        print("\n%d error(s), %d warning(s) — fix before rendering."
              % (len(errors), len(warnings)))
        return 1
    print("OK: spec is valid (%d slides, %d with speaker notes)%s"
          % (len(spec.get("slides", [])), notes,
             ", %d warning(s)" % len(warnings) if warnings else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
