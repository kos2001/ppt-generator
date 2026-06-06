#!/usr/bin/env python3
"""Validate a presentation spec before rendering.

Catches the common mistakes (wrong layout name, missing required fields,
malformed chart/table data) and prints clear, line-oriented feedback so they
can be fixed without a failed render. Exits non-zero if any errors are found;
warnings alone do not fail.

Usage:
    python validate_spec.py spec.json
"""
import json
import sys

VALID_LAYOUTS = {
    "title", "section", "bullets", "content", "two_column", "comparison",
    "metrics", "quote", "image", "table", "chart", "closing", "diagram",
}

VALID_DIAGRAMS = {"process", "cycle", "hierarchy", "pyramid", "funnel", "timeline"}

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
}


def validate(spec):
    errors, warnings = [], []

    if not isinstance(spec, dict):
        return ["spec must be a JSON object"], []
    if "slides" not in spec or not isinstance(spec["slides"], list):
        errors.append("spec must have a 'slides' array")
        return errors, warnings
    if not spec["slides"]:
        errors.append("'slides' is empty — add at least one slide")

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
        req_keys, one_of_groups = REQUIRED[layout]
        for k in req_keys:
            if not s.get(k):
                errors.append("%s (%s): missing required field '%s'" % (tag, layout, k))
        for group in one_of_groups:
            if not any(s.get(k) for k in group):
                errors.append("%s (%s): needs at least one of %s"
                              % (tag, layout, group))

        if layout == "chart":
            ch = s.get("chart", {})
            if not ch.get("categories") and ch.get("type") != "pie":
                warnings.append("%s: chart has no 'categories'" % tag)
            if not ch.get("series"):
                errors.append("%s: chart has no 'series' data" % tag)
        if layout == "table":
            tb = s.get("table", {})
            if not tb.get("rows"):
                errors.append("%s: table has no 'rows'" % tag)
        if layout == "diagram":
            dg = s.get("diagram", {})
            dt = dg.get("type", "process")
            if dt not in VALID_DIAGRAMS:
                errors.append("%s: unknown diagram type '%s' (valid: %s)"
                              % (tag, dt, ", ".join(sorted(VALID_DIAGRAMS))))
            if not dg.get("nodes"):
                errors.append("%s: diagram has no 'nodes'" % tag)
        if layout == "metrics" and len(s.get("metrics", [])) > 4:
            warnings.append("%s: more than 4 metrics — only the first 4 render" % tag)
        if layout == "image" and s.get("image"):
            import os
            if not os.path.exists(s["image"]):
                warnings.append("%s: image path not found: %s (placeholder will render)"
                                % (tag, s["image"]))

    if "template" in spec:
        sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
        from templates import THEMES
        if spec["template"].strip().lower() not in THEMES:
            warnings.append("unknown template '%s' — will fall back to default"
                            % spec["template"])
    return errors, warnings


def main(argv=None):
    argv = argv or sys.argv[1:]
    if not argv:
        print("usage: python validate_spec.py spec.json", file=sys.stderr)
        return 2
    with open(argv[0], encoding="utf-8") as f:
        spec = json.load(f)
    errors, warnings = validate(spec)
    for w in warnings:
        print("WARN: " + w)
    for e in errors:
        print("ERROR: " + e)
    if errors:
        print("\n%d error(s) — fix before rendering." % len(errors))
        return 1
    print("OK: spec is valid (%d slides)%s"
          % (len(spec.get("slides", [])),
             ", %d warning(s)" % len(warnings) if warnings else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
