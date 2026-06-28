#!/usr/bin/env python3
"""Validate a presentation spec (for build_pptx.py) or an edit spec (for
edit_pptx.py) before running it.

Catches the common mistakes — wrong layout/op name, missing required fields,
malformed chart/table data, bad fit modes — and prints clear, line-oriented
feedback so they can be fixed without a failed run. Exits non-zero if any
errors are found; warnings alone do not fail.

The schema is auto-detected: a top-level 'operations' array is an edit spec, a
'slides' array is a presentation spec. Force it with --edit or --deck.

Usage:
    python validate_spec.py spec.json            # auto-detect
    python validate_spec.py --edit edits.json     # force edit-spec checks
"""
import json
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

VALID_LAYOUTS = {
    "title", "section", "bullets", "content", "two_column", "comparison",
    "metrics", "quote", "image", "table", "chart", "closing", "diagram",
}

VALID_DIAGRAMS = {"process", "cycle", "hierarchy", "pyramid", "funnel",
                  "timeline", "flowchart"}

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


# --------------------------------------------------------------------------- #
# Edit spec (for edit_pptx.py)
# --------------------------------------------------------------------------- #
VALID_OPS = {
    "replace_text", "set_text", "set_table_cell", "set_chart_data",
    "fit_image", "replace_image", "add_image", "add_textbox",
    "add_chrome", "renumber_pages",
    "delete_shape", "duplicate_slide", "delete_slide", "move_slide",
}
VALID_FITS = {"contain", "cover", "stretch"}


def _is_int(v):
    return isinstance(v, int) and not isinstance(v, bool)


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def validate_edits(spec):
    import os
    errors, warnings = [], []

    if not isinstance(spec, dict):
        return ["edit spec must be a JSON object"], []
    if "operations" not in spec or not isinstance(spec["operations"], list):
        return ["edit spec must have an 'operations' array"], []
    if not spec["operations"]:
        warnings.append("'operations' is empty — nothing to do")
    if "source" not in spec:
        warnings.append("no 'source' — must be supplied with -i on the command line")
    if "output" not in spec:
        warnings.append("no 'output' — must be supplied with -o on the command line")

    def need(op, tag, *fields, kind=None, check=None):
        for f in fields:
            if f not in op:
                errors.append("%s: missing required field '%s'" % (tag, f))
            elif kind and not kind(op[f]):
                errors.append("%s: field '%s' has wrong type" % (tag, f))

    for i, op in enumerate(spec["operations"]):
        tag = "operation %d" % (i + 1)
        if not isinstance(op, dict):
            errors.append("%s: must be an object" % tag)
            continue
        kind = op.get("op")
        if kind not in VALID_OPS:
            errors.append("%s: unknown op %r (valid: %s)"
                          % (tag, kind, ", ".join(sorted(VALID_OPS))))
            continue
        tag = "%s (%s)" % (tag, kind)

        if kind == "replace_text":
            need(op, tag, "find", kind=lambda v: isinstance(v, str))
            if "slides" in op and not (isinstance(op["slides"], list)
                                       and all(_is_int(x) for x in op["slides"])):
                errors.append("%s: 'slides' must be a list of integers" % tag)
        elif kind == "set_text":
            need(op, tag, "slide", "shape", kind=_is_int)
            need(op, tag, "text", kind=lambda v: isinstance(v, str))
        elif kind == "set_table_cell":
            need(op, tag, "slide", "shape", "row", "col", kind=_is_int)
            need(op, tag, "text", kind=lambda v: isinstance(v, str))
        elif kind == "set_chart_data":
            need(op, tag, "slide", "shape", kind=_is_int)
            series = op.get("series")
            if not isinstance(series, list) or not series:
                errors.append("%s: 'series' must be a non-empty list" % tag)
            else:
                for j, ser in enumerate(series):
                    if not isinstance(ser, dict) or "values" not in ser:
                        errors.append("%s: series %d needs 'values'" % (tag, j + 1))
                    elif not (isinstance(ser["values"], list)
                              and all(_is_num(x) for x in ser["values"])):
                        errors.append("%s: series %d 'values' must be numbers"
                                      % (tag, j + 1))
            if "categories" in op and not isinstance(op["categories"], list):
                errors.append("%s: 'categories' must be a list" % tag)
        elif kind == "fit_image":
            need(op, tag, "slide", "shape", kind=_is_int)
            if "fit" in op and op["fit"] not in VALID_FITS:
                errors.append("%s: 'fit' must be one of %s"
                              % (tag, ", ".join(sorted(VALID_FITS))))
            if "into_shape" in op and not _is_int(op["into_shape"]):
                errors.append("%s: 'into_shape' must be an integer" % tag)
            if "template" in op and not isinstance(op["template"], str):
                errors.append("%s: 'template' must be a string" % tag)
        elif kind == "replace_image":
            need(op, tag, "slide", "shape", kind=_is_int)
            need(op, tag, "image", kind=lambda v: isinstance(v, str))
            if op.get("image") and not os.path.exists(op["image"]):
                warnings.append("%s: image not found: %s" % (tag, op["image"]))
            if "fit" in op and op["fit"] not in VALID_FITS:
                errors.append("%s: 'fit' must be one of %s"
                              % (tag, ", ".join(sorted(VALID_FITS))))
        elif kind == "add_image":
            need(op, tag, "slide", kind=_is_int)
            need(op, tag, "image", kind=lambda v: isinstance(v, str))
            if op.get("image") and not os.path.exists(op["image"]):
                warnings.append("%s: image not found: %s" % (tag, op["image"]))
            if "into_shape" in op and not _is_int(op["into_shape"]):
                errors.append("%s: 'into_shape' must be an integer" % tag)
            if "template" in op and not isinstance(op["template"], str):
                errors.append("%s: 'template' must be a string" % tag)
            if "fit" in op and op["fit"] not in VALID_FITS:
                errors.append("%s: 'fit' must be one of %s"
                              % (tag, ", ".join(sorted(VALID_FITS))))
            if ("into_shape" not in op and "template" not in op
                    and (("width_in" in op) ^ ("height_in" in op))):
                warnings.append("%s: give both width_in and height_in for a fixed "
                                "box, or neither for native size" % tag)
        elif kind == "add_textbox":
            need(op, tag, "slide", kind=_is_int)
            need(op, tag, "text", kind=lambda v: isinstance(v, str))
            if "align" in op and op["align"] not in ("left", "center", "right"):
                errors.append("%s: 'align' must be left, center, or right" % tag)
        elif kind == "add_chrome":
            need(op, tag, "slide", kind=_is_int)
            if "template" in op and not isinstance(op["template"], str):
                errors.append("%s: 'template' must be a string" % tag)
            for f in ("slide_no", "total"):
                if f in op and not _is_int(op[f]):
                    errors.append("%s: '%s' must be an integer" % (tag, f))
            if "page_number" in op and not isinstance(op["page_number"], bool):
                errors.append("%s: 'page_number' must be true or false" % tag)
        elif kind == "renumber_pages":
            if "format" in op and not isinstance(op["format"], str):
                errors.append("%s: 'format' must be a string" % tag)
        elif kind == "delete_shape":
            need(op, tag, "slide", "shape", kind=_is_int)
        elif kind == "duplicate_slide":
            need(op, tag, "slide", kind=_is_int)
            if "count" in op and not (_is_int(op["count"]) and op["count"] >= 1):
                errors.append("%s: 'count' must be an integer >= 1" % tag)
            if "at" in op and not _is_int(op["at"]):
                errors.append("%s: 'at' must be an integer" % tag)
        elif kind == "delete_slide":
            need(op, tag, "slide", kind=_is_int)
        elif kind == "move_slide":
            need(op, tag, "from", "to", kind=_is_int)

    return errors, warnings


def main(argv=None):
    argv = argv or sys.argv[1:]
    flags = [a for a in argv if a.startswith("--")]
    paths = [a for a in argv if not a.startswith("--")]
    if not paths:
        print("usage: python validate_spec.py [--edit|--deck] spec.json",
              file=sys.stderr)
        return 2
    with open(paths[0], encoding="utf-8") as f:
        spec = json.load(f)

    # Pick which schema to validate against: explicit flag wins, else auto-detect.
    if "--edit" in flags:
        is_edit = True
    elif "--deck" in flags:
        is_edit = False
    elif isinstance(spec, dict) and "operations" in spec and "slides" not in spec:
        is_edit = True
    else:
        is_edit = False

    if is_edit:
        errors, warnings = validate_edits(spec)
        ok_msg = "OK: edit spec is valid (%d operation(s))" % len(
            spec.get("operations", []) if isinstance(spec, dict) else [])
    else:
        errors, warnings = validate(spec)
        ok_msg = "OK: spec is valid (%d slides)" % len(
            spec.get("slides", []) if isinstance(spec, dict) else [])

    for w in warnings:
        print("WARN: " + w)
    for e in errors:
        print("ERROR: " + e)
    if errors:
        print("\n%d error(s) — fix before %s."
              % (len(errors), "editing" if is_edit else "rendering"))
        return 1
    print(ok_msg + (", %d warning(s)" % len(warnings) if warnings else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
