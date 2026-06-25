#!/usr/bin/env python
"""Extract text and speaker notes from a PPTX by automating desktop PowerPoint
via COM (pywin32).

Unlike extract_pptx.py (which parses the file directly and cannot open
DRM-protected files), this drives the installed PowerPoint application. If the
current Windows user is authorized to open the file in PowerPoint — including
files unlocked by an installed DRM client — PowerPoint decrypts it and this
script reads the rendered content through the app. It does NOT bypass DRM; it
relies on PowerPoint's own authorized access.

Requirements:
    - Desktop Microsoft PowerPoint installed (COM-registered)
    - pip install pywin32
    - For DRM files: the corporate DRM client installed and the user signed in

Usage:
    python scripts/extract_pptx_com.py <path-to.pptx> [--json]
"""
import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

try:
    import win32com.client
except ImportError:
    print("pywin32 is required: pip install pywin32", file=sys.stderr)
    sys.exit(1)


def extract(path):
    path = os.path.abspath(path)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    # Open read-only, no window, no links update, no title prompt.
    pres = powerpoint.Presentations.Open(
        path, ReadOnly=True, Untitled=False, WithWindow=False
    )
    slides = []
    try:
        for idx, slide in enumerate(pres.Slides, start=1):
            texts = []
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    t = shape.TextFrame.TextRange.Text.strip()
                    if t:
                        texts.append(t)
            notes = ""
            try:
                ns = slide.NotesPage
                for shape in ns.Shapes:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        cand = shape.TextFrame.TextRange.Text.strip()
                        # The notes placeholder; skip the slide-image thumbnail.
                        if cand:
                            notes = cand
            except Exception:  # noqa: BLE001
                pass
            slides.append({"slide": idx, "texts": texts, "notes": notes})
    finally:
        pres.Close()
        powerpoint.Quit()
    return slides


def to_markdown(slides):
    out = []
    for s in slides:
        out.append(f"## Slide {s['slide']}")
        out.extend(s["texts"])
        if s["notes"]:
            out.append(f"\n> **Notes:** {s['notes']}")
        out.append("")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(
        description="Extract PPTX content via PowerPoint COM automation."
    )
    ap.add_argument("path", help="path to .pptx file")
    ap.add_argument("--json", action="store_true", help="output JSON")
    args = ap.parse_args()

    try:
        slides = extract(args.path)
    except Exception as e:  # noqa: BLE001
        print(f"Failed to open via PowerPoint: {e}", file=sys.stderr)
        print(
            "Ensure desktop PowerPoint is installed and (for DRM files) that "
            "you are signed into the DRM client with rights to this document.",
            file=sys.stderr,
        )
        sys.exit(1)

    if args.json:
        print(json.dumps(slides, ensure_ascii=False, indent=2))
    else:
        print(to_markdown(slides))


if __name__ == "__main__":
    main()
