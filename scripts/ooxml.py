#!/usr/bin/env python3
"""Unpack a .pptx to an editable XML tree and pack it back — the escape hatch
for edits python-pptx cannot reach.

A .pptx is a ZIP of XML parts. Most edits are easier through edit_pptx.py, but
some changes (theme colors, slide master tweaks, custom geometry, speaker-notes
internals) are only reachable by editing the raw Office Open XML. This mirrors
the unpack → edit XML → pack workflow used by Anthropic's official pptx skill.

Usage:
    python ooxml.py unpack deck.pptx unpacked/     # extract (XML pretty-printed)
    python ooxml.py pack   unpacked/ out.pptx       # rezip into a valid .pptx

Round-trip safe: unpack then pack reproduces a working deck. Edit the XML files
under the unpacked directory between the two steps.
"""
import argparse
import os
import sys
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def _pretty(xml_bytes):
    """Pretty-print XML for hand-editing; return original bytes on any failure."""
    try:
        from lxml import etree
        parser = etree.XMLParser(remove_blank_text=True)
        root = etree.fromstring(xml_bytes, parser)
        return etree.tostring(root, pretty_print=True, xml_declaration=True,
                              encoding="UTF-8")
    except Exception:  # noqa: BLE001 - non-XML parts pass through untouched
        return xml_bytes


def unpack(pptx_path, out_dir, pretty=True):
    if not zipfile.is_zipfile(pptx_path):
        raise SystemExit("not a .pptx (zip) file: %s" % pptx_path)
    os.makedirs(out_dir, exist_ok=True)
    count = 0
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            data = z.read(name)
            if pretty and name.lower().endswith((".xml", ".rels")):
                data = _pretty(data)
            target = os.path.join(out_dir, name)
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "wb") as f:
                f.write(data)
            count += 1
    return count


def pack(in_dir, pptx_path):
    if not os.path.isdir(in_dir):
        raise SystemExit("not a directory: %s" % in_dir)
    files = []
    for root, _, names in os.walk(in_dir):
        for n in names:
            full = os.path.join(root, n)
            arc = os.path.relpath(full, in_dir).replace(os.sep, "/")
            files.append((full, arc))
    # [Content_Types].xml must be the first entry for strict OPC readers.
    files.sort(key=lambda fa: (fa[1] != "[Content_Types].xml", fa[1]))
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for full, arc in files:
            z.write(full, arc)
    return len(files)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Unpack/pack a .pptx OOXML tree.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    u = sub.add_parser("unpack", help="extract a .pptx to an XML tree")
    u.add_argument("pptx")
    u.add_argument("out_dir")
    u.add_argument("--raw", action="store_true", help="do not pretty-print XML")
    p = sub.add_parser("pack", help="rezip an XML tree into a .pptx")
    p.add_argument("in_dir")
    p.add_argument("pptx")
    args = ap.parse_args(argv)

    if args.cmd == "unpack":
        n = unpack(args.pptx, args.out_dir, pretty=not args.raw)
        print("Unpacked %d part(s) to %s" % (n, args.out_dir))
    else:
        n = pack(args.in_dir, args.pptx)
        print("Packed %d part(s) into %s" % (n, args.pptx))
    return 0


if __name__ == "__main__":
    sys.exit(main())
