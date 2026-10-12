#!/usr/bin/env python3
"""Maintainer tool: convert Phosphor "fill" SVG icons into the bundled icon set.

The renderer draws icons as native, recolorable PowerPoint shapes (a:custGeom),
which only support move/line/cubic/close. This tool parses each SVG path,
converts every command (including arcs and quadratics) to absolute cubic
Béziers, and writes assets/icons/phosphor-fill.json. Run it only when
changing the curated icon list; the skill itself never touches SVG.

Usage:
    git clone --depth 1 --filter=blob:none --sparse \\
        https://github.com/phosphor-icons/core.git && \\
        (cd core && git sparse-checkout set assets/fill)
    python build_icon_set.py core/assets/fill -o ../assets/icons/phosphor-fill.json

Icons: Phosphor Icons (https://phosphoricons.com), MIT License — see
assets/icons/LICENSE-phosphor.txt.
"""
import argparse
import json
import math
import os
import re
import sys

sys.dont_write_bytecode = True

# Curated for business, product and engineering decks. Names are Phosphor
# names without the "-fill" suffix; missing names are skipped with a warning.
CURATED = """
address-book airplane alarm archive arrow-circle-right arrows-clockwise
arrows-left-right arrow-right arrow-up-right article atom bank barcode
battery-charging bell binoculars bookmark book-open books brain briefcase
broadcast bug buildings calendar calendar-check camera car chart-bar
chart-line chart-line-up chart-pie chart-donut chat-circle chats check-circle
checks clipboard-text clock cloud cloud-arrow-up code coffee coins compass
cpu credit-card crown cube currency-dollar currency-krw database desktop
device-mobile diamond door-open download envelope eye factory file file-text
files fingerprint fire flag flask flow-arrow folder funnel game-controller
gauge gear gift git-branch globe graduation-cap graph hand-coins handshake
hard-drives hash headset heart heartbeat house identification-card image
infinity info key keyboard lightbulb lightning link list-checks lock
lock-open magnifying-glass map-pin map-trifold medal megaphone microphone
money network newspaper note package paint-brush paper-plane-tilt path
pencil phone piggy-bank plant play plug plugs-connected presentation-chart
printer puzzle-piece question receipt recycle robot rocket rocket-launch
scales seal-check share-network shield shield-check shopping-cart sparkle
speedometer squares-four stack star storefront strategy sun target terminal
thumbs-up timer translate trash tree-structure trend-up trophy truck
umbrella user user-circle user-gear users users-three video-camera wallet
warning warning-circle wifi-high wrench x-circle
""".split()

_TOKEN = re.compile(r"[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?")


def _tokens(d):
    return _TOKEN.findall(d)


def _arc_to_cubics(x1, y1, rx, ry, phi, large, sweep, x2, y2):
    """SVG endpoint arc -> list of cubic segments (c1x,c1y,c2x,c2y,x,y)."""
    if (x1, y1) == (x2, y2):
        return []
    if rx == 0 or ry == 0:
        return [(x1, y1, x2, y2, x2, y2)]
    rx, ry = abs(rx), abs(ry)
    cphi, sphi = math.cos(math.radians(phi)), math.sin(math.radians(phi))
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p = cphi * dx + sphi * dy
    y1p = -sphi * dx + cphi * dy
    lam = (x1p ** 2) / rx ** 2 + (y1p ** 2) / ry ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    coef = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        coef = -coef
    cxp, cyp = coef * rx * y1p / ry, -coef * ry * x1p / rx
    cx = cphi * cxp - sphi * cyp + (x1 + x2) / 2
    cy = sphi * cxp + cphi * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(1, int(math.ceil(abs(dt) / (math.pi / 2) - 1e-9)))
    step = dt / n
    k = 4 / 3 * math.tan(step / 4)
    out = []
    for i in range(n):
        a1, a2 = t1 + i * step, t1 + (i + 1) * step
        c1, s1, c2, s2 = math.cos(a1), math.sin(a1), math.cos(a2), math.sin(a2)
        pts = [(c1 - k * s1, s1 + k * c1), (c2 + k * s2, s2 - k * c2), (c2, s2)]
        seg = []
        for ex, ey in pts:
            ex, ey = ex * rx, ey * ry
            seg += [cphi * ex - sphi * ey + cx, sphi * ex + cphi * ey + cy]
        out.append(tuple(seg))
    return out


def parse_path(d):
    """SVG path data -> list of subpaths; each a list of ops:
    ("M", x, y) / ("L", x, y) / ("C", x1, y1, x2, y2, x, y) / ("Z",)."""
    toks = _tokens(d)
    i, cmd = 0, None
    x = y = sx = sy = 0.0
    last_c2 = last_q = None
    ops = []

    def num():
        nonlocal i
        v = float(toks[i])
        i += 1
        return v

    def flag():
        # Arc flags may be packed ("01" -> 0, 1); the tokenizer splits digits
        # only when separated, so peel one character at a time.
        nonlocal i
        t = toks[i]
        if len(t) > 1 and t[0] in "01" and not t.startswith(("0.", "1.")):
            toks[i] = t[1:]
            return float(t[0])
        i += 1
        return float(t)

    while i < len(toks):
        if re.match(r"[A-Za-z]", toks[i]):
            cmd = toks[i]
            i += 1
            if cmd in "Zz":
                ops.append(("Z",))
                x, y = sx, sy
                last_c2 = last_q = None
                continue
        rel = cmd.islower()
        c = cmd.upper()
        ox, oy = (x, y) if rel else (0.0, 0.0)
        if c == "M":
            x, y = num() + ox, num() + oy
            sx, sy = x, y
            ops.append(("M", x, y))
            cmd = "l" if rel else "L"       # implicit lineto after moveto
            last_c2 = last_q = None
        elif c == "L":
            x, y = num() + ox, num() + oy
            ops.append(("L", x, y))
            last_c2 = last_q = None
        elif c == "H":
            x = num() + (x if rel else 0.0)
            ops.append(("L", x, y))
            last_c2 = last_q = None
        elif c == "V":
            y = num() + (y if rel else 0.0)
            ops.append(("L", x, y))
            last_c2 = last_q = None
        elif c == "C":
            x1, y1, x2, y2 = num() + ox, num() + oy, num() + ox, num() + oy
            x, y = num() + ox, num() + oy
            ops.append(("C", x1, y1, x2, y2, x, y))
            last_c2, last_q = (x2, y2), None
        elif c == "S":
            x1, y1 = (2 * x - last_c2[0], 2 * y - last_c2[1]) if last_c2 else (x, y)
            x2, y2 = num() + ox, num() + oy
            x, y = num() + ox, num() + oy
            ops.append(("C", x1, y1, x2, y2, x, y))
            last_c2, last_q = (x2, y2), None
        elif c in "QT":
            if c == "Q":
                qx, qy = num() + ox, num() + oy
            else:
                qx, qy = (2 * x - last_q[0], 2 * y - last_q[1]) if last_q else (x, y)
            ex, ey = num() + ox, num() + oy
            ops.append(("C", x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y),
                        ex + 2 / 3 * (qx - ex), ey + 2 / 3 * (qy - ey), ex, ey))
            x, y = ex, ey
            last_q, last_c2 = (qx, qy), None
        elif c == "A":
            rx, ry, phi = num(), num(), num()
            large, sweep = flag(), flag()
            ex, ey = num() + ox, num() + oy
            for seg in _arc_to_cubics(x, y, rx, ry, phi, int(large), int(sweep), ex, ey):
                ops.append(("C",) + seg)
            x, y = ex, ey
            last_c2 = last_q = None
        else:
            raise ValueError("unsupported path command %r" % cmd)
    return ops


def icon_ops(svg_text):
    paths = re.findall(r'<path[^>]*\sd="([^"]+)"', svg_text)
    ops = []
    for d in paths:
        ops += parse_path(d)
    return [[o[0]] + [round(v, 2) for v in o[1:]] for o in ops]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("fill_dir", help="phosphor core assets/fill directory")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args(argv)
    icons, missing = {}, []
    for name in CURATED:
        path = os.path.join(args.fill_dir, name + "-fill.svg")
        if not os.path.exists(path):
            missing.append(name)
            continue
        with open(path, encoding="utf-8") as f:
            svg = f.read()
        vb = re.search(r'viewBox="0 0 (\d+) (\d+)"', svg)
        icons[name] = {"size": int(vb.group(1)) if vb else 256, "ops": icon_ops(svg)}
    out = {"source": "Phosphor Icons (fill weight), MIT License, "
                     "https://github.com/phosphor-icons/core",
           "icons": icons}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, separators=(",", ":"))
    print("Wrote %d icons -> %s" % (len(icons), args.out))
    if missing:
        print("WARN: not found (skipped): %s" % ", ".join(missing), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
