#!/usr/bin/env python
"""Build native, editable BLOCK DIAGRAMS in PowerPoint via COM (pywin32).

This is the COM counterpart to build_pptx.py's `diagram` layout. python-pptx
can draw diagrams but CANNOT open DRM/EDM-protected decks; this drives the
installed PowerPoint application instead, so it works on a corporate DRM deck
the signed-in user can already edit manually (PowerPoint decrypts in the
authorized session, the shapes are added, Save() re-encrypts). It does NOT
bypass DRM.

Every diagram is made of real, editable PowerPoint shapes — rounded rectangles,
arrows, connectors, ovals, and freeform trapezoids — never a flattened image.
Seven diagram types, mirroring the native catalog:

    process    horizontal steps joined by arrows (a pipeline)
    cycle      steps on a ring with connectors (a loop)
    hierarchy  a root box with child boxes below (org/tree)
    pyramid    stacked bands, narrow top -> wide base
    funnel     stacked bands, wide top -> narrow bottom
    timeline   milestones along a horizontal spine
    flowchart  arbitrary directed graph (branches/merges/loop-backs),
               auto-laid out LR or TD — give nodes ids + an `edges` list

Usage:
    # Showcase deck with all types (proves the capability end-to-end):
    python scripts/diagram_com.py --demo --out diagrams_demo.pptx

    # Flowchart onto an existing deck's slide (the DRM use case):
    python scripts/diagram_com.py --in deck.pptx --slide 4 --type flowchart \
        --direction TD --title "처리 흐름" \
        --nodes '[{"id":"a","title":"수집"},{"id":"b","title":"분석"}]' \
        --edges '[["a","b"]]'

    # Build a new deck from a JSON spec (list of diagram slides):
    python scripts/diagram_com.py --spec diagrams.json --out out.pptx

    # Add ONE diagram onto an existing deck's slide N (the DRM use case):
    python scripts/diagram_com.py --in deck.pptx --slide 4 \
        --type process --title "처리 흐름" \
        --nodes '[{"title":"수집"},{"title":"분석"},{"title":"출력"}]'

Spec shape (for --spec): {"title": "...", "template": "samsung",
  "slides": [{"type": "process", "title": "...", "nodes": [{"title","desc"}]}]}

Requirements: desktop PowerPoint (COM-registered) + `pip install pywin32`.
Run on the machine where PowerPoint (and, for DRM, the DRM client) lives.
"""
import argparse
import json
import math
import os
import sys

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from flowchart_layout import layer_graph, normalize_edges  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# --- Office COM enum constants (numeric, so no makepy/early-binding needed) ---
MSO_FALSE, MSO_TRUE = 0, -1
MSO_SHAPE_ROUNDED_RECT = 5
MSO_SHAPE_RIGHT_ARROW = 33
MSO_SHAPE_OVAL = 9
MSO_CONNECTOR_STRAIGHT = 1
MSO_TEXTBOX_H = 1
MSO_ARROWHEAD_TRIANGLE = 5
MSO_LINE_DASH = 4  # msoLineDash
MSO_SEGMENT_LINE = 0
MSO_EDITINGTYPE_CORNER = 1
PP_LAYOUT_BLANK = 12
PP_ALIGN_LEFT, PP_ALIGN_CENTER, PP_ALIGN_RIGHT = 1, 2, 3
MSO_ANCHOR_TOP, MSO_ANCHOR_MIDDLE = 1, 3
PP_SAVE_PPTX = 24

# Slide geometry in POINTS (PowerPoint's unit). 13.333 x 7.5 in = 16:9.
IN = 72.0
SLIDE_W, SLIDE_H = 13.333 * IN, 7.5 * IN
MARGIN = 0.92 * IN
BAR_H = 1.0 * IN

# samsung house palette (mirrors scripts/templates.py).
THEME = {
    "primary": "111111", "secondary": "1428A0", "accent": "1428A0",
    "surface": "F4F5F7", "on_primary": "FFFFFF", "text": "1A1A1A",
    "text_muted": "70747C",
    "heading_font": "Malgun Gothic", "body_font": "Malgun Gothic",
    "header_label": "Confidential",
}


def rgb(h):
    """'RRGGBB' hex -> Office BGR-packed int (RGB() macro order)."""
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


# --------------------------------------------------------------------------- #
# Primitive shape helpers
# --------------------------------------------------------------------------- #
def _box(slide, l, t, w, h, fill, line=None, shape=MSO_SHAPE_ROUNDED_RECT):
    sp = slide.Shapes.AddShape(shape, l, t, w, h)
    sp.Fill.ForeColor.RGB = rgb(fill)
    if line:
        sp.Line.ForeColor.RGB = rgb(line)
        sp.Line.Weight = 1.25
    else:
        sp.Line.Visible = MSO_FALSE
    sp.Shadow.Visible = MSO_FALSE
    return sp


def _label(sp, title, desc=None, title_size=15, desc_size=11,
           color="FFFFFF"):
    tf = sp.TextFrame
    tf.WordWrap = MSO_TRUE
    tf.VerticalAnchor = MSO_ANCHOR_MIDDLE
    tr = tf.TextRange
    tr.Text = (title or "") + (("\r" + desc) if desc else "")
    tr.Font.Name = THEME["heading_font"]
    tr.Font.Color.RGB = rgb(color)
    tr.ParagraphFormat.Alignment = PP_ALIGN_CENTER
    tr.Paragraphs(1).Font.Size = title_size
    tr.Paragraphs(1).Font.Bold = MSO_TRUE
    if desc:
        tr.Paragraphs(2).Font.Size = desc_size
        tr.Paragraphs(2).Font.Bold = MSO_FALSE


def _textbox(slide, l, t, w, h, text, size, color, *, bold=True,
             align=PP_ALIGN_LEFT, anchor=MSO_ANCHOR_TOP, italic=False):
    tb = slide.Shapes.AddTextbox(MSO_TEXTBOX_H, l, t, w, h)
    tf = tb.TextFrame
    tf.WordWrap = MSO_TRUE
    tf.VerticalAnchor = anchor
    tr = tf.TextRange
    tr.Text = text
    tr.Font.Name = THEME["body_font"]
    tr.Font.Size = size
    tr.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    tr.Font.Italic = MSO_TRUE if italic else MSO_FALSE
    tr.Font.Color.RGB = rgb(color)
    tr.ParagraphFormat.Alignment = align
    return tb


def _arrow_block(slide, l, t, w, h, fill):
    return _box(slide, l, t, w, h, fill, shape=MSO_SHAPE_RIGHT_ARROW)


def _connector(slide, x1, y1, x2, y2, color, weight=2.0, arrow=True, dash=False):
    c = slide.Shapes.AddConnector(MSO_CONNECTOR_STRAIGHT, x1, y1, x2, y2)
    c.Line.ForeColor.RGB = rgb(color)
    c.Line.Weight = weight
    if dash:
        c.Line.DashStyle = MSO_LINE_DASH
    if arrow:
        c.Line.EndArrowheadStyle = MSO_ARROWHEAD_TRIANGLE
    c.Shadow.Visible = MSO_FALSE
    return c


def _poly(slide, pts, fill):
    ff = slide.Shapes.BuildFreeform(MSO_EDITINGTYPE_CORNER,
                                    float(pts[0][0]), float(pts[0][1]))
    for x, y in pts[1:]:
        ff.AddNodes(MSO_SEGMENT_LINE, MSO_EDITINGTYPE_CORNER,
                    float(x), float(y))
    ff.AddNodes(MSO_SEGMENT_LINE, MSO_EDITINGTYPE_CORNER,
                float(pts[0][0]), float(pts[0][1]))
    sp = ff.ConvertToShape()
    sp.Fill.ForeColor.RGB = rgb(fill)
    sp.Line.Visible = MSO_FALSE
    sp.Shadow.Visible = MSO_FALSE
    return sp


def _band_colors(n):
    pair = [THEME["primary"], THEME["secondary"]]
    return [pair[i % 2] for i in range(n)]


def _area():
    """Drawing rectangle under the header bar: (left, top, w, h) in points."""
    top = 2.0 * IN
    return MARGIN, top, SLIDE_W - 2 * MARGIN, SLIDE_H - top - 0.7 * IN


# --------------------------------------------------------------------------- #
# The six diagram types
# --------------------------------------------------------------------------- #
def diag_process(slide, nodes):
    left, top, w, h = _area()
    n = max(len(nodes), 1)
    arrow_w = 0.5 * IN
    box_w = (w - arrow_w * (n - 1)) / n
    box_h = 1.7 * IN
    box_top = top + (h - box_h) / 2
    colors = _band_colors(n)
    for i, node in enumerate(nodes):
        x = left + i * (box_w + arrow_w)
        sp = _box(slide, x, box_top, box_w, box_h, colors[i])
        _label(sp, node.get("title", ""), node.get("desc"), 15, 11)
        if i < n - 1:
            ax = x + box_w
            ay = box_top + box_h / 2
            _arrow_block(slide, ax + 0.05 * IN, ay - 0.18 * IN,
                         arrow_w - 0.1 * IN, 0.36 * IN, THEME["accent"])


def diag_cycle(slide, nodes):
    left, top, w, h = _area()
    n = max(len(nodes), 1)
    cx, cy = left + w / 2, top + h / 2
    box_w, box_h = 2.3 * IN, 1.05 * IN
    radius = min(w, h) / 2 - box_h
    colors = _band_colors(n)
    centers = []
    for i in range(n):
        ang = -math.pi / 2 + i * 2 * math.pi / n
        centers.append((cx + radius * math.cos(ang),
                        cy + radius * math.sin(ang)))
    # Connect each box edge to the next box edge (where the center-to-center line
    # crosses each box) so arrows span the gap and touch both boxes, rather than
    # floating as stubs in the middle.
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
        _connector(slide, sx, sy, ex, ey, THEME["accent"], 2.0)
    for i, node in enumerate(nodes):
        px, py = centers[i]
        sp = _box(slide, px - box_w / 2, py - box_h / 2, box_w, box_h, colors[i])
        _label(sp, node.get("title", ""), node.get("desc"), 13, 10)


def diag_hierarchy(slide, nodes):
    left, top, w, h = _area()
    if not nodes:
        return
    root, children = nodes[0], nodes[1:]
    root_w, root_h = 3.0 * IN, 1.0 * IN
    root_x = left + (w - root_w) / 2
    rsp = _box(slide, root_x, top, root_w, root_h, THEME["primary"])
    _label(rsp, root.get("title", ""), root.get("desc"), 16, 11)
    if not children:
        return
    m = len(children)
    gap = 0.4 * IN
    child_w = min((w - gap * (m - 1)) / m, 3.2 * IN)
    total = child_w * m + gap * (m - 1)
    cstart = left + (w - total) / 2
    child_h, child_top = 1.1 * IN, top + 2.4 * IN
    rcx, rcy = root_x + root_w / 2, top + root_h
    for i, child in enumerate(children):
        cx = cstart + i * (child_w + gap)
        ccx = cx + child_w / 2
        _connector(slide, rcx, rcy + 0.1 * IN, ccx, child_top,
                   THEME["secondary"], 1.5, arrow=False)
        csp = _box(slide, cx, child_top, child_w, child_h,
                   THEME["surface"], line=THEME["secondary"])
        _label(csp, child.get("title", ""), child.get("desc"), 13, 10,
               color=THEME["text"])


def diag_pyramid(slide, nodes, funnel=False):
    left, top, w, h = _area()
    m = max(len(nodes), 1)
    base_hw = (w * 0.72) / 2
    cx = left + w / 2
    level_h = h / m
    colors = _band_colors(m)

    def hw(depth):
        frac = depth / m
        return base_hw * ((1 - frac) if funnel else frac)

    for i, node in enumerate(nodes):
        y_top, y_bot = top + i * level_h, top + (i + 1) * level_h
        hwt, hwb = hw(i), hw(i + 1)
        _poly(slide, [(cx - hwt, y_top), (cx + hwt, y_top),
                      (cx + hwb, y_bot), (cx - hwb, y_bot)], colors[i])
        lbl = node.get("title", "")
        if node.get("desc"):
            lbl += "  —  " + node["desc"]
        _textbox(slide, left, y_top, w, level_h, lbl, 14, THEME["on_primary"],
                 align=PP_ALIGN_CENTER, anchor=MSO_ANCHOR_MIDDLE)


def diag_funnel(slide, nodes):
    diag_pyramid(slide, nodes, funnel=True)


def diag_timeline(slide, nodes):
    left, top, w, h = _area()
    n = max(len(nodes), 1)
    line_y = top + h / 2
    pad = 0.4 * IN
    x0, x1 = left + pad, left + w - pad
    _box(slide, x0, line_y - 0.03 * IN, x1 - x0, 0.06 * IN, THEME["secondary"])
    step = (x1 - x0) / (n - 1) if n > 1 else 0
    colors = _band_colors(n)
    box_w = 2.4 * IN
    for i, node in enumerate(nodes):
        mx = x0 if n == 1 else x0 + i * step
        r = 0.13 * IN
        _box(slide, mx - r, line_y - r, 2 * r, 2 * r, colors[i],
             shape=MSO_SHAPE_OVAL)
        above = (i % 2 == 0)
        by = (line_y - 1.55 * IN) if above else (line_y + 0.35 * IN)
        bx = max(left, min(mx - box_w / 2, left + w - box_w))
        sp = _box(slide, bx, by, box_w, 1.2 * IN, THEME["surface"],
                  line=THEME["secondary"])
        _label(sp, node.get("title", ""), node.get("desc"), 13, 10,
               color=THEME["text"])
        sy = (by + 1.2 * IN) if above else by
        _connector(slide, mx, line_y, mx, sy, THEME["secondary"], 1.25,
                   arrow=False)


def diag_flowchart(slide, spec):
    """Auto-laid-out flowchart (nodes + edges) as native COM shapes.

    Shares the layered layout + back-edge detection with the python-pptx
    renderer (flowchart_layout.layer_graph); feedback edges draw as muted
    return paths so cycles stay compact. spec = {nodes, edges, direction}.
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

    left, top, w, h = _area()
    gap_x, gap_y = 0.55 * IN, 0.4 * IN
    if td:
        box_w = min(2.6 * IN, (w - gap_x * (max_in_layer - 1)) / max_in_layer)
        box_h = min(1.05 * IN, (h - gap_y * (n_layers - 1)) / n_layers)
    else:
        box_w = min(2.3 * IN, (w - gap_x * (n_layers - 1)) / n_layers)
        box_h = min(1.2 * IN, (h - gap_y * (max_in_layer - 1)) / max_in_layer)

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

    clearance = 0.22 * IN
    for a, b, label in edges:
        if a == b:
            continue
        ax, ay = rects[a]
        bx, by = rects[b]
        if (a, b) not in feedback:
            # Forward step: straight solid arrow between adjacent box edges.
            if td:
                p1, p2 = (ax + box_w / 2, ay + box_h), (bx + box_w / 2, by)
            else:
                p1, p2 = (ax + box_w, ay + box_h / 2), (bx, by + box_h / 2)
            _connector(slide, p1[0], p1[1], p2[0], p2[1], THEME["accent"], 1.75)
            lx, ly = (p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2
        else:
            # Feedback / loop-back: DASHED path routed clear of the boxes, arrow
            # on the final leg, so the return direction is unmistakable.
            col = THEME["secondary"]
            if not td:  # LR -> out the right of the source, up, over, down
                yline = top - clearance
                x_out = min(ax + box_w + clearance, left + w)
                xb, ya = bx + box_w / 2, ay + box_h / 2
                _connector(slide, ax + box_w, ya, x_out, ya, col, 1.5, arrow=False, dash=True)
                _connector(slide, x_out, ya, x_out, yline, col, 1.5, arrow=False, dash=True)
                _connector(slide, x_out, yline, xb, yline, col, 1.5, arrow=False, dash=True)
                _connector(slide, xb, yline, xb, by, col, 1.5, arrow=True, dash=True)
                lx, ly = (x_out + xb) / 2, yline - 0.02 * IN
            else:       # TD -> route out the left side
                xline = left - clearance if left - clearance > 0.1 * IN else 0.1 * IN
                ya, yb = ay + box_h / 2, by + box_h / 2
                _connector(slide, ax, ya, xline, ya, col, 1.5, arrow=False, dash=True)
                _connector(slide, xline, ya, xline, yb, col, 1.5, arrow=False, dash=True)
                _connector(slide, xline, yb, bx, yb, col, 1.5, arrow=True, dash=True)
                lx, ly = xline, (ya + yb) / 2
        if label:
            _textbox(slide, lx - 0.7 * IN, ly - 0.22 * IN, 1.4 * IN, 0.3 * IN,
                     str(label), 10, THEME["text_muted"], bold=False,
                     align=PP_ALIGN_CENTER, anchor=MSO_ANCHOR_MIDDLE)

    colors = _band_colors(n_layers)
    for n in ids:
        x, y = rects[n]
        sp = _box(slide, x, y, box_w, box_h, colors[rank[n]])
        nd = nodemap[n]
        _label(sp, nd.get("title", n), nd.get("desc"), 13, 10,
               color=(THEME["on_primary"]))


DIAGRAMS = {
    "process": diag_process, "cycle": diag_cycle, "hierarchy": diag_hierarchy,
    "pyramid": diag_pyramid, "funnel": diag_funnel, "timeline": diag_timeline,
    "flowchart": diag_flowchart,
}


# --------------------------------------------------------------------------- #
# Chrome + slide assembly
# --------------------------------------------------------------------------- #
def _header(slide, title):
    """samsung header bar: single-line white title (no eyebrow) + Confidential."""
    _box(slide, 0, 0, SLIDE_W, BAR_H, THEME["primary"], shape=1)  # rectangle
    _textbox(slide, MARGIN, 0, SLIDE_W - 2 * MARGIN - 3.0 * IN, BAR_H, title,
             24, THEME["on_primary"], anchor=MSO_ANCHOR_MIDDLE)
    if THEME.get("header_label"):
        _textbox(slide, SLIDE_W - MARGIN - 3.0 * IN, 0.2 * IN, 3.0 * IN,
                 0.6 * IN, THEME["header_label"], 16, THEME["on_primary"],
                 align=PP_ALIGN_RIGHT)


def _add_blank_slide(pres, index):
    return pres.Slides.Add(index, PP_LAYOUT_BLANK)


def draw_diagram(slide, spec):
    """Draw the diagram described by `spec` ({type, nodes, [edges, direction]})."""
    dtype = spec.get("type", "process")
    fn = DIAGRAMS.get(dtype)
    if fn is None:
        raise SystemExit("unknown diagram type %r (choose: %s)"
                         % (dtype, ", ".join(DIAGRAMS)))
    if dtype == "flowchart":  # needs edges + direction, so takes the whole spec
        diag_flowchart(slide, spec)
    else:
        fn(slide, spec.get("nodes", []))


def _blank_background(slide):
    _box(slide, 0, 0, SLIDE_W, SLIDE_H, "FFFFFF", shape=1)


# --------------------------------------------------------------------------- #
# Demo content
# --------------------------------------------------------------------------- #
DEMO = [
    {"type": "process", "title": "Process · 처리 파이프라인", "nodes": [
        {"title": "수집", "desc": "데이터 인입"},
        {"title": "정제", "desc": "전처리·검증"},
        {"title": "분석", "desc": "모델 추론"},
        {"title": "배포", "desc": "결과 전달"}]},
    {"type": "cycle", "title": "Cycle · 반복 개선 루프", "nodes": [
        {"title": "계획", "desc": "Plan"}, {"title": "실행", "desc": "Do"},
        {"title": "점검", "desc": "Check"}, {"title": "개선", "desc": "Act"}]},
    {"type": "hierarchy", "title": "Hierarchy · 조직 구조", "nodes": [
        {"title": "플랫폼 본부"},
        {"title": "데이터", "desc": "수집·저장"},
        {"title": "모델", "desc": "학습·서빙"},
        {"title": "서비스", "desc": "API·UX"}]},
    {"type": "pyramid", "title": "Pyramid · 가치 계층", "nodes": [
        {"title": "비전"}, {"title": "전략"}, {"title": "실행 과제"},
        {"title": "기반 데이터"}]},
    {"type": "funnel", "title": "Funnel · 전환 깔때기", "nodes": [
        {"title": "방문"}, {"title": "가입"}, {"title": "활성"},
        {"title": "결제"}]},
    {"type": "timeline", "title": "Timeline · 로드맵", "nodes": [
        {"title": "1분기", "desc": "PoC"}, {"title": "2분기", "desc": "파일럿"},
        {"title": "3분기", "desc": "정식 출시"}, {"title": "4분기", "desc": "확장"}]},
    {"type": "flowchart", "title": "Flowchart · 분기/루프백 (LR)", "direction": "LR",
     "nodes": [
        {"id": "a", "title": "이슈 등록"}, {"id": "b", "title": "분석"},
        {"id": "c", "title": "코드 생성"}, {"id": "d", "title": "리뷰"},
        {"id": "e", "title": "병합"}, {"id": "f", "title": "수정 반영"}],
     "edges": [["a", "b"], ["b", "c"], ["c", "d"],
               {"from": "d", "to": "e", "label": "통과"},
               {"from": "d", "to": "f", "label": "반려"},
               {"from": "f", "to": "c", "label": "재시도"}]},
]


def build_new(slides_spec, out_path):
    pp = win32com.client.Dispatch("PowerPoint.Application")
    pres = pp.Presentations.Add(WithWindow=MSO_FALSE)
    pres.PageSetup.SlideWidth = SLIDE_W
    pres.PageSetup.SlideHeight = SLIDE_H
    try:
        for i, sp in enumerate(slides_spec, start=1):
            slide = _add_blank_slide(pres, i)
            _blank_background(slide)
            _header(slide, sp.get("title", ""))
            draw_diagram(slide, sp)
        pres.SaveAs(os.path.abspath(out_path))
        n = pres.Slides.Count
    finally:
        pres.Close()
        pp.Quit()
    print("Built %d diagram slide(s) -> %s" % (n, out_path))


def add_to_existing(in_path, slide_no, spec, title, out_path):
    pp = win32com.client.Dispatch("PowerPoint.Application")
    pres = pp.Presentations.Open(os.path.abspath(in_path), ReadOnly=MSO_FALSE,
                                 Untitled=MSO_FALSE, WithWindow=MSO_FALSE)
    try:
        slide = pres.Slides(slide_no)
        if title:
            _header(slide, title)
        draw_diagram(slide, spec)
        if out_path:
            pres.SaveAs(os.path.abspath(out_path))
        else:
            pres.Save()
    finally:
        pres.Close()
        pp.Quit()
    print("Added %s diagram to slide %d of %s"
          % (spec.get("type"), slide_no, out_path or in_path))


def main():
    ap = argparse.ArgumentParser(description="Build block diagrams via PowerPoint COM.")
    ap.add_argument("--demo", action="store_true", help="showcase all six types")
    ap.add_argument("--spec", help="JSON spec with a 'slides' list of diagrams")
    ap.add_argument("--in", dest="in_path", help="existing deck to add a diagram to")
    ap.add_argument("--slide", type=int, help="target slide number (with --in)")
    ap.add_argument("--type", help="diagram type (with --in)")
    ap.add_argument("--title", help="header title (with --in)")
    ap.add_argument("--nodes", help="JSON list of {title,desc} (with --in)")
    ap.add_argument("--edges", help="JSON list of [from,to] or {from,to,label} "
                    "(flowchart, with --in)")
    ap.add_argument("--direction", default="LR", help="flowchart direction: "
                    "LR (default) or TD")
    ap.add_argument("-o", "--out", help="output .pptx path")
    args = ap.parse_args()

    if args.in_path:
        if not (args.slide and args.type):
            ap.error("--in requires --slide and --type")
        spec = {
            "type": args.type,
            "nodes": json.loads(args.nodes) if args.nodes else [],
            "edges": json.loads(args.edges) if args.edges else [],
            "direction": args.direction,
        }
        add_to_existing(args.in_path, args.slide, spec, args.title, args.out)
        return 0

    if args.demo:
        slides_spec = DEMO
    elif args.spec:
        with open(args.spec, encoding="utf-8") as f:
            slides_spec = json.load(f).get("slides", [])
    else:
        ap.error("choose one of: --demo, --spec, or --in")
    out = args.out or "diagrams.pptx"
    build_new(slides_spec, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
