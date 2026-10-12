#!/usr/bin/env python3
"""Pure-Python layered layout for flowchart diagrams (no pptx/COM deps).

Shared by the two renderers — build_pptx.py (python-pptx) and diagram_com.py
(PowerPoint COM) — so both lay an arbitrary directed graph out the same way and
only differ in how they draw the resulting boxes/connectors.

The hard parts live here:
  * feedback (back) edge detection via DFS, so cycles/loop-backs don't stretch
    the layout — back edges are excluded from ranking and flagged for the
    renderer to draw as return paths.
  * longest-path layering over the remaining (forward) edges.

`layer_graph` is unit-agnostic: it returns ranks + layer groupings + the
feedback-edge set. Each renderer turns ranks into coordinates for its own units
(EMU for python-pptx, points for COM) and chosen direction (LR or TD).
"""
from collections import deque, defaultdict


def normalize_edges(raw):
    """Accept [from, to] pairs or {from, to, label} dicts -> [(from, to, label)]."""
    out = []
    for e in raw or []:
        if isinstance(e, dict):
            out.append((str(e.get("from")), str(e.get("to")), e.get("label")))
        else:
            out.append((str(e[0]), str(e[1]), e[2] if len(e) > 2 else None))
    return out


def layer_graph(ids, edges):
    """Layer a directed graph for drawing.

    ids:   ordered list of node ids (str).
    edges: list of (from, to, label) with ids that exist in `ids`.
    Returns (rank, layers, feedback):
      rank[id]      -> integer layer (0 = first column/row)
      layers[r]     -> list of ids in layer r, in input order
      feedback      -> set of (from, to) edges detected as back edges
    """
    ids = [n for n in ids]
    idset = set(ids)
    edges = [(a, b, l) for (a, b, l) in edges if a in idset and b in idset]

    adj = defaultdict(list)
    for a, b, _l in edges:
        if a != b:  # ignore self-loops for traversal
            adj[a].append(b)

    # DFS back-edge detection: an edge to a node currently on the stack (GRAY).
    WHITE, GRAY, BLACK = 0, 1, 2
    color = {n: WHITE for n in ids}
    feedback = set()
    for src in ids:
        if color[src] != WHITE:
            continue
        color[src] = GRAY
        stack = [(src, iter(adj[src]))]
        while stack:
            node, it = stack[-1]
            advanced = False
            for nb in it:
                if color[nb] == GRAY:
                    feedback.add((node, nb))
                elif color[nb] == WHITE:
                    color[nb] = GRAY
                    stack.append((nb, iter(adj[nb])))
                    advanced = True
                    break
                # BLACK -> forward/cross edge, no action
            if not advanced:
                color[node] = BLACK
                stack.pop()
    # self-loops are feedback too (so the renderer can flag them)
    for a, b, _l in edges:
        if a == b:
            feedback.add((a, b))

    # Longest-path ranking over forward edges only -> compact layers.
    fwd = [(a, b) for (a, b, _l) in edges if (a, b) not in feedback]
    children = defaultdict(list)
    indeg = {n: 0 for n in ids}
    for a, b in fwd:
        children[a].append(b)
        indeg[b] += 1
    rank = {n: 0 for n in ids}
    ind = dict(indeg)
    q = deque([n for n in ids if ind[n] == 0])
    while q:
        u = q.popleft()
        for v in children[u]:
            if rank[u] + 1 > rank[v]:
                rank[v] = rank[u] + 1
            ind[v] -= 1
            if ind[v] == 0:
                q.append(v)

    layers = {}
    for n in ids:
        layers.setdefault(rank[n], []).append(n)

    # Reduce edge crossings by reordering nodes within each layer.
    layers = _minimize_crossings(ids, layers, edges, rank, feedback)
    return rank, layers, feedback


def _minimize_crossings(ids, layers, edges, rank, feedback, passes=6):
    """Order nodes within each layer to minimize edge crossings.

    The ordering phase of the Sugiyama framework: a barycenter heuristic swept
    down then up across the layers, keeping whichever ordering yields the fewest
    crossings. Uses only forward edges (feedback edges are drawn as separate
    return paths, so they don't drive the main-flow ordering). Heuristic and
    cheap — ideal for the small graphs these flowcharts hold.
    """
    if max(layers) < 1:
        return layers
    # Forward adjacency (skip feedback + self-loops).
    preds, succs = defaultdict(list), defaultdict(list)
    for a, b, _l in edges:
        if a == b or (a, b) in feedback or rank[a] >= rank[b]:
            continue
        succs[a].append(b)
        preds[b].append(a)

    order = {r: list(layers[r]) for r in layers}
    max_r = max(order)

    def norm_pos(n):
        lst = order[rank[n]]
        return lst.index(n) / max(1, len(lst) - 1)

    def sweep(down):
        rng = range(1, max_r + 1) if down else range(max_r - 1, -1, -1)
        nbr = preds if down else succs
        for r in rng:
            lst = order[r]
            keys = {}
            for i, n in enumerate(lst):
                ns = nbr.get(n, [])
                keys[n] = (sum(norm_pos(x) for x in ns) / len(ns)) if ns \
                    else (i / max(1, len(lst) - 1))
            order[r] = sorted(lst, key=lambda n: keys[n])

    def crossings():
        idx = {n: order[rank[n]].index(n) for n in ids}
        total = 0
        for r in range(max_r):
            es = [(a, b) for (a, b, _l) in edges
                  if rank[a] == r and rank[b] == r + 1 and (a, b) not in feedback]
            for i in range(len(es)):
                a1, b1 = es[i]
                for j in range(i + 1, len(es)):
                    a2, b2 = es[j]
                    if (idx[a1] - idx[a2]) * (idx[b1] - idx[b2]) < 0:
                        total += 1
        return total

    best = {r: list(order[r]) for r in order}
    best_c = crossings()
    for p in range(passes):
        sweep(down=(p % 2 == 0))
        c = crossings()
        if c < best_c:
            best_c, best = c, {r: list(order[r]) for r in order}
            if best_c == 0:
                break
    return best
