# Slide Layout Catalog

Twelve layouts cover the great majority of real decks. Pick the layout that
matches the *communication job* of the slide, not just the content type — a
good deck alternates layouts so it doesn't read as page after page of bullets.

Each entry lists its fields and the situation it's for.

---

## `title`
Opening slide. Big title on a clean field with a left accent block.
- `title` (req), `subtitle`, `eyebrow`, `author`, `date`
- `background` (hex, e.g. `"000000"`) fills the whole slide for a dark cover and
  flips the text (and the top-right classification marker) to a light color
  automatically; `color` overrides that text color. Handy for a temporary cover
  before a house cover art is dropped in.
- Use once, at the start.

## `section`
Full-bleed divider in the primary color with an optional large number. Signals
"new chapter" and gives the audience a breath.
- `title` (req), `subtitle`, `number` (e.g. `"01"`)
- Use to separate the 2–5 major parts of a longer deck.

## `bullets`
The workhorse content slide. Title + optional one-line `lead` + bullet list
with up to two indent levels.
- `title` (req), `bullets[]` (req), `eyebrow`, `lead`
- Keep to ~5 top-level bullets. If you have more, split or switch layout.

## `content`
Title + flowing paragraphs (no bullets).
- `title` (req), `body` (string or string[], one paragraph each), `eyebrow`
- **Use sparingly.** The default style avoids prose: prefer `bullets` and write
  body points as short noun-phrase fragments, not sentences. Reach for `content`
  only when a genuine narrative paragraph is unavoidable.

## `two_column`
Two parallel lists or text blocks under one title. For related-but-distinct
groupings (e.g. "what worked / what to watch").
- `title` (req), `left` & `right`, each `{heading, bullets[] | body}`

## `comparison`
Two contrasting *cards* with colored headers. Stronger visual opposition than
`two_column` — use for Before/After, Option A/B, Us/Them, Pros/Cons.
- `title` (req), `left` & `right`, each `{heading, bullets[]}`

## `metrics`
Up to 4 KPI cards in a row, each a big number + label + small sublabel. The
fastest way to land "here are the numbers that matter."
- `title` (req), `metrics[]` of `{value, label, sublabel}` (max 4)
- Keep `value` short — a number with at most a tiny unit ("94%", "2 nm",
  "138억"). It renders very large, so a long value (especially Korean like
  "138억 년" or "930억 광년") wraps onto a second line and collides with the
  label. Move the unit/context into `sublabel` ("년", "광년", "+18% YoY").

## `quote`
One large pulled quote with attribution. For a customer voice, a principle, or
a mission statement. Give it its own slide — the whitespace is the point.
- `quote` (req), `attribution`, `title`
- On bar-style templates (e.g. `samsung`, `report`) it draws the same header bar
  as the other content slides — pass a short `title` for it (e.g. "맺는 말") — so
  the quote slide stays on-template instead of floating on a bare field. On the
  other templates it remains a clean, header-less full-field quote.

## `image`
A picture with optional supporting text.
- `image` (path), `position`: `right` (default) | `left` | `full`, `title`,
  `bullets[]` or `body`
- `full` makes it edge-to-edge with the title in a bar across the bottom.
- Image is fitted (never stretched). A missing path renders a labeled
  placeholder rather than failing, so the deck still builds.

## `table`
A data grid with an optional header row and zebra striping.
- `title` (req), `table` = `{header[], rows[][]}`
- Keep ≤ ~6 cols / ~8 rows per slide.

## `chart`
A native, editable chart (column/bar/line/pie/area). It stays editable in
PowerPoint — users can retype the data.
- `title` (req), `chart` (see `spec-schema.md` for the data shape)
- One clear comparison per chart beats a dense multi-series jumble.

## `diagram`
A conceptual diagram drawn as **native, editable PowerPoint shapes** (boxes,
arrows, connectors, polygons) — never a flattened image. Every box and label
can be moved, re-typed, and recolored in PowerPoint/Keynote afterward. Use this
to show relationships and flow, not just lists.
- `title` (req), `eyebrow`, `diagram` = `{type, nodes[]}`
- `type`: one of
  - `process` — horizontal steps connected by arrows (a workflow/pipeline)
  - `cycle` — steps arranged in a ring with arrows (a repeating loop)
  - `hierarchy` — a root box with child boxes below (org chart; 2 levels:
    `nodes[0]` is the root, the rest are its children)
  - `pyramid` — stacked bands, narrow top → wide base (layered model)
  - `funnel` — stacked bands, wide top → narrow bottom (conversion/filtering)
  - `timeline` — milestones along a horizontal spine (roadmap/schedule)
  - `flowchart` — an arbitrary directed graph with branches/merges, auto-laid
    out left→right into layers (a "Mermaid-like" flow, but native shapes)
- each node is `{title, desc}` (desc optional, shown smaller). For `flowchart`,
  give each node an `id` and add `edges`: a list of `[from_id, to_id]` or
  `{from, to, label}` (the optional `label` annotates the connector, e.g. a
  branch condition). Layout is automatic from the edges, with crossing
  minimization; feedback/loop-back edges are detected and drawn as a dashed
  return path routed clear of the forward arrows (so the direction stays
  unambiguous). Works in both directions; `LR` reads most naturally for wide
  16:9 slides.
- Aim for 3–6 nodes (a bit more is fine for `flowchart`); beyond that it gets
  cramped. For a picture diagram you already have as a file, use the `image`
  layout instead (not editable, but drops in any pre-made graphic).

## `closing`
Full-bleed closing slide ("Thank you", "Questions?", contact/CTA).
- `title` (req, defaults to "Thank You"), `subtitle`

---

## Composing a deck

A reliable arc for a ~10–15 slide deck:

1. `title`
2. `bullets` agenda or a `section`
3. `metrics` or `chart` to establish the situation with evidence
4. alternating `bullets` / `two_column` / `content` for the argument
5. `comparison` or `chart` at the turning point
6. `quote` for a human beat
7. `section` dividers between major parts (longer decks)
8. `content` for the forward look
9. `closing`

Vary the layout every 1–2 slides. If three `bullets` slides appear in a row,
convert one to `metrics`, `two_column`, `chart`, or `quote`.
