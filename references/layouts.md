# Slide Layout Catalog

Eighteen layouts cover the great majority of real decks. Pick the layout that
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
- `title` (req), `bullets[]` (req), `eyebrow`, `subtitle`
- (`lead` is a legacy alternative to `subtitle` that sits below it; prefer
  `subtitle`.)
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
- `alt` sets the picture's alt text (defaults to the slide title); screen
  readers announce it, and the audit flags pictures without one.
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
A native, editable chart (column, bar, stacked, line, area, pie, doughnut),
styled with the template's palette and fonts. It stays editable in
PowerPoint — users can retype the data.
- `title` (req), `chart` (see `spec-schema.md` for the data shape)
- One clear comparison per chart beats a dense multi-series jumble. For a
  single series, `highlight` the bar that carries the message and state the
  takeaway in `subtitle`.
- Ranking → `bar` (sorted, largest first); trend over time → `line` or
  `column`; share of a whole (≤ 5 parts) → `doughnut`; composition over time →
  `stacked_column`.

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

## `agenda`
Numbered agenda with round number badges — a real roadmap of the talk instead
of a bullet list. Up to 8 items; more than 5 flow into two columns.
- `title` (req), `items[]` (req): strings or `{title, desc}`, `current`
  (1-based, optional) highlights one item and dims the rest — reuse the same
  agenda before each section as a progress tracker.

## `big_number`
One hero figure, its label, and a one-line takeaway band at the bottom. For
the single number the audience must leave with ("9.8억 원", "61%").
- `title` (req), `value` (req, short), `label`, `takeaway`
- Use once or twice per deck; it loses force when repeated.

## `cards`
2–4 parallel cards in a row: options, pillars, workstreams, features. Each
card: heading, optional big `value`, and `bullets` or a short `body`. Cards
size to the tallest content.
- `title` (req), `cards[]` of `{title, value?, bullets[] | body}`
- Prefer `comparison` for exactly two *contrasting* items; `cards` is for
  parallel peers.

## `matrix`
2×2 matrix for prioritization or positioning (impact × effort, risk ×
likelihood). Quadrants are given in reading order.
- `title` (req), `quadrants[4]` (req) of `{title, items[]}` in order
  top-left, top-right, bottom-left, bottom-right
- `x_axis` / `y_axis` = `{"low": "...", "high": "..."}` label the axes;
  `highlight` (1–4) fills the quadrant the slide argues for.

## `icon_grid`
3–6 short items, each with an icon in a soft circle, a title and a one-line
description: capabilities, benefits, principles, service lines. 4 or fewer sit
in one row; 5–6 wrap to two rows of three.
- `title` (req), `items[]` (req) of `{icon, title, desc}`
- `icon` is a Phosphor name or a concept alias (`references/icons.md`); an
  unknown icon renders the item's number instead.
- `cards` also accept an `icon` per card, drawn beside the card heading.

## Numbered lists
Any `bullets` slide (and `two_column` columns / `comparison` cards) can set
`"numbered": true` for real auto-numbered items (`1.`, `2.`, …) — use it for
steps, ranked items and checklists; sub-items keep dash bullets. `start_at`
sets the first number.

## `closing`
Full-bleed closing slide ("Thank you", "Questions?", contact/CTA).
- `title` (req, defaults to "Thank You"), `subtitle`

---

## Composing a deck

A reliable arc for a ~10–15 slide deck:

1. `title`
2. `agenda` or a `section`
3. `metrics`, `big_number` or `chart` to establish the situation with evidence
4. alternating `bullets` / `two_column` / `content` for the argument
5. `comparison` or `chart` at the turning point
6. `quote` for a human beat
7. `section` dividers between major parts (longer decks)
8. `content` for the forward look
9. `closing`

In an `"assertion"` deck the same arc works, but each content slide's title
carries its conclusion and the layout below it is the proof: a `chart` or
`metrics` under a claim about numbers, a `comparison` under a before/after
claim, a `diagram` under a claim about how something works.

Vary the layout every 1–2 slides. If three `bullets` slides appear in a row,
convert one to `metrics`, `two_column`, `chart`, or `quote` (the validator
warns on any three-in-a-row).

## Continuation slides

A `bullets` or `table` slide that would have to shrink below 85% to fit is
split automatically at item boundaries (a top-level bullet stays with its
sub-bullets; table rows stay whole) into balanced parts. Continuation parts
get the title plus " (계속)" (Korean titles) or " (cont.)", repeat the table
header, continue numbering, and drop the subtitle and notes. The build and the
validator report each `SPLIT`. Set `"split": false` on a slide (or
`"auto_split": false` on the deck) to keep it on one slide — then it shrinks
and may `OVERFLOW` instead. Splitting is a safety net: it is still better to
cut words or choose a denser layout.

## How the renderer fits text

Every text box is measured before it is written. Titles first shrink (to 80%)
to stay on one line, then wrap to two lines and push the body down. In an
`"assertion"` deck titles keep full size and wrap instead (they only shrink,
by up to 10%, to avoid stranding one word on line two), so headlines stay a
consistent size. Body text,
card text, table cells and diagram labels scale down uniformly (to 70%) to fit
their box; KPI values shrink to stay on one line. Anything that still doesn't
fit is reported as `OVERFLOW`. Comparison cards and tables size themselves to
their content, and table columns get width in proportion to their text.

Measurements use Malgun Gothic's real Hangul advance (1 em) and PowerPoint's
line rule (single spacing = 1.2 × size, scaled by the line-spacing percent).
The renderer deliberately writes no `normAutofit`: PowerPoint only applies a
stored shrink and recomputes it on edit, while LibreOffice recomputes it on
every render — which would hide real overflow in the preview check.
