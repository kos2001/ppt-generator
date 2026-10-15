# Presentation Spec Schema

A presentation is a single JSON object: top-level deck settings plus a `slides`
array. Each slide names a `layout` and carries the fields that layout uses.
Build it once, render into any template. Render with:

```
python scripts/build_pptx.py spec.json -o deck.pptx
```

## Top-level fields

| Field | Type | Default | Notes |
|-------|------|---------|-------|
| `template` | string | `"standard"` | One of the template names (see `templates.md`). |
| `title` | string | `"presentation"` | Used for the default output filename. |
| `footer` | string | — | Footer text shown on interior content slides. Per-slide `footer` overrides it. |
| `slide_numbers` | bool | `true` | Show page numbers on interior content slides. |
| `title_style` | string | `"topic"` | `"topic"` = noun-phrase labels; `"assertion"` = each content-slide title is a one-sentence takeaway (≤ 2 lines). A slide may override it with its own `title_style`. |
| `font` | string | template's | Override heading + body Latin font for the whole deck. |
| `ea_font` | string | `"Malgun Gothic"` | East Asian (Korean/CJK) font written to every run. |
| `auto_split` | bool | `true` | Split overlong `bullets` / `table` slides onto continuation slides. |
| `header_label` | string | template's | Top-right classification label (`report` template); `""` disables. |
| `slides` | array | **required** | The slides, in order. |

Title / section / closing / full-bleed image slides never show a footer or
page number — they are meant to read as clean dividers.

## Slide object

Every slide has a `layout` (defaults to `"bullets"` if omitted) plus
layout-specific fields. Common optional fields:

- `eyebrow` — small uppercase kicker above the title (most content layouts).
- `subtitle` — one italic line under the header, above the body. Works on every
  content layout (bullets, content, two_column, comparison, metrics, table,
  chart, image, diagram). Use it to state the slide's angle in a phrase, e.g.
  "The skill does the content design, not just text dumping."
- `footer` — overrides the deck footer for this one slide.
- `title_style` — overrides the deck's `title_style` for this slide (e.g.
  `"topic"` on an agenda slide in an assertion deck).
- `notes` — speaker notes (string, or list of strings = one paragraph each).
  Put the detail and the spoken explanation here, not on the slide.

See `layouts.md` for the full per-layout field reference and when to use each.

## Bullet items

Anywhere `bullets` appears, each item is either a plain string or an object:

```json
{ "text": "Main point", "bold": true }
{ "text": "Supporting detail", "level": 1 }
```

- `level` (0–2) indents and lightens sub-points. Keep most bullets at level 0;
  use level 1 sparingly for support. Level 2 exists but rarely reads well.
- `bold` emphasizes a single bullet.

## Quick layout-to-fields map

```
title       title, subtitle, eyebrow, author, date, background(hex), color(hex)
section     title, subtitle, number
bullets     title, eyebrow, subtitle, bullets[], numbered, start_at, split   (lead: legacy)
content     title, eyebrow, body (string | string[])
two_column  title, left{heading,bullets|body}, right{...}
comparison  title, left{heading,bullets}, right{heading,bullets}
metrics     title, metrics[]{value,label,sublabel}        (max 4; keep value short)
quote       quote, attribution, title (title shows in the bar on bar themes)
image       title, image (path), alt, position(right|left|full), bullets|body
table       title, table{header[], rows[][]}, split
chart       title, chart{type,categories[],series,highlight,number_format}
diagram     title, diagram{type, nodes[]{id?,title,desc}, edges[]?, direction?}
agenda      title, items[] (str | {title,desc}), current
big_number  title, value, label, takeaway
cards       title, cards[]{title,icon,value,bullets|body}    (2–4)
icon_grid   title, items[]{icon,title,desc}               (3–6)
matrix      title, quadrants[4]{title,items[]}, x_axis{low,high}, y_axis{low,high}, highlight
closing     title, subtitle
(any)       notes, footer, eyebrow, subtitle
```

## Diagram data

Native, editable diagrams (boxes/arrows/connectors stay editable in PowerPoint):

```json
"diagram": {
  "type": "process",
  "nodes": [
    {"title": "Collect", "desc": "Gather inputs"},
    {"title": "Process", "desc": "Transform"},
    {"title": "Ship",    "desc": "Release"}
  ]
}
```

`type` ∈ `process` | `cycle` | `hierarchy` | `pyramid` | `funnel` | `timeline` |
`flowchart`. For `hierarchy`, `nodes[0]` is the root and the rest are its
children. Keep to 3–6 nodes.

`flowchart` draws an arbitrary directed graph (branches, merges, loop-backs),
auto-laid out: give each node an `id` and add `edges` (`[from, to]` pairs or
`{from, to, label}`), plus optional `direction` (`"LR"` default | `"TD"`).
Layering, back-edge handling, and crossing minimization are automatic.

```json
"diagram": {
  "type": "flowchart", "direction": "LR",
  "nodes": [{"id": "a", "title": "수집"}, {"id": "b", "title": "분석"},
            {"id": "c", "title": "배포"}],
  "edges": [["a", "b"], {"from": "b", "to": "c", "label": "승인"}]
}
```

See `layouts.md` for what each type communicates.

## Chart data

```json
"chart": {
  "type": "column",                     // see types below
  "categories": ["Q1", "Q2", "Q3"],
  "series": [                           // multi-series: list of {name, values}
    {"name": "2024", "values": [3, 4, 5]},
    {"name": "2025", "values": [4, 5, 7]}
  ],
  "number_format": "0.0",               // optional Excel format for axis + labels
  "data_labels": false                  // optional; see defaults below
}
```

`type` ∈ `column` | `bar` (horizontal, first category on top) |
`stacked_column` | `stacked_bar` | `line` | `area` | `pie` | `doughnut`.

- **Values must be numbers** with the same count as `categories`. Units go in
  `number_format`, e.g. `"#,##0"`, `"0%"` (values as fractions),
  `"\"$\"0.0\"M\""`, `"#,##0\"억\""`.
- **`highlight`** (single-series column/bar): a category name or index. That
  bar gets the accent color and the rest turn grey — use it to point at the
  one bar the slide is about.
- **`data_labels`** defaults to on for single-series charts (≤ 12 points) and
  pie/doughnut, off for multi-series. Pie/doughnut labels show percentages
  (`"percent": false` shows raw values).
- Series colors come from the template palette, so charts match the deck.

For a single series you may pass a bare number list instead of the `{name,
values}` form. `series_name`, `data_labels`, and `type` still apply alongside it:

```json
"chart": {
  "type": "column",
  "categories": ["Q2", "Q3"],
  "series": [12, 4],
  "series_name": "Incidents",
  "data_labels": true
}
```

Pie charts use one series; its values map to the categories.

## Table data

```json
"table": {
  "header": ["Account", "ARR", "Status"],
  "rows": [
    ["Globex", "$640K", "Renewed"],
    ["Initech", "$410K", "Expansion"]
  ]
}
```

Header is optional. Rows are styled with alternating background bands. Keep to
~6 columns and ~8 rows so cells stay legible; split large tables across slides.

## Minimal example

```json
{
  "template": "minimal",
  "title": "Launch Plan",
  "slides": [
    {"layout": "title", "title": "Launch Plan", "subtitle": "Q1 rollout"},
    {"layout": "bullets", "title": "Goals",
     "bullets": ["Ship to 3 regions", "Hit 99.9% uptime", "Onboard 50 teams"]},
    {"layout": "closing", "title": "Thank You"}
  ]
}
```
