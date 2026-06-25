# Edit Spec Schema (modifying an existing deck)

Two ways to work with a deck exist in this skill:

- **Generate** a new deck from a *presentation spec* → `build_pptx.py` (see `spec-schema.md`).
- **Edit** an existing `.pptx` in place from an *edit spec* → `edit_pptx.py` (this file).

Editing keeps every untouched shape, image, master, and bit of formatting
exactly as it was — it does not re-render the deck. Use it to update an existing
file (fix a number, swap a logo, change a date across all slides) without
disturbing its design.

> Works on ordinary, unprotected `.pptx` files only. A DRM-locked file must
> first be exported to a plaintext copy through its DRM client by an authorized
> user. These tools do not bypass DRM.

## Workflow

```bash
# 1. See what's there and how to address each shape (slide + shape index)
python scripts/inspect_pptx.py deck.pptx          # add --json for machine output

# 2. (optional) Dump full text / tables / images
python scripts/extract_pptx.py deck.pptx --images ./imgout

# 3. Write an edit spec, validate it, then apply it
python scripts/validate_spec.py edits.json        # checks ops/fields (auto-detects edit spec)
python scripts/edit_pptx.py edits.json            # source/output read from spec
python scripts/edit_pptx.py edits.json -i in.pptx -o out.pptx
```

`validate_spec.py` auto-detects an edit spec (top-level `operations`) and checks
op names, required fields, `fit` modes, numeric chart values, and image paths
before you run it. Force the check with `--edit` if needed.

Shapes are addressed by **`slide`** (1-based, as shown by `inspect_pptx.py`) and
**`shape`** (0-based index within that slide's shape list). Always run
`inspect_pptx.py` first to get the right indices.

## Top-level fields

| Field | Type | Notes |
|-------|------|-------|
| `source` | string | Path to the input `.pptx`. Overridden by `-i`. |
| `output` | string | Path to write the edited `.pptx`. Overridden by `-o`. |
| `operations` | array | Operations applied in order. |

Use a different `output` than `source` to keep the original intact.

## Operations

Each operation is an object with an `op` field plus its arguments.

### `replace_text` — find/replace (formatting-preserving)
Replaces text across the whole deck, or only the listed slides. Works inside
text boxes, grouped shapes, and table cells. Each match keeps the character
formatting (font, size, color, bold) of the run where it begins — even when the
matched text is split across several runs.

```json
{"op": "replace_text", "find": "2024", "replace": "2025"}
{"op": "replace_text", "find": "Confidential", "replace": "Public", "slides": [1, 2]}
```
| Field | Required | Notes |
|-------|----------|-------|
| `find` | yes | Literal substring to match (not a regex). |
| `replace` | no | Replacement (default `""`, i.e. delete). |
| `slides` | no | List of 1-based slide numbers; omit for all slides. |

### `set_text` — replace one shape's text
Replaces all text in a single text shape, keeping the first run's font.

```json
{"op": "set_text", "slide": 3, "shape": 2, "text": "New Title"}
```

### `set_table_cell` — edit one table cell
```json
{"op": "set_table_cell", "slide": 7, "shape": 1, "row": 0, "col": 2, "text": "Q3"}
```
`row`/`col` are 0-based. The shape must be a table (see `inspect_pptx.py`).

### `set_chart_data` — replace a chart's data
Replaces the categories and series of an existing chart in place, keeping its
type and styling. Get the chart's `(slide, shape)` address and current shape
from `inspect_pptx.py`; dump its current data with `extract_pptx.py`.

```json
{"op": "set_chart_data", "slide": 9, "shape": 4,
 "categories": ["Q1", "Q2", "Q3", "Q4"],
 "series": [
   {"name": "Revenue", "values": [10, 22, 18, 30]},
   {"name": "Cost",    "values": [8, 12, 11, 15]}
 ]}
```
| Field | Required | Notes |
|-------|----------|-------|
| `series` | yes | List of `{name, values}`; each `values` length should match `categories`. |
| `categories` | no | Category labels; defaults to the chart's current categories. |

The shape must be a chart. Works for column/bar/line/pie and other category
charts (replace_data semantics from python-pptx).

### `fit_image` — resize an existing picture to a fixed box
Resizes/repositions a picture already on the slide to fit a box, without
distortion. With no box given, the box is the **whole slide** — so this snaps an
image to fill the slide template exactly. Useful for decks whose slides are
single full-page images that sit slightly off the slide (letterboxed or
misaligned).

```json
// Make every full-page image fill the slide template exactly:
{"op": "fit_image", "slide": 1, "shape": 0, "fit": "cover"}

// Or fit into a specific template box / placeholder:
{"op": "fit_image", "slide": 3, "shape": 0, "into_shape": 2, "fit": "contain"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `slide`, `shape` | yes | The picture to refit. |
| `fit` | no | `contain` / `cover` / `stretch` (default `cover`). |
| `into_shape` | no | Fit into this shape's box instead of the whole slide. |
| `left_in`+`top_in`+`width_in`+`height_in` | no | An explicit inch box instead of the whole slide. |

Any prior crop on the picture is cleared before refitting.

### `replace_image` — swap a picture's image, keep its box
```json
{"op": "replace_image", "slide": 1, "shape": 0, "image": "new-logo.png"}
{"op": "replace_image", "slide": 1, "shape": 0, "image": "photo.jpg", "fit": "contain"}
```
The target shape must be a picture. Its **box (position + size) is preserved** —
the fixed template frame stays put. Because the new image rarely shares the old
one's aspect ratio, it is re-fitted into that box per `fit` (default `cover`),
and any previous crop is cleared first, so the image is never distorted.

### `add_image` — add a new picture, fitted to a box
Place an image into a fixed box, scaled to fit without distortion. The box can
be an existing template shape (the image takes that shape's exact geometry) or
explicit inch coordinates.

```json
// Fill a template placeholder shape, replacing it:
{"op": "add_image", "slide": 5, "into_shape": 2, "image": "chart.png",
 "fit": "cover", "replace_target": true}

// Or place into an explicit box:
{"op": "add_image", "slide": 5, "image": "chart.png",
 "left_in": 1, "top_in": 1.5, "width_in": 4, "height_in": 3, "fit": "contain"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `image` | yes | Path to an image file. |
| `into_shape` | no | Index of an existing shape whose box (left/top/width/height) the image is fitted into — e.g. a template "photo here" rectangle. Found via `inspect_pptx.py`. |
| `replace_target` | no | With `into_shape`, delete that placeholder shape after reading its box (default `false`). |
| `left_in`, `top_in` | no | Position in inches when no `into_shape` (default 1, 1). |
| `width_in`, `height_in` | no | Box size in inches. Omit both (and `into_shape`) to add at native size. |
| `fit` | no | See below (default `contain` for `add_image`). |

### Fit modes (`fit`)
How an image is sized to a fixed box, used by `add_image` and `replace_image`:

| Mode | Behavior |
|------|----------|
| `contain` | Scale to fit **inside** the box, centered; no crop, no distortion (may leave margins). |
| `cover` | Scale to **fill** the box exactly, cropping the overflow; no distortion. Matches how PowerPoint template photo frames behave. |
| `stretch` | Stretch to the box exactly; may distort. |

### `add_textbox` — add a new text box
Adds new text onto a slide (vs `set_text`, which only edits an existing shape).

```json
{"op": "add_textbox", "slide": 1, "text": "DRAFT",
 "left_in": 0.5, "top_in": 0.3, "width_in": 3, "height_in": 0.6,
 "size_pt": 24, "bold": true, "color": "CC0000", "align": "left"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `slide`, `text` | yes | Target slide and the text. |
| `left_in`/`top_in`/`width_in`/`height_in` | no | Inch box (defaults 1,1,4,1). |
| `font`, `size_pt`, `bold`, `color`(hex), `align`(left/center/right) | no | Formatting. |

### `delete_shape` — remove a shape
```json
{"op": "delete_shape", "slide": 4, "shape": 3}
```

### `delete_slide` — remove a whole slide
```json
{"op": "delete_slide", "slide": 5}
```

### `move_slide` — reorder a slide
Moves the slide at `from` to position `to` (both 1-based).
```json
{"op": "move_slide", "from": 11, "to": 2}
```

> **Ordering caveat:** `delete_slide`/`move_slide` change slide numbers, and
> `delete_shape`/`add_*` change shape indices. Put structural ops last, or
> re-run `inspect_pptx.py` between edits. Content-matched ops (`replace_text`)
> are immune to index shifts.

## Addressing caveats

- Shape indices shift if you **delete** or **add** shapes earlier in the same
  slide. Order `delete_shape`/`add_image` ops last, or re-inspect between runs.
- `replace_text` is the safest bulk edit — it matches by content, not index.
- To restyle a deck into a different *template* (not just edit content), use the
  generation path instead: change the `template` field in a presentation spec
  and re-render with `build_pptx.py`.
