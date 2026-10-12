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

// Or snap into a named template's content box (under header, above footer):
{"op": "fit_image", "slide": 2, "shape": 0, "template": "samsung", "fit": "contain"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `slide`, `shape` | yes | The picture to refit. |
| `fit` | no | `contain` / `cover` / `stretch` (default `cover`). |
| `template` | no | Fit into this template's standard content box (e.g. `samsung`), sized to the deck's actual slide dimensions. Box precedence: `template` > `into_shape` > inch box > whole slide. |
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

// Or drop into a named template's content box:
{"op": "add_image", "slide": 5, "image": "chart.png",
 "template": "samsung", "fit": "contain"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `image` | yes | Path to an image file. |
| `template` | no | Fit into this template's standard content box (e.g. `samsung`), sized to the deck's slide dimensions. Box precedence: `template` > `into_shape` > inch box. |
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

### `add_chrome` — brand a slide with a template's header/footer
Adds a template's chrome to one slide **in place** — the header (a filled bar or
an accent header, whichever the template uses), the brand / classification
marker, and a `current / total` page number — reusing the same drawing code as
`build_pptx.py`, so a chromed slide is identical to a natively rendered one.

This is the missing piece for a **mixed deck**: a few natively rendered template
slides plus foreign full-bleed slides dropped in (NotebookLM, Gamma, translated
or design-tool exports), which arrive with no chrome and look out of place.
`wrap_images.py` would rebuild the whole deck from images (losing the native
slides); the COM `--chrome` path stamps *every* slide (doubling chrome on the
ones that already have it). `add_chrome` brands only the slides you name, keeping
the deck and its slide count intact.

```json
// Typical recipe per foreign image slide: pull the picture under the header,
// then brand it. Run a renumber_pages at the end to unify page numbers.
{"op": "fit_image", "slide": 3, "shape": 0, "template": "samsung", "fit": "contain"}
{"op": "add_chrome", "slide": 3, "template": "samsung",
 "eyebrow": "solutions", "title": "저장소 내 AI 코파일럿"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `slide` | yes | 1-based slide to brand. |
| `template` | no | Template whose chrome to draw (default `samsung`). |
| `title` | no | Header title text. |
| `eyebrow` | no | Small label above the title (rendered uppercase). |
| `subtitle` | no | Optional line under the header. |
| `footer` | no | Footer text (left of the page number). Normally **omit it** — the footer should show only the page number. |
| `slide_no`, `total` | no | Page number values (default: this slide's physical position / deck length). |
| `page_number` | no | Set `false` to omit the page number (default `true`). |

Note that `fit_image` with the same `template` fits the picture into the box
*under* the header band, so the new bar never overlaps the image. The header is
drawn last (on top), so call `add_chrome` after `fit_image`. For templates with a
"bar" header (e.g. `samsung`) on a full-bleed foreign slide that has its own
baked-in title, expect the bar title and the image's own title to both show —
either keep the bar title short, set it to match, or omit `title`.

### `renumber_pages` — re-stamp page numbers across the deck
Rewrites every `N / M` page-number box to the slide's physical position over the
current total, keeping each box's formatting. After inserting, deleting, or
reordering slides the baked footer numbers go stale (a 6-slide deck grown to 9
still reads `1 / 6` … `6 / 6`); this re-stamps them all to `1 / 9` … `9 / 9` in
one pass. It matches the `current / total` footer that `build_pptx.py` and
`add_chrome` emit; a box that doesn't look like a page number is left untouched,
so it won't clobber body text that contains a slash.

```json
{"op": "renumber_pages"}
{"op": "renumber_pages", "format": "{n}/{total}"}        // no spaces
{"op": "renumber_pages", "format": "Page {n} of {total}"}
```
| Field | Required | Notes |
|-------|----------|-------|
| `format` | no | Number string template (default `"{n} / {total}"`); use `{n}` and `{total}`. |

Run it **last**, after any `add_chrome` / `duplicate_slide` / `delete_slide` /
`move_slide`, so it numbers the final slide order.

### `delete_shape` — remove a shape
```json
{"op": "delete_shape", "slide": 4, "shape": 3}
```

### `duplicate_slide` — copy a slide N times
Clones a slide and everything on it — shapes, text, and **images** (the image
relationships are remapped so the copies render correctly, which a naive XML
copy gets wrong). Copies land immediately after the source slide by default, or
at 1-based position `at`.

```json
{"op": "duplicate_slide", "slide": 2}                 // one copy, right after slide 2
{"op": "duplicate_slide", "slide": 2, "count": 14}    // 14 copies (e.g. one frame per image)
{"op": "duplicate_slide", "slide": 1, "count": 3, "at": 5}  // 3 copies starting at position 5
```
| Field | Required | Notes |
|-------|----------|-------|
| `slide` | yes | 1-based source slide to copy. |
| `count` | no | Number of copies (default 1). |
| `at` | no | 1-based position for the first copy (default: right after the source). |

**Recipe — wrap an image-only deck in a fixed template.** This is the main use:
you have N slide images and want each framed by a template's chrome (header bar,
logo, footer). python-pptx can't add framed slides on its own, so:

1. Open the template `.pptx`; pick one slide whose chrome you want as the frame.
   Strip its body content with `delete_shape` so only the chrome remains (or
   start from a slide that is already chrome-only).
2. `duplicate_slide` that frame `count: N-1` times (you already have one) — or
   delete the original content slides afterward and duplicate `count: N`.
3. `add_image` each picture into the body region of each framed slide
   (`fit: contain` keeps it undistorted).

Because the frame slide and the image count rarely match, `duplicate_slide` is
what makes "insert this image deck into our corporate template" actually work.

### `delete_slide` — remove a whole slide
```json
{"op": "delete_slide", "slide": 5}
```

### `move_slide` — reorder a slide
Moves the slide at `from` to position `to` (both 1-based).
```json
{"op": "move_slide", "from": 11, "to": 2}
```

> **Ordering caveat:** `duplicate_slide`/`delete_slide`/`move_slide` change slide numbers, and
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
