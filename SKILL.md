---
name: ppt-generator
description: >-
  Create, read, or edit PowerPoint (.pptx) decks. CREATE complete editable
  presentations from a topic, outline, notes, or source document in a chosen
  visual template — use WHENEVER the user wants to create, build, make, or
  draft a slide deck / presentation / PPT / PPTX / slides / "발표 자료" /
  "자료 만들어" / pitch deck / keynote, even if they only give the topic
  ("make slides about X", "turn this doc into a deck", "투자자용 10장 피치").
  EDIT or MODIFY an existing .pptx in place ("이 파일 수정해줘", "edit this
  deck", "fix the title", "swap the logo", "update the chart numbers"):
  replace text (formatting preserved), set/add text boxes, edit tables and
  chart data, replace/add/refit images, delete or reorder slides — keeping the
  original design. FIT images to a fixed template box or placeholder without
  distortion ("템플릿 크기에 맞춰 이미지 넣어줘"). RESTYLE / re-render a deck
  into another template or look (corporate, minimal, dark, vibrant, academic;
  "다른 템플릿으로", "이 템플릿으로 바꿔줘"). READ / EXTRACT a deck's text,
  tables, chart data, and images ("내용 추출", "이 ppt 읽어줘"). Produces a
  real, fully editable .pptx via python-pptx — not a description of one.
---

# PPT Generator

Turn content into a finished `.pptx` deck. The approach separates **what** goes
on each slide (a JSON *presentation spec*) from **how** it looks (a *template*).
You author the spec; a bundled renderer turns it into native, editable
PowerPoint. The same spec can be re-rendered in any template by changing one
field.

This mirrors how AI presentation tools like Presenton work — content generation
feeds a structured layout, which a renderer turns into the final file — but it
runs fully locally with no server.

The skill works in two directions:
- **Generate** a new deck from a presentation spec (the main workflow below).
- **Read / edit** an existing `.pptx` — extract its content, or modify it in
  place without disturbing the design (see "Reading and editing existing decks").

## When to use this

Any request to create, draft, or restyle a slide deck / presentation / PPT.
The user usually gives you a topic, an outline, rough notes, or a document and
expects slides back. Your job is to do the *content design* (what each slide
says and which layout carries it best), then render it.

## Requirements

Needs Python with `python-pptx` (charts/tables/images included). Pillow is used
for image fitting. Check and install if missing:

```bash
python3 -c "import pptx" 2>/dev/null || pip install python-pptx Pillow
```

Optional, only for visual QA (`scripts/thumbnail.py`): LibreOffice (`soffice`)
to convert `.pptx`→PDF, plus `pdftoppm` (poppler) or PyMuPDF to rasterize it.
The reading/editing scripts (`extract_pptx.py`, `inspect_pptx.py`,
`edit_pptx.py`, `ooxml.py`) need only `python-pptx` + Pillow.

## Workflow

### 1. Understand the content and audience

Before writing slides, get clear on:
- **Topic & key message** — what should the audience remember?
- **Audience & setting** — execs, investors, students, a team? Big screen?
- **Length** — how many slides (default to a tight 8–14 if unspecified).
- **Source material** — if the user gave a document/notes, read it and extract
  the structure; don't invent facts that aren't there.

If the request is vague ("make me a deck about our roadmap"), make reasonable
choices and proceed — produce a solid draft rather than stalling on questions.
Ask only if something genuinely blocks you (e.g. no topic at all).

### 2. Choose a template

Match the audience and tone. Default to `corporate` when unsure. See
`references/templates.md` for the full catalog:

`corporate` (business default) · `minimal` (design-forward) · `dark`
(tech/keynote) · `vibrant` (pitch/marketing) · `academic` (lectures/research).

### 3. Design the slides as a spec

Write a JSON spec (schema in `references/spec-schema.md`, layouts in
`references/layouts.md`, full worked example in `assets/spec.example.json`).

The craft is in **slide design**, not just dumping text:
- **One idea per slide.** If a slide has two messages, split it.
- **Titles are noun phrases, never sentences.** End every title on a noun —
  "Revenue Growth" or "18% Revenue Growth", never "Revenue grew 18%". A title
  labels the slide; it does not make a claim with a verb. Avoid imperative
  titles too ("Pick a Template" → "Template Choice").
- **No full sentences in the body.** Write body content as short noun phrases
  and fragments, not narrative prose. Drop articles and trailing verbs:
  "Enterprise segment outperformed plan" → "Enterprise above plan". This keeps
  slides scannable and stops them from becoming a document.
- **Every body point is a bullet.** Separate each idea with its own bullet, and
  default to the `bullets` / `two_column` / `comparison` layouts. Avoid the
  prose `content` layout — reserve it for a genuine exception, and even then
  break the points into bullets. (A single italic `subtitle` lead-in line under
  the title is the one place a full sentence is fine.)
- **Pick the layout that fits the job** — numbers → `metrics` or `chart`;
  contrast → `comparison`; relationships/flow → `diagram`; a human beat →
  `quote`. Don't default everything to one layout.
- **Vary the rhythm.** Never stack three same-layout slides in a row. Use
  `section` dividers to break a longer deck into parts.
- **Keep slides sparse.** ~5 bullets max, ≤4 metrics, ≤~6×8 tables. Slides
  support the speaker; they aren't the document.

A typical arc: `title` → agenda/`section` → `metrics`/`chart` (situation) →
alternating content layouts (argument) → `comparison`/`quote` (turn) →
`content` (forward look) → `closing`. See `references/layouts.md` for more.

Write the spec to a file, e.g. `deck.spec.json`.

### 4. Validate, then render

```bash
python3 scripts/validate_spec.py deck.spec.json     # catches mistakes early
python3 scripts/build_pptx.py   deck.spec.json -o deck.pptx
```

The validator reports missing fields, bad layout names, and malformed
chart/table data. Fix any errors it prints before rendering.

### 5. Verify and iterate

State the real outcome: how many slides, which template, where the file is.
Don't claim it "looks great" — you can't see it render. For visual QA, render a
contact-sheet grid of every slide and inspect it (needs LibreOffice plus
`pdftoppm` or PyMuPDF):

```bash
python3 scripts/thumbnail.py deck.pptx          # -> deck.grid.png (+ per-slide PNGs)
```

Open the grid to catch text cutoff, overflow, or clashing colors. If LibreOffice
isn't installed, the script says so — then tell the user to open the file in
PowerPoint/Keynote and offer concrete next steps: adjust length, swap the
template (one-field change → re-render), tighten wording, add a chart, etc.

To **restyle** an existing deck into another look, change only the top-level
`template` field in the spec and re-render — content is untouched.

## Reading and editing existing decks

Besides generating new decks, the skill can read an existing `.pptx` and modify
it in place — keeping every untouched shape, image, and bit of formatting.

> Works on ordinary, unprotected `.pptx` only. A DRM-protected file must first
> be exported to a plaintext copy through its DRM client by an authorized user;
> these tools do not bypass DRM.

**Read / extract** content:

```bash
python3 scripts/inspect_pptx.py deck.pptx          # map slides + addressable shapes
python3 scripts/extract_pptx.py deck.pptx           # text + tables + chart data (markdown)
python3 scripts/extract_pptx.py deck.pptx --images ./imgout --json
```

`extract_pptx.py` pulls text, tables, **chart data** (type, categories, series),
and embedded images; `inspect_pptx.py` also reports each chart's type and
series/category counts.

`inspect_pptx.py` is the starting point for edits — it prints each shape's
`(slide, shape)` address, type, and text so you know what to target.

**Edit** in place via a JSON *edit spec* (schema in `references/edit-spec.md`).
Validate it first — `validate_spec.py` auto-detects an edit spec and checks op
names, required fields, fit modes, and chart values:

```bash
python3 scripts/validate_spec.py edits.json         # catches mistakes early
python3 scripts/edit_pptx.py edits.json             # source/output from spec
python3 scripts/edit_pptx.py edits.json -i in.pptx -o out.pptx
```

Supported operations: `replace_text` (formatting-preserving find/replace across
runs), `set_text`, `set_table_cell`, `set_chart_data` (replace a chart's
categories/series in place), `replace_image`, `add_image`, `add_textbox` (add
new text), `delete_shape`, and slide-level `delete_slide` / `move_slide`.
`replace_text` is the safest bulk edit — it matches by content, so it is
unaffected by shape indices. Structural ops (delete/move slide, delete shape)
shift indices, so order them last or re-inspect between edits.

Images are fitted to a **fixed box** without distortion: `add_image` places into
a template placeholder (`into_shape`) or inch box, `replace_image` keeps the
existing picture's frame, and `fit_image` resizes a picture already on the slide
to a box — defaulting to the whole slide, so it snaps full-page images to fill
the slide template exactly. All take a `fit` mode (`contain` = letterbox inside,
`cover` = fill and crop, `stretch`).

Editing preserves the original design. To instead **restyle** a deck into a
different template, use the generation path: change the `template` field in a
presentation spec and re-render.

**Image-only decks.** Some decks are slides exported as one full-page image each
(common with translated or design-tool exports) — `inspect_pptx.py` shows every
shape as `PICTURE` with no text. There is no text to edit. Options: refit the
images to the slide template with `fit_image`; wrap each image in a template
(header/footer chrome via `add_image into_shape` + `add_textbox`); replace
images; or, to get editable text, recreate the content as a spec and re-render
(the generation path). Don't promise text edits on such a deck.

**Escape hatch — raw OOXML.** For changes python-pptx can't reach (theme colors,
slide-master tweaks, custom geometry), unpack the deck to its XML tree, edit the
XML, and repack:

```bash
python3 scripts/ooxml.py unpack deck.pptx unpacked/   # XML pretty-printed
# …edit files under unpacked/ppt/… …
python3 scripts/ooxml.py pack   unpacked/ out.pptx
```

## Output conventions

- Save specs and `.pptx` files where the user is working, or `claudedocs/` if
  there's no obvious home. Name by topic (e.g. `q4-review.pptx`).
- The output is a standard editable `.pptx` — charts and tables remain
  editable in PowerPoint, so the user can refine after you hand it off.

## Notes on extending

- New template: copy an entry in `scripts/templates.py`, adjust palette/fonts.
- New layout: add a `layout_*` function in `scripts/build_pptx.py` and register
  it in the `LAYOUTS` dict; document it in `references/layouts.md`.
- Images: pass a local file path in an `image` slide; a missing path renders a
  labeled placeholder so the build never fails mid-deck.
