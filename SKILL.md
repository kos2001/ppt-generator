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

## When to use this

Any request to create, draft, restyle, read, or edit a slide deck / presentation
/ PPT. The user gives you a topic, an outline, notes, a document, or an existing
`.pptx`, and expects slides back.

## Two modes — establish which one first

This skill does two **independent** jobs that share nothing but the file format.
Decide which the request needs *before* touching any tool, because the workflow,
scripts, and even the failure modes differ completely:

- **Mode A · Generate from a layout** — build a brand-new deck from a topic /
  outline / notes / source document, laid out into a chosen template. You do the
  content design (what each slide says, which layout carries it) and render a
  fresh `.pptx` from a spec. → **Mode A workflow** below.
- **Mode B · Work on an existing deck** — read or modify a `.pptx` the user
  already has, *keeping its design*: extract its content, or replace text, edit
  tables/charts, refit images, duplicate/reorder slides in place. → **Mode B
  workflow** below.

**How to pick.** If the request clearly implies one mode, go straight to it
("make slides about Q4 churn" → A; "fix the date in deck.pptx" → B). If it's
ambiguous — the user just points at a file without saying what to do, asks for
something that could be either ("이 내용으로 발표자료 만들어줘" when a `.pptx`
is attached), or names both creating and an existing file — **ask before
starting**:

> "두 가지로 진행할 수 있어요 — (A) 레이아웃에 맞춰 **새 덱을 생성**할까요, 아니면
> (B) **기존 파일을 수정**할까요?"

Why the choice comes first: generating re-renders the whole deck from a spec, so
it *replaces* the original design; editing preserves every untouched shape and
bit of formatting. Picking the wrong mode silently discards exactly what the
user meant to keep (or rebuilds something they wanted made fresh) — so the two
flows stay separate and you commit to one.

## Requirements

Needs Python with `python-pptx` (charts/tables/images included). Pillow is used
for image fitting. Check and install if missing:

```bash
python3 -c "import pptx" 2>/dev/null || pip install python-pptx Pillow
```

> **Windows:** `python3` is often a Microsoft Store stub that prints nothing and
> exits — use `python` or `py` instead (e.g. `py -c "import pptx"`). Substitute
> it in every `python3 …` command below. Paths with backslashes also need
> forward slashes or doubled backslashes inside JSON specs.

Optional, only for visual QA (`scripts/thumbnail.py`): LibreOffice (`soffice`)
to convert `.pptx`→PDF, plus `pdftoppm` (poppler) or PyMuPDF to rasterize it.
The reading/editing scripts (`extract_pptx.py`, `inspect_pptx.py`,
`edit_pptx.py`, `ooxml.py`) need only `python-pptx` + Pillow.

## Mode A — Generate a deck from a layout

### 1. Confirm the brief before building

A deck is expensive to redo: the template, length, audience, and key message
shape *every* slide, so guessing them wrong means rebuilding the whole thing.
The reader almost always has these in their head but doesn't volunteer them.
So **confirm the brief before you render** — don't silently pick defaults and
generate. Get clear on:

- **Topic & key message** — what should the audience remember?
- **Audience & setting** — execs, investors, students, a team? Big screen?
- **Length** — how many slides?
- **Template / look** — default to the user's house style if they have one
  (this user uses **`samsung`** only — don't ask, just use it). Otherwise
  confirm one of `corporate`, `minimal`, `dark`, `vibrant`, `samsung`, `report`,
  `academic` (see `references/templates.md`). The renderer default is `samsung`.
- **Source material** — if the user gave a document/notes, read it and extract
  the structure; don't invent facts that aren't there.

**What to do:** ask for whatever the request didn't already specify *and that
isn't fixed by a house style*, in **one short consolidated question** (not an
interrogation) — e.g. "몇 장 정도로, 누구 대상으로 만들까요?" — with your
recommended default offered as the first option so it's a one-tap confirm. Then
build. Skip the question
only when the user already gave the essentials, attached a spec, or explicitly
said to just go ahead ("알아서 만들어줘", "draft something") — in that case pick
sensible defaults (tight 8–14 slides, `corporate`) and proceed, naming the
choices you made so they can correct course.

The point isn't to stall — it's that a 30-second confirm beats regenerating 15
slides in the wrong template or length. When in doubt, ask.

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

## Mode B — Work on an existing deck

This mode reads an existing `.pptx` and modifies it in place — keeping every
untouched shape, image, and bit of formatting. Use it whenever the user has a
file already and wants it changed rather than rebuilt.

> Works on ordinary, unprotected `.pptx` only. A DRM-protected file must first
> be exported to a plaintext copy through its DRM client by an authorized user;
> these tools do not bypass DRM.

**DRM-protected decks.** Enterprise document-security DRM (e.g. Fasoo, MarkAny,
Softcamp — common in Korean corporations) encrypts the file so python-pptx can't
open it; `extract_pptx.py`/`inspect_pptx.py` will fail. **Do not help bypass or
crack the protection** — that defeats an intentional security control and
typically violates company policy and law. Refuse requests to circumvent it,
even when the user clearly has access. Instead, the first step is always to
obtain an *authorized* unprotected rendering, then proceed normally:

1. **Authorized export** — use the DRM client's decrypt / 반출(export) workflow
   (often needs manager/security approval). The result is a normal `.pptx`.
2. **Ask the owner / security admin** for an authorized plaintext copy if you're
   not the document owner.
3. **Export from a licensed viewer** — if you may view it and policy permits,
   save-as or export slides to images/PDF, then process those (image-only path).

If printing/export/capture is blocked by policy, that block is the intended
control — go through the approval process, don't work around it. Once an
unprotected copy or its images exist, the read/edit and image-only workflows
here apply unchanged.

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
new text), `delete_shape`, and slide-level `duplicate_slide` /
`delete_slide` / `move_slide`. `duplicate_slide` copies a slide and its images
N times — the primitive for wrapping an image-only deck in a fixed template
(duplicate one chrome-only frame slide per image, then `add_image` into each;
see `references/edit-spec.md`).
`replace_text` is the safest bulk edit — it matches by content, so it is
unaffected by shape indices. Structural ops (duplicate/delete/move slide, delete
shape) shift indices, so order them last or re-inspect between edits.

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
(common with NotebookLM, translated, or design-tool exports) — `inspect_pptx.py`
shows every slide as a single `PICTURE` with no text. There is **no text to
edit**, so don't promise text edits; pick a path by what the user actually wants:

| Goal | Path | Editable text? |
|------|------|----------------|
| Fix wording / restyle / make it editable | Recreate content as a spec → re-render (**Mode A**). View each image, transcribe its title/bullets/table into the spec, map to the layout that fits (comparison, diagram, table…), keep the slide count. | ✅ yes |
| Just tidy sizing — snap images to the slide cleanly | `fit_image` (Mode B), default box = whole slide | ❌ no |
| Wrap each image in a template's chrome (header/logo/footer) | **`scripts/wrap_images.py`** — one command does it end-to-end (see below). Internally: chrome-only frame deck → `add_image` each picture into the body, fitted. | ❌ no |
| Swap a picture for a new one | `replace_image` (Mode B), keeps the frame | ❌ no |

The recreate path is the only one that yields editable text — everything in the
pixels stays baked in. For bulk transcription, OCR (Tesseract/PaddleOCR) can
seed a first draft, but verify it against the image: OCR mangles diagram labels
and Korean text, so a vision pass is still needed.

**Wrapping images in a template frame — `scripts/wrap_images.py`.** The "wrap in
chrome" path is bundled as one command, so you don't re-orchestrate it each time:

```bash
# from an image-only deck (one picture per slide):
python scripts/wrap_images.py --from-pptx deck.pptx --template samsung \
    --eyebrow "PROJECT NAME" --footer "Confidential" -o out.pptx
# or from a folder of images (natural-sorted slide1, slide2, … slide10):
python scripts/wrap_images.py --images ./imgs --template samsung -o out.pptx
```

It renders a chrome-only frame deck in the chosen template (header bar, brand
marker, page numbers), then fits each image into the body (`--fit contain`
default, `cover` to fill-and-crop). Output slide count = number of images, so
the deck's length is preserved. Use `--eyebrow` for a running label in the bar
and `--footer` for footer text. It does **not** make text editable — for that,
use the recreate path above.

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
- New edit op: add an `op_*` in `scripts/edit_pptx.py` and register it in `OPS`;
  add a validation branch in `scripts/validate_spec.py` (`VALID_OPS` + the
  per-op checks); document it in `references/edit-spec.md`.
- Images: pass a local file path in an `image` slide; a missing path renders a
  labeled placeholder so the build never fails mid-deck.
