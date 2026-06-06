---
name: ppt-generator
description: >-
  Generate complete, editable PowerPoint (.pptx) presentations from a topic,
  outline, notes, or source document, rendered in a chosen visual template.
  Use this skill WHENEVER the user wants to create, build, make, or draft a
  slide deck / presentation / PPT / PPTX / slides / "발표 자료" / pitch deck /
  keynote — even if they only describe the topic ("make slides about X",
  "turn this doc into a deck", "I need a 10-slide pitch for investors") and
  don't mention a tool or file format. Also use when the user wants to restyle
  an existing deck into a different template/theme, or asks for a specific
  look (corporate, minimal, dark, vibrant, academic). Produces a real, fully
  editable .pptx via python-pptx — not a description of one.
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
Don't claim it "looks great" — you can't see it render. If LibreOffice is
available (`soffice`), you can generate thumbnails to inspect:

```bash
soffice --headless --convert-to pdf --outdir /tmp deck.pptx   # then view the PDF
```

Otherwise, tell the user to open it in PowerPoint/Keynote and offer concrete
next steps: adjust length, swap the template (one-field change → re-render),
tighten wording, add a chart, etc.

To **restyle** an existing deck into another look, change only the top-level
`template` field in the spec and re-render — content is untouched.

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
