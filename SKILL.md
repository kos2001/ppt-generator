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
for image fitting; PyMuPDF rasterizes previews. Check and install if missing:

```bash
python3 -c "import pptx, PIL, pymupdf" 2>/dev/null || pip install python-pptx Pillow pymupdf
```

For the visual check (step 5) you also need LibreOffice (`soffice`):
macOS `brew install --cask libreoffice`, Debian/Ubuntu
`apt install libreoffice-impress`. If it isn't available, skip step 5 and say
so — the validator's fit check still runs.

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

`corporate` (business default) · `report` (formal internal/exec report: black
header bar, "Confidential" label — set `header_label` to change it) ·
`minimal` (design-forward) · `dark` (tech/keynote) · `vibrant`
(pitch/marketing) · `academic` (lectures/research).

### 3. Design the slides as a spec

Write a JSON spec (schema in `references/spec-schema.md`, layouts in
`references/layouts.md`, full worked example in `assets/spec.example.json`).

The craft is in **slide design**, not just dumping text:
- **One idea per slide.** If a slide has two messages, split it.
- **Titles follow the deck's `title_style`** (top-level spec field):
  - `"topic"` (default) — **titles are noun phrases, never sentences.** End
    every title on a noun — "Revenue Growth" or "18% Revenue Growth", never
    "Revenue grew 18%". A title labels the slide; it does not make a claim
    with a verb. Avoid imperative titles too ("Pick a Template" → "Template
    Choice").
  - `"assertion"` — every content-slide title states the slide's **takeaway**
    as one complete claim, ≤ 2 lines, no trailing period: "Churn fell 18%
    after the onboarding redesign"; in Korean, 개조식 결론 "자동 해결률
    61%로 4개 분기 연속 상승". The body is the *evidence* for that claim (a
    chart, metrics, a diagram), so don't repeat the claim in the `subtitle` —
    use the subtitle for scope or source. Title, section and closing slides
    stay noun-phrase labels; give an agenda slide `"title_style": "topic"`.
    Read the titles alone, in order: they should tell the whole story.
  - Use `"assertion"` when the user asks for headline / action / 결론형 titles.
    For decks meant to persuade or drive a decision (exec reports, proposals,
    research findings) you can suggest it in your outcome summary — studies of
    assertion-evidence slides (Garner & Alley, 2013) found better
    comprehension and recall than topic headings with bullets — but keep the
    default unless the user opts in.
  - The validator flags titles that don't match the style: sentence-like
    titles in a `topic` deck, label-like titles (< 4 words) in an `assertion`
    deck.
- **No full sentences in the body.** Write body content as short noun phrases
  and fragments, not narrative prose. Drop articles and trailing verbs:
  "Enterprise segment outperformed plan" → "Enterprise above plan". This keeps
  slides scannable and stops them from becoming a document.
- **Every body point is a bullet.** Separate each idea with its own bullet, and
  default to the `bullets` / `two_column` / `comparison` layouts. Avoid the
  prose `content` layout — reserve it for a genuine exception, and even then
  break the points into bullets. (A single `subtitle` lead-in line under the
  title is the one place a full sentence is fine. Use `subtitle` on every
  layout; `lead` is an older bullets-only field — don't use both.)
- **Pick the layout that fits the job** — numbers → `metrics` or `chart`;
  the one number that matters → `big_number`; contrast → `comparison`;
  parallel options/pillars → `cards`; 3–6 capabilities or benefits →
  `icon_grid` (icons from `references/icons.md`, Korean aliases like
  "보안" work); prioritization → `matrix` (2×2);
  relationships/flow → `diagram`; steps or ranked items → `bullets` with
  `"numbered": true`; the talk's structure → `agenda`; a human beat →
  `quote`. Don't default everything to one layout.
- **Vary the rhythm.** Never stack three same-layout slides in a row. Use
  `section` dividers to break a longer deck into parts.
- **Keep slides sparse.** ~5 bullets max (validator warns above 6), each
  bullet ≤ 2 lines, ≤4 metrics with short values ("$4.2M", "94%"), ≤~6×8
  tables. Slides support the speaker; they aren't the document.
- **Put the explanation in `notes`.** Every content slide gets speaker notes
  (`"notes": "..."`, 1–4 sentences): the detail, numbers' context, and the
  transition. Text on screen that duplicates the narration hurts recall; a
  sparse slide plus notes is the evidence-backed split.
- **Make charts say one thing.** Single series + `highlight` the category
  that matters (others turn grey); put units in `number_format`, never in
  the data (`"series": [2.3, 1.1]`, `"number_format": "\"$\"0.0\"M\""`).
  Use the `subtitle` to state the takeaway.
- **Never invent numbers.** Every figure on a slide must come from the user's
  material; if a layout wants a number you don't have, use a different
  layout. Simple arithmetic on given figures (a share of a total, a
  before/after delta) is fine — show the inputs in the speaker notes.
- **Fill the slide or change the layout.** Three short bullets leave most of a
  slide empty; numbers belong in `metrics`, a before/after in `comparison`,
  a set of parallel items in a `diagram`.

A typical arc: `title` → agenda/`section` → `metrics`/`chart` (situation) →
alternating content layouts (argument) → `comparison`/`quote` (turn) →
`content` (forward look) → `closing`. See `references/layouts.md` for more.

Write the spec to a file, e.g. `deck.spec.json`.

### 4. Validate, then render

```bash
python3 scripts/validate_spec.py deck.spec.json     # structure + lint + dry-run fit
python3 scripts/build_pptx.py   deck.spec.json -o deck.pptx
```

The validator checks three things:
- **Structure** (ERROR): unknown layouts, missing fields, chart series whose
  length doesn't match the categories, non-numeric chart values.
- **Density & rhythm** (WARN): too many bullets, bullets that wrap past two
  lines, three same-layout slides in a row, oversized tables, long metric
  values, theme colors below WCAG contrast.
- **Fit** (dry-run of the real renderer): text it must shrink below 85%
  (`FIT`, WARN) and text that overflows even at 70% (`OVERFLOW`, ERROR).
  Overlong `bullets`/`table` slides are split onto "(계속)/(cont.)"
  continuation slides instead (`SPLIT`, WARN) — see `references/layouts.md`.

Fix every ERROR. Treat WARNs as edits to make: an `OVERFLOW` or heavy `FIT`
means the slide carries too much — split it or cut words, don't just accept
the shrink; an `ORPHAN` title leaves one short word alone on its last line —
shorten or rephrase it; a `SPARSE` bullets slide fills under 35% of the
body — add substance or move it to `cards`, `metrics`, `big_number` or a
`diagram`. The renderer prints the same fit report after it
saves, or `Fit: all text fits…` when there is nothing to report. The fit
check is an estimate: the visual check in step 5 is what catches awkward
wrapping it can't predict.

The renderer fits text automatically: titles shrink, then wrap to two lines
(the body moves down to make room); body text, cards, table cells and diagram
labels scale down uniformly to fit their boxes. That is a safety net, not a
license to overfill.

### 5. Look at the slides, then iterate

Render the deck to images and **look at them** (Read the PNG):

```bash
python3 scripts/render_preview.py deck.pptx        # -> deck_preview/contact_sheet.png
```

Open `contact_sheet.png` (all slides in a numbered grid; individual
`slide-NN.png` files are next to it for a closer look) and check, in order:
1. Text cut off, overflowing a box, or overlapping another element
2. Awkward wrapping — a title or bullet with a single orphaned word
3. Empty-looking or overcrowded slides; three similar slides in a row
4. Charts: readable labels, the highlighted bar is the one that matters
5. Contrast: anything hard to read on its background

`render_preview.py` also measures the rendered pixels and prints a `LAYOUT`
line for any slide whose bottom third or more is empty (also saved as
`layout_metrics.json`). Title, section and quote slides are sparse on
purpose; elsewhere grow the content or switch layout.

Before delivering, write down each issue you found as **blocker** (text cut
off, overlap, unreadable, wrong number), **major** (overflow risk, orphaned
title word, empty or overcrowded slide, chart highlighting the wrong thing)
or **minor** (spacing, wording polish). Do not hand off while any blocker or
major remains — fix and re-render. Mention unresolved minors in your summary.

Fix issues in the spec, re-render, and re-check (usually one or two rounds).
LibreOffice substitutes uninstalled fonts, so judge layout and density, not
the exact line breaks.

Then state the real outcome: how many slides, which template, where the file
is, what you checked visually (or that you could not, if LibreOffice is
missing). Offer concrete next steps: adjust length, swap the template
(one-field change → re-render), tighten wording, add a chart, etc.

To **restyle** an existing deck into another look, change only the top-level
`template` field in the spec and re-render — content is untouched.

## Korean and other CJK text

Handled automatically: every run gets an East Asian font (`a:ea`, default
**Malgun Gothic** — ships with Windows and Office; macOS falls back to Apple SD
Gothic Neo) plus `lang="ko-KR"`, and Korean paragraphs break between words
(어절) instead of mid-word. Korean runs are never italicized (PowerPoint would
slant only the Latin digits in the line). Fitting accounts for Hangul being
~2× the width of Latin letters, so Korean decks need fewer words per bullet. To use a brand or
modern font, set top-level `"ea_font": "Pretendard"` (or `"font"` to change the
Latin font too) — warn the user it must be installed on the presenting machine.
Write titles and bullets in the user's language.

## Output conventions

- Save specs and `.pptx` files where the user is working, or `claudedocs/` if
  there's no obvious home. Name by topic (e.g. `q4-review.pptx`).
- The output is a standard editable `.pptx` — charts and tables remain
  editable in PowerPoint, so the user can refine after you hand it off.

## Notes on extending

- New template: copy an entry in `scripts/templates.py`, adjust palette/fonts,
  then run the validator on any spec using it — it checks the theme's
  contrast ratios.
- New layout: add a `layout_*` function in `scripts/build_pptx.py` and register
  it in the `LAYOUTS` dict; document it in `references/layouts.md`.
- Images: pass a local file path in an `image` slide; a missing path renders a
  labeled placeholder so the build never fails mid-deck.
- After changing the renderer, run the regression tests from the skill root:
  `python3 -m unittest discover -s tests` (builds every template, checks fit,
  Korean fonts, bullets, chart styling, notes, and validator behavior).
