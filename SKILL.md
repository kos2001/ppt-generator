---
name: ppt-generator
description: >-
  Create, read, edit, or diagram PowerPoint (.pptx) decks — use WHENEVER a user
  wants slides made, changed, or diagrammed. CREATE a full editable deck from a
  topic, outline, notes, or document ("make slides about X", "발표 자료
  만들어줘", "투자자용 10장 피치", "turn this doc into a deck"). EDIT an existing
  .pptx in place, keeping its design — replace text (formatting preserved),
  tables, chart data, images; add/delete/reorder slides ("이 파일 수정해줘",
  "fix the title", "swap the logo"). DIAGRAM as
  native editable shapes (not images): process, cycle, org chart/hierarchy,
  pyramid, funnel, timeline, branching flowchart ("다이어그램/플로우차트/조직도/
  순서도 그려줘"). FIT images to a template box without distortion ("템플릿
  크기에 맞춰"). RESTYLE into another template (corporate, minimal, dark,
  vibrant, academic, samsung, report). READ/EXTRACT text, tables, chart data,
  images ("내용 추출", "이 ppt 읽어줘"). Handles ordinary .pptx via python-pptx
  and corporate DRM/EDM decks via authorized PowerPoint COM. Produces a real,
  fully editable .pptx — not a description of one.
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

**How to pick — infer it, don't interrogate.** Almost every request resolves on
its own; read the cue and go:

- An existing `.pptx` named with *change* language — "수정/고쳐/바꿔/정리/통일",
  "fix", "edit", "update", "swap", "clean up" — is **Mode B**. Don't ask what to
  change; the Mode B workflow inspects the deck and acts (see below).
- A topic, outline, notes, or source doc with *make* language — "만들어줘",
  "발표자료", "deck", "slides about…", "turn this into a deck" — is **Mode A**.

Only one case is genuinely 50/50 and worth a single quick question: a `.pptx` is
attached **and** the ask is "make a deck from this" — which can mean *edit this
file* or *use it as source material for a fresh deck*. There, ask once:

> "(A) 이걸 소재로 **새 덱을 생성**할까요, 아니면 (B) 이 **파일을 직접 수정**할까요?"

Why even that one matters: generating re-renders the whole deck from a spec, so
it *replaces* the original design, while editing preserves every untouched shape
and bit of formatting. Guessing wrong here silently discards exactly what the
user meant to keep — so resolve this single fork, but don't manufacture a
question when the cue above already settles it.

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

Optional, only for visual QA (`scripts/thumbnail.py`): desktop PowerPoint with
`pywin32` — it exports each slide straight to PNG via COM, so there is no
external converter to install. The reading/editing scripts (`extract_pptx.py`,
`inspect_pptx.py`, `edit_pptx.py`, `ooxml.py`) need only `python-pptx` + Pillow.

> **Windows COM path (DRM/EDM decks):** the `*_com.py` scripts
> (`extract_pptx_com.py`, `resize_pptx_com.py`) drive desktop PowerPoint via
> COM and require `pywin32` plus an installed, signed-in PowerPoint. Install
> with `pip install pywin32`. Verify with `py -c "import win32com.client"`.
> See the **Authorized COM path** section under Mode B for usage.

## Mode A — Generate a deck from a layout

### 1. Confirm the brief — at most one quick question

A deck is expensive to redo, so the few things that shape *every* slide are
worth getting right. But most of them are already fixed for this user, so the
confirm collapses to a single lean question — and often to none:

- **Template / look** — **always `samsung`** for this user. Never ask; never
  offer the catalog. (The other templates in `references/templates.md` exist for
  restyling, not for a pre-build choice.)
- **Length** — default **~10 slides** (tight 8–14). Only a default; easy to redo.
- **Audience & setting** — assume a general professional audience on a big
  screen unless the request implies otherwise.
- **Topic, key message, source material** — read from the request itself. If the
  user gave a document/notes, extract the structure from it; don't invent facts.

**So:** if the request already implies topic + roughly who/how long, **just
build** and name the choices you made ("samsung·10장·일반 청중 기준으로 만들었어요
— 길이/대상 바꿀 점 있으면 말씀해 주세요") so the user can course-correct in one
reply. Only when length *and* audience are both genuinely unclear, ask **one**
consolidated question with the defaults pre-filled as a one-tap answer — e.g.
"몇 장 정도로, 누구 대상으로 만들까요? (기본: 10장·일반 청중)" — then build.

The bar for asking is high on purpose: a single optional confirm is fine, but
template choice, a multi-question intake, or re-asking what the request already
says all just add friction. When the essentials are present, momentum beats
confirmation.

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
contact-sheet grid of every slide and inspect it (needs desktop PowerPoint +
`pywin32`; PowerPoint exports each slide to PNG via COM):

```bash
python3 scripts/thumbnail.py deck.pptx          # -> deck.grid.png (+ per-slide PNGs)
```

Open the grid to catch text cutoff, overflow, or clashing colors. If PowerPoint
or `pywin32` isn't available, the script says so — then tell the user to open the
file in PowerPoint/Keynote and offer concrete next steps: adjust length, swap the
template (one-field change → re-render), tighten wording, add a chart, etc.

To **restyle** an existing deck into another look, change only the top-level
`template` field in the spec and re-render — content is untouched.

## Mode B — Work on an existing deck

This mode reads an existing `.pptx` and modifies it in place — keeping every
untouched shape, image, and bit of formatting. Use it whenever the user has a
file already and wants it changed rather than rebuilt.

**When the ask is vague ("이 파일 수정해 줘", "fix this deck") — don't ask what
to change. Inspect first, then act.** A bare "fix it" is an invitation to find
the problems, not a prompt to interview the user. Run `inspect_pptx.py` (and a
`thumbnail.py` grid when the issue might be visual) to see the real state.

**First, always confirm the *actual physical slide count*** — the number of
slides `inspect_pptx.py` reports — and compare it to what the footer page numbers
claim (e.g. a "N / 12" stamp). They diverge often: image-only slides exported
from another tool get inserted *between* content slides, so a deck whose footers
read "/ 12" may physically hold 29 slides. Treat the physical count from
`inspect_pptx.py` as ground truth — it drives `renumber_pages` (which restamps
each slide as physical `position / total`), tells you whether foreign image
slides snuck in, and prevents off-by-N edits when you address shapes by slide
index. Note the real total in your first report. Then
sort what you find into two buckets — and treat them differently:

- **Safe, reversible fixes → just do them, now.** These only adjust *form*, never
  destroy content, so there's nothing to second-guess: brand/fit foreign or
  off-template slides to match the deck (full-bleed images with no chrome),
  renumber stale page numbers to physical `position / total` (`renumber_pages`
  does exactly this — don't talk yourself out of it), refit distorted or
  overflowing images, correct an obviously wrong value. Apply every fix in this
  bucket before you consider asking anything.
- **Destructive or judgment calls → flag them *after*, don't freeze on them.**
  Deleting slides, dropping content, or rewriting wording can't be undone and may
  not match intent, so these you raise rather than assume. But raising them is the
  *last* step, not a gate on the safe fixes above — never withhold a reversible
  improvement just because some adjacent decision needs the user.

The trap to avoid: a slide can be *off-template* (foreign styling — a safe
branding fix) **and** *off-topic* (content that may not belong — a delete
judgment call) at once. Brand it regardless; that's reversible and on-template
either way. Then, separately, note the content concern: "이미지 슬라이드 N장을
템플릿에 맞춰 브랜딩하고 번호를 다시 매겼어요. 다만 그 12장은 주제와 달라 보이는데
— 뺄까요, 둘까요?" That mirrors the right instinct: fix what's clearly fixable,
surface what genuinely needs a human call. Reserve an up-front question only when
the deck has **no clear defect to fix** and the intent is truly underdetermined.
The default is momentum: inspect, make every safe fix, report, then ask.

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

**Authorized COM path (Windows + PowerPoint).** When the user can already open
and edit the file manually in desktop PowerPoint on their corporate PC — i.e.
the DRM/EDM grants them edit rights — you can drive that *same authorized
PowerPoint* via COM automation (pywin32) instead of extracting a plaintext copy.
PowerPoint decrypts in the signed-in session, the script edits, and `Save()`
re-encrypts on write. This is **not** a DRM bypass; it uses the user's own
rights. Two bundled COM scripts (Windows only; `pip install pywin32`):

- `scripts/extract_pptx_com.py` — read text + speaker notes from a DRM file.
- `scripts/resize_pptx_com.py` — **resize pictures** in a DRM file, including
  fitting them to a template's content box. Resizing changes only frame geometry
  and never reads the image pixels, so it works even when DRM blocks image
  *extraction* (if manual resize works, this does too). Run it in order:

  ```bash
  python scripts/resize_pptx_com.py deck.pptx --check    # can COM even open it?
  python scripts/resize_pptx_com.py deck.pptx --list     # pictures + sizes
  python scripts/resize_pptx_com.py deck.pptx --template # fit all into template box
  ```

  `--template [NAME]` (default `samsung`) fits each picture into the template
  content box; `--box WxH`, `--scale F`, `--width/--height` are the other modes;
  `--fit contain|stretch`. In-place saves back up the original first.

  `--chrome [NAME]` — **add the template's header/footer chrome** to a DRM deck.
  `--template` only moves picture *frames*; `--chrome` inserts *new native
  shapes* (header bar, `SAMSUNG DS` brand marker, page numbers, optional
  `--eyebrow`/`--footer`) on every slide, mirroring the bar-style chrome in
  `build_pptx.py`. PowerPoint decrypts in the authorized session, the shapes are
  added, and `Save()` re-encrypts — so a DRM deck the user can edit manually
  gets the same look as a natively rendered one (the COM equivalent of
  `wrap_images.py`, which can't run on DRM because it needs python-pptx). It is
  idempotent — a re-run removes its prior chrome before re-adding. Typical
  image-deck flow is two passes: fit the images, then add chrome:

  ```bash
  python scripts/resize_pptx_com.py deck.pptx --template samsung --out tmp.pptx
  python scripts/resize_pptx_com.py tmp.pptx  --chrome  samsung \
      --eyebrow "Project" --footer "Confidential" --out final.pptx
  ```

  **Fitting and chrome go together.** `--template` alone leaves a picture
  centered on a *bare white slide* — it is not "applying the template," only
  reframing. To make an image slide actually look like the template you must
  follow the fit pass with a `--chrome` pass; do both whenever the goal is an
  on-brand slide, and never offer the fit by itself as if it were the whole job.

  `--slides LIST` (with `--chrome`) — **brand only the listed slides**, e.g.
  `--slides 3,5` or `--slides 3-5` (1-based; `--slide N` also works for one).
  This is the **DRM equivalent of the mixed-deck flow** below: in a deck that is
  mostly native template slides with a few foreign image slides dropped in,
  chrome *only* the foreign slides so the native ones don't get a second header
  bar. Without it, `--chrome` stamps every slide and double-chromes the native
  ones. Page numbers always use each slide's physical 1-based position over the
  full deck. If the deck's existing numbers deliberately skip the image slides
  (a common generation artifact), add `--no-numbers` so the branded slides match
  that scheme instead of colliding with a neighbor's number:

  ```bash
  # mixed DRM deck: fit images, then chrome ONLY the foreign image slides 3 & 5
  python scripts/resize_pptx_com.py deck.pptx --template samsung --out tmp.pptx
  python scripts/resize_pptx_com.py tmp.pptx  --chrome samsung --slides 3,5 \
      --no-numbers --footer "Confidential" --out final.pptx
  ```

  Chrome geometry scales with the deck's page height (relative to a 7.5in-tall
  16:9 baseline), so the bar, brand marker, and page number match natively-built
  chrome at any page size — including oversized 16:9 decks (e.g. 17.8×10in). When
  fitting full-bleed images on such a deck, match the box to the slides that are
  already fitted (measure one with `--list`) rather than the default template box,
  so all content slides line up; `--slides LIST` also restricts `--template`/resize
  to just the full-bleed slides without disturbing already-fitted ones.

  `--renumber [NAME]` — **renumber the whole deck `1/N .. N/N`** by physical
  position via COM. **This is the standard final step after ANY edit** (add /
  delete / tidy / mixed-deck branding): always leave the deck numbered
  sequentially over its real total — never leave a stale logical count like
  "n / 12" on a deck that now physically holds 29 slides, and don't leave image
  slides unnumbered. It updates an existing `N / M` page-number textbox in place
  (native slides) and adds one bottom-right where none exists (title, closing,
  freshly-chromed image slides). Prefer this over `--no-numbers`: rather than
  matching a skip-scheme, unify the entire deck.

  **The mixed-deck tidy recipe (three `--out` passes).** This is the workflow
  that comes up most: a deck that is mostly native template slides with a band of
  foreign full-bleed image slides dropped in, footers reading a stale count.
  First classify — which slides are full-bleed images with no chrome (these need
  fit + chrome) versus already-fitted/native ones (leave untouched). `--list`
  shows picture sizes: a picture filling the whole slide is full-bleed; one inset
  at an offset is already fitted. Then fit **only** the full-bleed slides, chrome
  **only** the same set (no numbers yet), and renumber the whole deck last:

  ```bash
  # full-bleed slides here are 1-3,13-24,34-35; fit them into the SAME box the
  # already-fitted slides use (measure one with --list) so all content lines up
  python scripts/resize_pptx_com.py deck.pptx --slides 1-3,13-24,34-35 \
      --box 34.6x19.3 --left 5.27 --top 3.89 --out a.pptx
  python scripts/resize_pptx_com.py a.pptx --chrome samsung --slides 1-3,13-24,34-35 \
      --no-numbers --out b.pptx
  python scripts/resize_pptx_com.py b.pptx --renumber samsung --out final.pptx
  ```

  On a deck with no already-fitted slides to match, `--template samsung` (which
  reserves the header area automatically) is fine for the fit pass instead of an
  explicit `--box`. Always render a `thumbnail.py` grid of the result and confirm
  the new bars line up with the native ones before promoting it over the target.

  `--delete-slides LIST` — **delete slides** (e.g. `10-16,18-27`, 1-based) via
  COM, the DRM-safe counterpart to `edit_pptx.py`'s `delete_slide`. **Deleting is
  destructive — confirm with the user first** (see the Mode B buckets above); a
  "수정/fix" request is *not* a delete request. Deletes in descending order so
  indices don't shift mid-pass; renumber afterward.

  **COM save hazard:** a lingering `POWERPNT.EXE` that still holds the file open
  can silently re-save the pre-edit version over your output. After a COM save,
  verify the on-disk slide count (`zipfile` namelist) and check
  `tasklist | grep -i powerpnt` for zombies; prefer `--out` chains over repeated
  in-place saves.

  Some strict DRM products allow manual editing but block automation/COM
  specifically — then `--check` fails at `Open()`. That block is intended; stop
  there, don't work around it.

**Block diagrams on any deck — `scripts/diagram_com.py`.** The COM counterpart
to the native `diagram` layout. python-pptx can draw diagrams but can't open DRM
files; this draws them as real, editable PowerPoint shapes (rounded boxes,
arrows, connectors) onto a slide through the authorized session — so a DRM deck
gets native diagrams, not a flattened image. Seven types: `process`, `cycle`,
`hierarchy`, `pyramid`, `funnel`, `timeline`, and `flowchart` (an arbitrary
branching graph, auto-laid out `LR`/`TD` with crossing minimization — share the
layout logic with `build_pptx.py` via `flowchart_layout.py`).

```bash
python scripts/diagram_com.py --demo --out diagrams.pptx          # showcase all 7
python scripts/diagram_com.py --in deck.pptx --slide 4 --type flowchart \
    --direction TD --title "처리 흐름" \
    --nodes '[{"id":"a","title":"수집"},{"id":"b","title":"분석"}]' \
    --edges '[["a","b"]]'
```

For a non-DRM deck, prefer the native `diagram` layout in a spec (Mode A) — same
seven types, including `flowchart` (see `references/layouts.md`).

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
new text), `add_chrome` (brand a slide with a template's header/footer in place),
`renumber_pages` (re-stamp page numbers over the new total), `delete_shape`, and
slide-level `duplicate_slide` / `delete_slide` / `move_slide`. `duplicate_slide`
copies a slide and its images N times — the primitive for wrapping an image-only
deck in a fixed template (duplicate one chrome-only frame slide per image, then
`add_image` into each; see `references/edit-spec.md`).
`replace_text` is the safest bulk edit — it matches by content, so it is
unaffected by shape indices. Structural ops (duplicate/delete/move slide, delete
shape) shift indices, so order them last or re-inspect between edits.

Images are fitted to a **fixed box** without distortion: `add_image` places into
a template placeholder (`into_shape`) or inch box, `replace_image` keeps the
existing picture's frame, and `fit_image` resizes a picture already on the slide
to a box — defaulting to the whole slide, so it snaps full-page images to fill
the slide template exactly. `add_image`/`fit_image` also accept a `template`
field (e.g. `"template": "samsung"`) that fits the image into that template's
standard **content box** — the region under the header bar and above the
footer/page-number band, sized to the deck's actual slide dimensions — so an
edited image lands exactly where the template wants it. All take a `fit` mode
(`contain` = letterbox inside, `cover` = fill and crop, `stretch`).

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

**Mixed decks — native template slides + foreign image slides.** A common case
is a deck that is *mostly* templated (rendered by `build_pptx.py`, or already
on-brand) but has a few full-bleed image slides dropped in — NotebookLM/Gamma
exports, translated slides, a screenshot. Those inserted slides have no chrome,
so they look foreign, and the page numbers no longer match the new slide count.
You want to brand just those slides while keeping every native slide and the
slide count intact. `wrap_images.py` won't do (it rebuilds the whole deck from
images, discarding the native slides). Pick the path by whether the file is
protected:

- **Unprotected `.pptx`** → edit in place with the `add_chrome` op, one foreign
  slide at a time, then unify numbering once (below).
- **DRM-protected** (python-pptx can't open it) → use the COM path with
  `--chrome --slides 3,5` to brand *only* the foreign slides (see the
  **Authorized COM path** section). Bare `--chrome` would stamp every slide and
  double-chrome the native ones, so the slide list is required here.

For the unprotected path:

```jsonc
// edits.json — brand slides 3-5, then renumber the whole deck
{"op": "fit_image", "slide": 3, "shape": 0, "template": "samsung", "fit": "contain"},
{"op": "add_chrome", "slide": 3, "template": "samsung", "eyebrow": "solutions", "title": "..."},
// …repeat fit_image + add_chrome for each foreign slide…
{"op": "renumber_pages"}   // run last: re-stamps every "N / M" to physical order over the new total
```

`fit_image` (same `template`) pulls each picture into the box *under* the header
band; `add_chrome` then draws the matching header bar, brand marker, and page
number — pixel-identical to a native render because it reuses `build_pptx.py`'s
own chrome code. See `references/edit-spec.md` for the full field list. (If a
foreign image already has its own baked-in title, the template's bar title and
the image's title will both show — keep the bar title short, match it, or omit
`title`.)

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
