# Slide Review Taxonomy

The checklist for the visual review of any deck — one you generated (Mode A,
step 5) or one you edited (Mode B). It follows the **SlideAudit** taxonomy of
slide design flaws (Zhang, Chen, Zhong & Wobbrock, *SlideAudit: A Dataset and
Taxonomy for Automated Evaluation of Presentation Slides*, UIST '25,
[arXiv:2508.03630](https://arxiv.org/abs/2508.03630)): 19 single-slide
categories in four groups, derived with 17 design experts and used to
annotate 2,400 slides.

**Why a fixed taxonomy:** in that study, prompting models with the full
taxonomy raised flaw-detection F1 for every model tested (e.g. GPT-4o
0.476 → 0.655, Claude 0.519 → 0.612). The best results also gave the model a
structured description of the slide's objects next to the image. So review
with both inputs:

```bash
python3 scripts/audit_pptx.py deck.pptx                 # automated checks (tagged with these categories)
python3 scripts/inspect_pptx.py deck.pptx --geometry     # object description: boxes, positions, font sizes
# + the rendered image: thumbnail.py (PowerPoint) or render_preview.py (LibreOffice)
```

Walk each slide through every category below and decide *present / absent*.
Most slides have none or one or two flaws; "no flaw" is a valid answer —
the taxonomy's other benefit was fewer false alarms.

## Composition & Layout

| Category | Look for | Auto |
|---|---|---|
| Lack of Visual Hierarchy | Nothing stands out as most important; title, body and data compete | — |
| Cluttered Layout | Too many elements for the space; no breathing room | — |
| Unbalanced Space or Margin Usage | Content crowded to one side or the top; big empty band; elements hugging edges | `LAYOUT` (render_preview) |
| Misaligned Elements | Edges that almost line up but don't; ragged columns | audit |
| Content Overflow or Truncation | Text past its box or cut off; shapes off the slide | audit + fit report |
| Occluded Content | Text or data hidden under another element | audit |

## Typography

| Category | Look for | Auto |
|---|---|---|
| Illegible Typeface Usage | Decorative/script fonts; fallback fonts that look wrong | — |
| Inappropriate Font Sizing | Text too small for the room (< ~10pt body) or oversized | audit |
| Excessive Text Density | A paragraph on a slide; more than ~6 bullets or ~90 words | audit + validator |
| Inconsistent or Distracting Text Styling | Bold/italic/underline/color used everywhere; many typefaces; Hangul without a pinned font | audit |
| Improper Line or Character Spacing | Lines jammed together or floating apart; letter-spaced body text | — |
| Lack of Text Hierarchy | Headings, subheadings and body look the same | — |

## Color

| Category | Look for | Auto |
|---|---|---|
| Insufficient Contrast | Text below 4.5:1 (3:1 for large/bold) against what is behind it | audit + validator (themes) |
| Overuse or Inconsistent Color | More than one accent; colors that change meaning slide to slide | — |
| Poor or Jarring Color Combinations | Clashing hues, saturated pairs that vibrate | — |

## Imagery

| Category | Look for | Auto |
|---|---|---|
| Irrelevant or Off-topic Visuals | Images or icons that don't support the slide's point | — |
| Low-Quality or Distorted Images | Pixelated (low effective DPI), compression artifacts | audit |
| Improper Image Scaling or Size | Stretched/squashed pictures; images too small to read | audit |
| Inconsistent Visual Style | Flat icons mixed with photos or 3D clip art | — |

Deck-level categories from the full taxonomy (no navigation cues,
inconsistent layout across slides, illogical content flow) and animation
categories are outside this per-slide checklist, but keep them in mind when
you read the contact sheet as a whole.

## Severity

Record each flaw you confirm as:
- **blocker** — content lost or unreadable: overflow/truncation, occlusion,
  contrast far below threshold, a wrong number;
- **major** — clearly hurts reading: excessive density, tiny text, misleading
  hierarchy, distorted image, an empty or overcrowded slide;
- **minor** — polish: near-miss alignment, spacing, style consistency.

Do not hand off a deck with an open blocker or major.
