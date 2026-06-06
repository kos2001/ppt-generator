# ppt-generator

A [Claude Code](https://claude.com/claude-code) skill that turns a topic,
outline, notes, or document into a complete, **editable PowerPoint (`.pptx`)**
deck — rendered in a visual template of your choice.

It separates **what** goes on each slide (a JSON *presentation spec*) from
**how** it looks (a *template*), so the same content can be re-rendered in any
template by changing one field. Inspired by the content → layout → render
pipeline of [Presenton](https://github.com/presenton/presenton), but fully
local: no server, no API, native `.pptx` via [python-pptx](https://python-pptx.readthedocs.io/).

## Features

- **6 templates** — `corporate`, `minimal`, `dark`, `vibrant`, `academic`,
  `report` (black header bar, page numbers, optional "Confidential" marker).
- **13 layouts** — title, section, bullets, content, two_column, comparison,
  metrics, quote, image, table, chart, **diagram**, closing.
- **Editable diagrams** — `process`, `cycle`, `hierarchy`, `pyramid`, `funnel`,
  `timeline`, drawn as native PowerPoint shapes (not flattened images), so every
  box stays editable after generation.
- **Native charts & tables** that remain editable in PowerPoint.
- Restyle a whole deck by changing one `template` field.

## Requirements

```bash
pip install python-pptx Pillow
```

## Quick start

```bash
# 1. Write a presentation spec (see assets/spec.example.json)
# 2. Validate it
python3 scripts/validate_spec.py my-deck.spec.json
# 3. Render to .pptx
python3 scripts/build_pptx.py my-deck.spec.json -o my-deck.pptx
```

Minimal spec:

```json
{
  "template": "corporate",
  "title": "Launch Plan",
  "slides": [
    {"layout": "title", "title": "Launch Plan", "subtitle": "Q1 rollout"},
    {"layout": "bullets", "title": "Goals",
     "bullets": ["Ship to 3 regions", "99.9% uptime", "Onboard 50 teams"]},
    {"layout": "closing", "title": "Thank You"}
  ]
}
```

## Repository layout

| Path | What |
|------|------|
| `SKILL.md` | The skill definition and workflow (entry point for Claude Code). |
| `scripts/build_pptx.py` | Renderer: spec JSON + template → `.pptx`. |
| `scripts/templates.py` | Template (theme) definitions. |
| `scripts/validate_spec.py` | Spec validator. |
| `references/` | Spec schema, layout catalog, template catalog. |
| `assets/spec.example.json` | Full worked example exercising every layout. |
| `decks/` | Example generated decks and their specs. |
| `ppt-generator.skill` | Packaged skill, installable into Claude Code. |

## Install as a Claude Code skill

Copy this folder into your Claude Code skills directory:

```bash
cp -R ppt-generator ~/.claude/skills/ppt-generator
```

or install the packaged `ppt-generator.skill` file.

## Content style

Slide titles are noun phrases (not sentences); body content is short bulleted
noun-phrase fragments rather than prose. See `SKILL.md` for the full design
guidance.

## License

MIT
