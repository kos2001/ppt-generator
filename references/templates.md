# Template Catalog

A template is a complete visual identity — palette + fonts + accent treatment —
applied uniformly across all layouts. Set it once via the spec's top-level
`template` field; every slide restyles automatically. The same content can be
re-rendered in any template by changing that one field.

Choose by audience and tone, not personal taste. When unsure, `corporate` is
the safe default.

| Template | Look | Best for |
|----------|------|----------|
| `corporate` | Navy + warm orange on white, sans-serif | Business reviews, sales, exec updates, board decks. The safe default. |
| `minimal` | Black on white, lots of whitespace, thin accents | Design-conscious audiences, product/portfolio, when content should dominate. |
| `dark` | Dark slate background, mint accent, high contrast | Tech/product keynotes, demos, engineering talks, on-stage in a dark room. |
| `vibrant` | Purple + coral, colored panels, bold | Startup pitches, marketing, launches, internal hype. Energetic, less formal. |
| `academic` | Serif (Georgia), muted browns, brick-red accent | Lectures, research talks, reports, anything text-dense and scholarly. |
| `report` | Black header bar with white title, no divider line, clean white body, page numbers | Formal corporate / internal reports, confidential-style documents. |

## Picking a template

- **Formality:** academic / corporate (high) → minimal → vibrant (low).
- **Room:** presenting on a large screen in a dark room → `dark` reads best.
- **Density:** lots of text per slide → `academic` (serif + supports density)
  handles it most gracefully; `minimal` is the worst fit for dense slides.
- **Brand energy:** a launch or pitch that should feel exciting → `vibrant`.

## Fonts

Templates name common fonts (Calibri, Helvetica Neue, Arial, Verdana, Georgia).
PowerPoint substitutes a near-match if a font isn't installed on the viewer's
machine, so decks stay portable. To match a specific brand, edit the
`heading_font` / `body_font` values in `scripts/templates.py`.

## Adding or customizing a template

Templates live in `scripts/templates.py` as plain dicts. To create a brand
theme, copy an existing entry, rename the key, and adjust:

- `primary` — titles, section backgrounds, table headers
- `accent` — the single pop color: bars, metric values, highlights
- `secondary` — eyebrows and sublabels
- `background` / `surface` — slide and card fills
- `text` / `text_muted` / `on_primary` — body, captions, text on colored fills
- `heading_font` / `body_font` and the `*_pt` sizes

Keep every key present; the renderer reads the full set. Pick one accent color
and use it sparingly — restraint is what makes a custom theme look designed.
