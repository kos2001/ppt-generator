# Template Catalog

A template is a complete visual identity — palette + fonts + accent treatment —
applied uniformly across all layouts. Set it once via the spec's top-level
`template` field; every slide restyles automatically. The same content can be
re-rendered in any template by changing that one field.

The renderer default is `standard` (this skill's standing default — see
SKILL.md). When the user names a template, use exactly that one; among the
generic looks, `corporate` is the safe choice for business decks.

| Template | Look | Best for |
|----------|------|----------|
| `corporate` | Navy + warm orange on white, sans-serif | Business reviews, sales, exec updates, board decks. The safe generic choice. |
| `minimal` | Black on white, lots of whitespace, thin accents | Design-conscious audiences, product/portfolio, when content should dominate. |
| `dark` | Dark slate background, mint accent, high contrast | Tech/product keynotes, demos, engineering talks, on-stage in a dark room. |
| `vibrant` | Purple + coral, colored panels, bold | Startup pitches, marketing, launches, internal hype. Energetic, less formal. |
| `academic` | Serif (Georgia), muted browns, brick-red accent | Lectures, research talks, reports, anything text-dense and scholarly. |
| `report` | Black header bar with white title, no divider line, clean white body, page numbers | Formal corporate / internal reports, confidential-style documents. |
| `standard` | Black header bar, top-right "Confidential" classification marker, blue accents, white body, Malgun Gothic; footer shows only the page number | Korean corporate / internal brand decks following the standard internal-report style. |

## Picking a template

- **Formality:** academic / corporate (high) → minimal → vibrant (low).
- **Room:** presenting on a large screen in a dark room → `dark` reads best.
- **Density:** lots of text per slide → `academic` (serif + supports density)
  handles it most gracefully; `minimal` is the worst fit for dense slides.
- **Brand energy:** a launch or pitch that should feel exciting → `vibrant`.

## Fonts

Templates name common fonts (Calibri, Helvetica Neue, Arial, Verdana, Georgia).
PowerPoint substitutes a near-match if a font isn't installed on the viewer's
machine, so decks stay portable. Korean/CJK characters use the separate East
Asian font (`ea_font`, default Malgun Gothic), which the renderer writes into
every run and into the deck's theme fonts — so text the user adds later in
PowerPoint matches too. To match a specific brand, edit the
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

Optional keys:
- `bleed` / `on_bleed` — fill and text for full-bleed `section` / `closing`
  slides (default `primary` / `on_primary`; `dark` keeps them dark)
- `on_accent` — text on accent-filled headers (e.g. comparison's right card)
- `chart_colors` — series palette (default: primary, accent, secondary, …)
- `grid` — chart gridline color; `muted_fill` — non-highlighted bars
- `ea_font` — East Asian (Korean/CJK) font; default Malgun Gothic

## Contrast

The validator checks every theme against WCAG contrast: 4.5:1 for body text
(text, muted text, text on cards, text on primary fills) and 3:1 for large
headings (title color, heading on accent). All bundled templates pass. When
you add or edit a theme, run the validator and darken any color it flags.
