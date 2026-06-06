"""Visual templates (themes) for the PPT generator.

Each theme is a self-contained design system: a color palette, font choices,
and a few sizing knobs. The renderer (build_pptx.py) reads these values and
applies them consistently across every slide layout, so picking a different
template restyles the whole deck without touching its content.

Colors are hex strings WITHOUT the leading '#'. Fonts are family names as they
appear in PowerPoint; if a font is not installed on the viewer's machine,
PowerPoint substitutes a close match, so prefer widely available families.

To add a new template, copy an existing entry and adjust the values. Keep every
key present — the renderer expects the full set.
"""

THEMES = {
    # Clean, trustworthy, the safe default for business decks.
    "corporate": {
        "label": "Corporate",
        "description": "Clean professional blue/grey. Safe default for business, sales, and exec decks.",
        "background": "FFFFFF",
        "surface": "EEF3FA",       # card / panel fill
        "primary": "1F3A5F",       # deep navy — titles, bars
        "secondary": "2E75B6",     # mid blue — accents, sublabels
        "accent": "E8833A",        # warm orange — highlights, metric values
        "text": "21262E",          # near-black body text
        "text_muted": "6B7785",    # captions, footers
        "on_primary": "FFFFFF",    # text drawn on top of primary fills
        "heading_font": "Calibri",
        "body_font": "Calibri",
        "title_pt": 40,
        "heading_pt": 28,
        "body_pt": 18,
    },
    # Lots of whitespace, restrained palette, editorial feel.
    "minimal": {
        "label": "Minimal",
        "description": "Black on white, generous whitespace, thin accents. Design-forward, low clutter.",
        "background": "FFFFFF",
        "surface": "F4F4F4",
        "primary": "111111",
        "secondary": "555555",
        "accent": "111111",
        "text": "1A1A1A",
        "text_muted": "8A8A8A",
        "on_primary": "FFFFFF",
        "heading_font": "Helvetica Neue",
        "body_font": "Helvetica Neue",
        "title_pt": 44,
        "heading_pt": 28,
        "body_pt": 18,
    },
    # Dark mode, high contrast, modern/tech product vibe.
    "dark": {
        "label": "Dark",
        "description": "Dark background, high contrast, neon accent. Modern, technical, product/keynote.",
        "background": "0F1419",
        "surface": "1C242E",
        "primary": "FFFFFF",
        "secondary": "9AA7B4",
        "accent": "3DDC97",        # mint green
        "text": "E6EAEE",
        "text_muted": "8A97A4",
        "on_primary": "0F1419",
        "heading_font": "Arial",
        "body_font": "Arial",
        "title_pt": 42,
        "heading_pt": 28,
        "body_pt": 18,
    },
    # Bold, colorful, energetic — startups, marketing, pitches.
    "vibrant": {
        "label": "Vibrant",
        "description": "Bold purple/coral with colored panels. Energetic — pitches, marketing, launches.",
        "background": "FFFFFF",
        "surface": "F3EEFF",
        "primary": "5B2A86",       # purple
        "secondary": "9C27B0",
        "accent": "FF5252",        # coral red
        "text": "2A2233",
        "text_muted": "7A7287",
        "on_primary": "FFFFFF",
        "heading_font": "Verdana",
        "body_font": "Verdana",
        "title_pt": 42,
        "heading_pt": 28,
        "body_pt": 18,
    },
    # Black header bar with white title, no divider line, clean white body.
    # Formal corporate/report style (e.g. confidential internal documents).
    "report": {
        "label": "Report",
        "description": "Black header bar with white title, no divider line, clean white body, page numbers. Formal corporate/internal-report style.",
        "background": "FFFFFF",
        "surface": "F2F2F2",
        "primary": "111111",       # near-black — header bar, section bg, table header
        "secondary": "595959",     # grey — eyebrows, sublabels
        "accent": "2F5597",        # corporate blue — metric values, charts (not a header line)
        "text": "1A1A1A",
        "text_muted": "7F7F7F",
        "on_primary": "FFFFFF",
        "heading_font": "Calibri",
        "body_font": "Calibri",
        "title_pt": 44,
        "heading_pt": 30,
        "body_pt": 22,
        "header_style": "bar",
        "header_label": "Confidential",  # white label at the right of the header bar; "" to disable
    },
    # Serif, muted, dense — lectures, research, reports.
    "academic": {
        "label": "Academic",
        "description": "Serif fonts, muted scholarly palette, supports denser text. Lectures, research, reports.",
        "background": "FBFAF7",
        "surface": "EFEBE2",
        "primary": "3A2E1F",       # dark brown
        "secondary": "7A6A52",
        "accent": "8C2F23",        # brick red
        "text": "2B2620",
        "text_muted": "726B5E",
        "on_primary": "FBFAF7",
        "heading_font": "Georgia",
        "body_font": "Georgia",
        "title_pt": 38,
        "heading_pt": 26,
        "body_pt": 18,
    },
}

DEFAULT_THEME = "corporate"


def get_theme(name):
    """Return the theme dict for `name`, falling back to the default.

    Unknown names fall back rather than raise, so a typo in a spec still
    produces a deck (with a warning printed by the caller).
    """
    return THEMES.get((name or "").strip().lower(), THEMES[DEFAULT_THEME])


def theme_names():
    return list(THEMES.keys())
