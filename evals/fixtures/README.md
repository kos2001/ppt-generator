# Eval fixtures (not version-controlled)

Evals 3-5 in `../evals.json` exercise the Mode B (edit-an-existing-deck) path,
which needs a real **mixed deck** as input: a deck that is mostly native
template slides with a band of *foreign full-bleed image slides* dropped in
between them, and footer page numbers that no longer match the physical slide
count. That is the exact mess the tidy workflow (`resize_pptx_com.py` →
`--template/--box` fit, `--chrome`, `--renumber`) is designed to fix.

These fixtures are **binary and not tracked** (`*.pptx` here is gitignored) —
deck files are outputs, kept out of the repo. To run evals 3-5, drop matching
decks here using the names `evals.json` expects:

| eval | fixture file | shape |
|------|--------------|-------|
| 3, 4 | `semiconductor-outlook.orig-29.pptx` | 29 slides, standard 13.3×7.5in page, ~17 foreign full-bleed image slides interleaved |
| 5    | `Claude_Cowork_No-Code_Automation.orig.pptx` | 35 slides, oversized 17.8×10in page, ~17 un-chromed full-bleed image slides |

Any deck with the same shape works — the assertions check structure (slide
count preserved, sequential `i / N` numbering, consistent bar heights), not
specific content. Grade an output deck with `python evals/grade_edit.py
<output.pptx> --expect-slides <N>`.
