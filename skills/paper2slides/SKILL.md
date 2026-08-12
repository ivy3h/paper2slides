---
name: paper2slides
description: Turn a scientific paper (PDF path or arXiv id/URL) into an editable reading-group slide deck (a .pptx with speaker notes) via a local multi-agent pipeline (parse → plan deck → build → per-slide visual critique → finalize). Use when the user wants to make slides / a talk / a presentation / a deck from a paper.
---

# paper2slides

Convert a paper into a ~12 to 15 slide reading-group talk as an **editable `.pptx`** (clean navy+gold theme, speaker notes on every slide). Deterministic scripts handle parse/build/render; **you (the agent) do the judgment**: deck planning, figure selection, and visual critique. Final `.pptx` is uploadable to Google Drive / openable in PowerPoint/Keynote.

Let `REPO` = this skill's repo root, `PY` = `$REPO/.venv/bin/python`, `W` = a fresh work dir (e.g. `$REPO/work`).

## 0. Prerequisites (once)
- System tools on PATH: **LibreOffice** (`soffice`) and **poppler** (`pdftoppm`). On macOS: `brew install --cask libreoffice && brew install poppler`.
- Python deps: `python3 -m venv "$REPO/.venv" && "$PY" -m pip install -r "$REPO/requirements.txt"`

## 1. Parse the paper
```
"$PY" "$REPO/scripts/parse_paper.py" <pdf-or-arxiv> "$W"
```
→ `W/content.md` (text), `W/pages/page-N.png` (page renders), `W/assets/figNN.png` (embedded figures), `W/figures.json`.

## 2. Choose figures
Let the detector propose boxes first. It works from the captions and the vector paths beside them, so it finds the vector figures and booktabs tables that `parse_paper.py` cannot:
```
"$PY" "$REPO/scripts/find_figures.py" <pdf-or-arxiv> "$W" [--pages 2,3,5]
```
→ `W/crops_auto.json`. On the paper in `examples/` these average **0.91 IoU** against hand-tuned boxes, so nudge them rather than re-deriving. Then read `W/content.md`, skim `W/pages/*.png`, and pick the few highest-value figures/tables (architecture/pipeline diagram, main results table, a key plot). Embedded raster figures are already in `W/assets/`. Crop **key equations** the same way you crop figures: a numbered equation the talk leans on (a loss, an estimator, a gradient) is worth a slide of its own, and cropping keeps the paper's own typography instead of trying to retype maths in `python-pptx`. For **vector** figures (diagrams that didn't extract), crop from a page render:
```
"$PY" "$REPO/scripts/crop.py" "$W/pages/page-3.png" 0.08 0.10 0.95 0.42 "$W/assets/fig_pipeline.png"
```
Once a box is right, **record it in `W/crops.json`** so the deck can be rebuilt from the PDF alone (see `examples/spurious-rewards.crops.json`). Replay or re-cut at another resolution with:
```
"$PY" "$REPO/scripts/apply_crops.py" "$W/crops.json" "$W" [dpi]
```

## 3. Plan the deck → `W/deck.json`
Write `W/deck.json` (schema below). Aim for **12 to 15 slides** in a logical talk flow:
title → motivation/problem → contributions → method (with the pipeline/architecture figure) → data/setup → **main results table on its OWN full slide** (legibility) → analysis/why → key finding(s) → **Conclusions** → (optional) **Future work**.
Rules: bullets **short** (≤5 per slide, ≤~14 words each); **bold** the key phrase in each bullet with `**...**`; write 2 to 4 sentence **speaker notes** per slide; set `figure` to an asset stem in `W/assets`. Keep every number faithful to the paper. **Never use an em dash or en dash anywhere in the deck** (titles, kickers, bullets, captions, notes): recast with a comma, colon, semicolon, parentheses, or a second sentence. Hyphens inside compound words and model names are fine, and use `-` for negative numbers. **No subtitle under the paper title** on the cover, and the title slide carries no one-line summary. Body text is **one size everywhere** (`BODY_SIZE` in build_slides.py); do not vary it per slide. Name the closing slides **Conclusions** and **Future work**: frame what is left open as work to do, not as caveats. **Whenever a slide names an external dataset, method or benchmark** (DeepScaleR, GRPO, MATH-500, a baseline you re-ran), credit it with `cite` and mark the body text so the reader can see something is cited. Pass `cite` as a **list**: entries are auto-numbered and rendered footer-left. Then tag the corresponding words with `[^1]`, `[^2]`, or `[^1,2]` for two at once, which render as superscripts. Format each entry `Name: First author et al. (arXiv:XXXX.XXXXX)`, taking the identifier straight from the paper's own bibliography; if a work has no arXiv id use its venue, e.g. `(Notion Blog, 2025)`. The footer shrinks itself to stay on one line, so four references still fit.

### deck.json schema
```json
{
  "meta": {"title","authors","affiliation","venue","presenter","date"},
  "slides": [
    { "layout": "title|section|bullets|bullets_figure|figure|figure_bullets|table|two_column|matrix|takeaways",
      "title": "…", "kicker": "short label (optional)", "number": "1 (for section)",
      "bullets": ["**bold** then normal text", {"text":"sub point","level":1}],
      "columns": [["left bullets"],["right bullets"]],   // two_column only
      "figure": "figNN or fig_pipeline (stem in assets/)", "figure_caption": "…",
      "cite": ["GRPO: Shao et al. (arXiv:2402.03300)", "…"],  // footer credit, auto-numbered
      "figure_height": 3.0,                              // figure_bullets only (inches)
      "matrix": {"cols":["A","B"], "rows":["R1","R2"],   // matrix only
                 "cells":[["…","…"],["…","…"]], "dead":[[0,1],[1,0]],
                 "height": 2.9, "label_width": 3.4},
      "notes": "speaker notes" }
  ]
}
```
Layouts: `title` & `section` are full navy slides (use meta / number+title); `bullets`; `bullets_figure` (bullets left, figure right, best for tall figures); `figure` (big centered figure); `figure_bullets` (full-width figure band on top, takeaways under it; use for **wide, short** figures and tables, aspect ratio ≳2, which a full-slide `figure` would squash into a thin strip); `table` (full-bleed figure, optionally with a few bullets); `two_column`; `matrix` (case-analysis grid of native shapes; cells listed in `dead` render muted, good for "which branch survives" logic); `takeaways` (gold callout box).

Footer: a gold hairline, the optional `cite` credit bottom-left, and the page number bottom-right, all on one line. No running title.

The `takeaways` callout auto-sizes to its bullet count, so short conclusion lists do not leave a half-empty gold box.

## 4. Lint, then build + render
Check the spec before you render it. Errors are things that render wrong or mislead (a dash, a `[^N]` with no matching reference, a missing figure, a slide with no notes); warnings are worth a look.
```
"$PY" "$REPO/scripts/lint_deck.py" "$W/deck.json" "$W/assets"
```
```
"$PY" "$REPO/scripts/build_slides.py" "$W/deck.json" "$W/slides.pptx" "$W/assets"
"$PY" "$REPO/scripts/render_slides.py" "$W/slides.pptx" "$W/render" 120
```
`build_slides.py` warns on stderr when a text block needs more room than its box, which is the overflow the critique loop used to catch by eye. If you change anything in `scripts/`, run `"$PY" "$REPO/scripts/selftest.py"`: it builds the shipped example and asserts the things that have broken before.

## 5. Visual critique loop (multi-agent)
The mechanical failures are caught for you now: `lint_deck.py` covers dashes, dangling citations, missing notes and unresolved figures, and `build_slides.py` warns when a text block will not fit its box. What is left needs judgment, so **read `W/render/slide-N.png` yourself**, slide by slide, looking for: an illegible table (give it its own full slide or summarize the key numbers), a bad title wrap, a figure cropped through an axis label, a sub-bullet that reads as though it belongs to the wrong parent, left/right imbalance. Apply fixes to `deck.json`, rebuild, re-render, go again (≥2 passes). For a very long deck you can fan out one agent per slide with the Workflow tool, but for a normal 15 to 25 slide talk reading them inline is faster and cheaper. **Re-verify any table you retype against the source PDF**: it is the easiest thing to get wrong.

## 6. Deliver
```
soffice --headless --convert-to pdf "$W/slides.pptx" --outdir "$W"   # PDF preview
```
Hand the user `W/slides.pptx` (editable, with speaker notes) + the PDF.

## Customizing the look
Edit the color constants and layout helpers at the top of `scripts/build_slides.py` (`NAVY`, `GOLD`, fonts, sizes). 16:9 by default (`SW,SH`).
