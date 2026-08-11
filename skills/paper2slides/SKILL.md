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
Read `W/content.md`; skim `W/pages/*.png`. Pick the few highest-value figures/tables (architecture/pipeline diagram, main results table, a key plot). Embedded raster figures are already in `W/assets/`. For **vector** figures (diagrams that didn't extract), crop from a page render:
```
"$PY" "$REPO/scripts/crop.py" "$W/pages/page-3.png" 0.08 0.10 0.95 0.42 "$W/assets/fig_pipeline.png"
```

## 3. Plan the deck → `W/deck.json`
Write `W/deck.json` (schema below). Aim for **12 to 15 slides** in a logical talk flow:
title → motivation/problem → contributions → method (with the pipeline/architecture figure) → data/setup → **main results table on its OWN full slide** (legibility) → analysis/why → key finding(s) → takeaways → (optional) limitations.
Rules: bullets **short** (≤5 per slide, ≤~14 words each); **bold** the key phrase in each bullet with `**...**`; write 2 to 4 sentence **speaker notes** per slide; set `figure` to an asset stem in `W/assets`. Keep every number faithful to the paper. **Never use an em dash or en dash anywhere in the deck** (titles, kickers, bullets, captions, notes): recast with a comma, colon, semicolon, parentheses, or a second sentence. Hyphens inside compound words and model names are fine, and use `-` for negative numbers.

### deck.json schema
```json
{
  "meta": {"title","subtitle","authors","affiliation","venue","presenter","date"},
  "slides": [
    { "layout": "title|section|bullets|bullets_figure|figure|figure_bullets|table|two_column|matrix|takeaways",
      "title": "…", "kicker": "short label (optional)", "number": "1 (for section)",
      "bullets": ["**bold** then normal text", {"text":"sub point","level":1}],
      "columns": [["left bullets"],["right bullets"]],   // two_column only
      "figure": "figNN or fig_pipeline (stem in assets/)", "figure_caption": "…",
      "figure_height": 3.0,                              // figure_bullets only (inches)
      "matrix": {"cols":["A","B"], "rows":["R1","R2"],   // matrix only
                 "cells":[["…","…"],["…","…"]], "dead":[[0,1],[1,0]],
                 "height": 2.9, "label_width": 3.4},
      "notes": "speaker notes" }
  ]
}
```
Layouts: `title` & `section` are full navy slides (use meta / number+title); `bullets`; `bullets_figure` (bullets left, figure right, best for tall figures); `figure` (big centered figure); `figure_bullets` (full-width figure band on top, takeaways under it; use for **wide, short** figures and tables, aspect ratio ≳2, which a full-slide `figure` would squash into a thin strip); `table` (full-bleed figure, optionally with a few bullets); `two_column`; `matrix` (case-analysis grid of native shapes; cells listed in `dead` render muted, good for "which branch survives" logic); `takeaways` (gold callout box).

Footer: a gold hairline plus the page number, bottom-right. No running title.

## 4. Build + render
```
"$PY" "$REPO/scripts/build_slides.py" "$W/deck.json" "$W/slides.pptx" "$W/assets"
"$PY" "$REPO/scripts/render_slides.py" "$W/slides.pptx" "$W/render" 120
```

## 5. Visual critique loop (multi-agent)
Fan out **one agent per slide** (use the Workflow tool) to read `W/render/slide-N.png` + that slide's `deck.json` entry and return concrete fixes: text overflow / running under the footer, too many or too-long bullets, an illegible table (give it its own full slide or summarize key numbers), bad title wrap, imbalance, missing notes. Then apply fixes to `deck.json`, rebuild, re-render. Repeat until clean (≥2 passes). **Re-verify any tables against the source PDF**: agents sometimes mis-transcribe numbers.

## 6. Deliver
```
soffice --headless --convert-to pdf "$W/slides.pptx" --outdir "$W"   # PDF preview
```
Hand the user `W/slides.pptx` (editable, with speaker notes) + the PDF.

## Customizing the look
Edit the color constants and layout helpers at the top of `scripts/build_slides.py` (`NAVY`, `GOLD`, fonts, sizes). 16:9 by default (`SW,SH`).
