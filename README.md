# paper2slides

Turn a scientific paper into an **editable reading-group slide deck** (a `.pptx` with speaker notes) using a local multi-agent pipeline (Claude Code). Deterministic Python scripts handle parsing / building / rendering; the agent does the judgment (deck planning, figure selection, and a **visual-in-the-loop critique** where it renders each slide, looks at it, and fixes overflow/legibility/balance).

Output is a clean **16:9, navy + gold themed** PPTX, uploadable to Google Drive or editable in PowerPoint / Keynote.

Inspired by [PPTAgent](https://github.com/icip-cas/PPTAgent) and [Paper2Poster](https://github.com/Paper2Poster/Paper2Poster), but lighter: no GPU/model stack: figures are extracted with PyMuPDF and slides are built natively with `python-pptx`.

![example](examples/medfact_slides.pdf)

## What you get
- `medfact-style` deck: title → motivation → contributions → method (with figure) → data → **full results table (own slide)** → analysis → takeaways → limitations
- Every slide has **speaker notes**
- Fully editable `.pptx` (native shapes/text, not screenshots)

## Install
System tools (once):
```bash
# macOS
brew install --cask libreoffice    # provides `soffice` (pptx -> pdf/png)
brew install poppler               # provides `pdftoppm`
```
Python:
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Use it as a Claude Code skill (recommended)
Drop the skill where Claude Code finds it:
```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)/skills/paper2slides" ~/.claude/skills/paper2slides   # or cp -R
```
Then in Claude Code:
```
/paper2slides  turn  https://arxiv.org/abs/1706.03762  into reading-group slides
```
or just ask: *"make reading-group slides from paper.pdf"*. The skill ([skills/paper2slides/SKILL.md](skills/paper2slides/SKILL.md)) drives the full pipeline.

## Use it standalone (no agent)
```bash
PY=.venv/bin/python
# 1) parse: PDF path or arXiv id/url -> text + figures + page renders
$PY scripts/parse_paper.py 1706.03762 work
# 2) write work/deck.json yourself (see examples/medfact.deck.json and the schema in SKILL.md)
# 3) build + render
$PY scripts/build_slides.py work/deck.json work/slides.pptx work/assets
$PY scripts/render_slides.py work/slides.pptx work/render 120
# 4) PDF preview
soffice --headless --convert-to pdf work/slides.pptx --outdir work
```

## Repo layout
```
scripts/parse_paper.py    PDF/arXiv -> content.md + assets/figNN.png + pages/page-N.png
scripts/crop.py           crop a vector figure from a page render
scripts/build_slides.py   deck.json -> editable .pptx (themed, 16:9, speaker notes)
scripts/render_slides.py  .pptx -> per-slide PNGs (for the visual critique loop)
skills/paper2slides/SKILL.md   the Claude Code skill (the multi-agent workflow)
examples/medfact.deck.json     a real deck spec
examples/medfact_slides.pdf    the rendered result
```

## deck.json
See the schema and layout list in [SKILL.md](skills/paper2slides/SKILL.md) and the worked example in [examples/medfact.deck.json](examples/medfact.deck.json). Customize colors/fonts at the top of `scripts/build_slides.py`.

## License
MIT
