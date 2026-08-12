#!/usr/bin/env python
"""Propose crop boxes for the figures and tables a paper actually has.

Usage: python find_figures.py <pdf-path | arxiv-id> <outdir> [--pages 2,3,5]

`parse_paper.py` only pulls out embedded rasters, which in most ML papers means
it finds almost nothing: the figures are vector art and the tables are text.
This locates them instead by working from the captions, which are reliable, and
growing a region from the vector paths next to each one.

Writes <outdir>/crops_auto.json in the same schema as crops.json. The boxes are
proposals: skim the page renders, adjust, then save as crops.json.
"""
import sys, os, json, re
import fitz
from parse_paper import resolve

# A caption is "Figure 3." or "Table 1:" and never "Figure 2 presents the ...",
# which is a body sentence that starts the same way.
CAP_RE = re.compile(r"^(Figure|Table)\s+(\d+)\s*[.:]", re.I)
GAP = 0.030        # page-height fraction: how far apart two parts of one figure may sit
PAD = 0.006        # breathing room around the final box
MAX_REACH = 0.60   # a figure never lives further than this from its caption
TOP, BOT = 0.070, 0.935   # live area: below the running head, above the page number
FURNITURE = 0.072         # blocks wholly above this (or below 1-this) are page furniture


def blocks_of(page):
    out = []
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4].strip()
        if text:
            out.append({"rect": fitz.Rect(x0, y0, x1, y1), "text": text})
    return out


def union(rects):
    """Bounding box of rects, computed by coordinate.

    fitz.Rect.__or__ ignores zero-area rects, and a booktabs rule is exactly that:
    a line of zero height. Unioning them the built-in way silently yields nothing.
    """
    return fitz.Rect(min(r.x0 for r in rects), min(r.y0 for r in rects),
                     max(r.x1 for r in rects), max(r.y1 for r in rects))


def same_column(a, b, w, tol=0.06):
    """Two boxes belong to the same column if their x-spans mostly agree."""
    return abs(a.x0 - b.x0) / w < tol and abs(a.x1 - b.x1) / w < tol * 2


def find_on_page(page):
    w, h = page.rect.width, page.rect.height
    blocks = blocks_of(page)
    caps = [b for b in blocks if CAP_RE.match(b["text"])]
    # a booktabs rule is ~0.4pt tall, so requiring both dimensions would discard
    # exactly the lines that delimit a table
    draws = [d["rect"] for d in page.get_drawings()
             if d["rect"].width > 2 or d["rect"].height > 2]
    found = []

    for cap in caps:
        kind, num = CAP_RE.match(cap["text"]).group(1).title(), CAP_RE.match(cap["text"]).group(2)
        above = kind == "Figure"          # figure captions sit under the art, table captions over it
        cx0, cx1 = cap["rect"].x0, cap["rect"].x1
        col = fitz.Rect(cx0 - 0.012 * w, TOP * h, cx1 + 0.012 * w, BOT * h)

        # a figure cannot reach past the next caption sharing its column
        limit = TOP * h if above else BOT * h
        for other in caps:
            if other is cap or not same_column(other["rect"], cap["rect"], w):
                continue
            if above and other["rect"].y1 <= cap["rect"].y0:
                limit = max(limit, other["rect"].y1)
            if not above and other["rect"].y0 >= cap["rect"].y1:
                limit = min(limit, other["rect"].y0)

        # vector paths inside this column, on the correct side, within reach
        cand = []
        for r in draws:
            if r.x0 < col.x0 - 0.01 * w or r.x1 > col.x1 + 0.01 * w:
                continue          # spans past the column, so it is a rule or another figure
            if above and not (limit <= r.y0 and r.y1 <= cap["rect"].y0 + 2):
                continue
            if not above and not (cap["rect"].y1 - 2 <= r.y1 and r.y1 <= limit):
                continue
            if abs((cap["rect"].y0 if above else cap["rect"].y1) - (r.y1 if above else r.y0)) > MAX_REACH * h:
                continue
            cand.append(r)
        if not cand:
            continue

        box = None
        if not above:
            # A booktabs table is a set of rules sharing one x-span, with rows of text
            # between them. Grouping by x-span spans the whole table; growing by
            # proximity would stop at the header rule, because the rows are not drawings.
            groups = {}
            for r in cand:
                groups.setdefault((round(r.x0, 1), round(r.x1, 1)), []).append(r)
            best = None
            for rs in groups.values():
                if len(rs) < 2:
                    continue
                top = min(x.y0 for x in rs)
                if best is None or top < best[0]:
                    best = (top, rs)
            if best:
                box = union(best[1])

        if box is None:
            # grow outward from the path nearest the caption, stopping at a real gap
            cand.sort(key=lambda r: -r.y1 if above else r.y0)
            keep = [cand[0]]
            for r in cand[1:]:
                b = union(keep)
                gap = (b.y0 - r.y1) if above else (r.y0 - b.y1)
                if gap > GAP * h:
                    break
                keep.append(r)
            box = union(keep)

        # pull in adjoining text: axis labels, legends, panel letters like "(a) Qwen2.5-7B".
        # A table's cells already sit inside its rules, so only take text that overlaps.
        for _ in range(0 if not above else 4):
            grown = False
            for b in blocks:
                if b is cap or CAP_RE.match(b["text"]):
                    continue
                r = b["rect"]
                if r.y1 < FURNITURE * h or r.y0 > (1 - FURNITURE) * h:
                    continue          # running head or page number, not part of the figure
                if r.x1 < cx0 - 0.02 * w or r.x0 > cx1 + 0.02 * w:
                    continue
                if above and not (limit <= r.y0 and r.y1 <= cap["rect"].y0):
                    continue
                if not above and not (cap["rect"].y1 <= r.y0 and r.y1 <= limit):
                    continue
                near = (r.y0 <= box.y1 + GAP * h) and (r.y1 >= box.y0 - GAP * h)
                if near and not box.contains(r):
                    # a long paragraph sitting apart from the art is body text
                    if len(b["text"]) > 220 and not box.intersects(r):
                        continue
                    box = union([box, r])
                    grown = True
            if not grown:
                break

        box = fitz.Rect(max(box.x0 - PAD * w, 0), max(box.y0 - PAD * h, 0),
                        min(box.x1 + PAD * w, w), min(box.y1 + PAD * h, h))
        box &= col            # a figure lives in its caption's column and the live area
        if box.is_empty or box.height < 0.02 * h:
            continue
        found.append({
            "name": f"{'fig' if kind == 'Figure' else 'tab'}{num}",
            "page": page.number + 1,
            "box": [round(box.x0 / w, 4), round(box.y0 / h, 4),
                    round(box.x1 / w, 4), round(box.y1 / h, 4)],
            "note": cap["text"][:70].replace("\n", " "),
        })
    return found


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    src, outdir = sys.argv[1], sys.argv[2]
    only = None
    if "--pages" in sys.argv:
        only = {int(x) for x in sys.argv[sys.argv.index("--pages") + 1].split(",")}
    os.makedirs(outdir, exist_ok=True)
    doc = fitz.open(resolve(src, outdir))

    crops = []
    for page in doc:
        if only and page.number + 1 not in only:
            continue
        crops += find_on_page(page)

    spec = {"source": src, "dpi": 320, "crops": crops}
    path = os.path.join(outdir, "crops_auto.json")
    json.dump(spec, open(path, "w"), indent=2)
    embedded = []
    man = os.path.join(outdir, "figures.json")
    if os.path.exists(man):
        embedded = json.load(open(man, encoding="utf-8"))

    for c in crops:
        b = c["box"]
        print(f"  {c['name']:8s} p{c['page']:<3d} [{b[0]:.3f} {b[1]:.3f} {b[2]:.3f} {b[3]:.3f}]  {c['note'][:52]}")
    if embedded:
        pages = sorted({e["page"] for e in embedded})
        print(f"  (plus {len(embedded)} embedded raster(s) already in assets/, from page(s) "
              f"{', '.join(str(p) for p in pages)}; see figures.json)")
    print(f"{len(crops)} proposals -> {path}")
    print("These are starting points: check them against the page renders and adjust before saving as crops.json.")


if __name__ == "__main__":
    main()
