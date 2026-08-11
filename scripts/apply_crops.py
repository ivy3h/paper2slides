#!/usr/bin/env python
"""Re-cut every figure a deck uses, from a recorded crops.json.

Usage: python apply_crops.py crops.json <workdir> [dpi]

`crop.py` is for exploring: you eyeball a page render and try a box. Once a box
is right, record it here so the deck can be rebuilt from the PDF alone, at any
resolution, by anyone. Without this the crop coordinates live only in shell
history and the deck is not reproducible.

crops.json:
{
  "source": "2506.10947",          # arXiv id, or a path to the PDF
  "dpi": 320,
  "crops": [
    {"name": "fig_headline", "page": 2, "box": [0.118, 0.074, 0.858, 0.337],
     "note": "Figure 1"}
  ]
}
box is [x0, y0, x1, y1] as fractions of the page, top-left origin.
"""
import sys, os, json, io
import fitz
from PIL import Image
from parse_paper import resolve


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    spec = json.load(open(sys.argv[1], encoding="utf-8"))
    outdir = sys.argv[2]
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else int(spec.get("dpi", 320))

    assets = os.path.join(outdir, "assets")
    os.makedirs(assets, exist_ok=True)
    doc = fitz.open(resolve(str(spec["source"]), outdir))

    # Render each needed page once, then cut boxes with exactly the arithmetic
    # crop.py uses, so exploring with crop.py and replaying here cannot disagree.
    pages = {}
    for n in sorted({c["page"] for c in spec["crops"]}):
        pages[n] = Image.open(io.BytesIO(doc[n - 1].get_pixmap(dpi=dpi).tobytes("png")))

    for c in spec["crops"]:
        im = pages[c["page"]]
        W, H = im.size
        x0, y0, x1, y1 = c["box"]
        out = os.path.join(assets, c["name"] + ".png")
        cut = im.crop((int(x0 * W), int(y0 * H), int(x1 * W), int(y1 * H)))
        cut.save(out)
        print(f"  {c['name']:26s} p{c['page']:<3d} {cut.size[0]}x{cut.size[1]}  {c.get('note', '')}")

    print(f"{len(spec['crops'])} crops at {dpi} dpi -> {assets}")


if __name__ == "__main__":
    main()
