#!/usr/bin/env python
"""Parse a paper into text + figures + page renders for the slide pipeline.

Usage: python parse_paper.py <pdf-path | arxiv-id | arxiv-url> <outdir>

Outputs:
  <outdir>/content.md          full text (page-marked)
  <outdir>/pages/page-N.png    full-page renders (to view / crop vector figures)
  <outdir>/assets/figNN.png    embedded raster figures (>= ~300x300)
  <outdir>/figures.json        manifest of extracted figures
"""
import sys, os, json, re, io, urllib.request
import fitz  # PyMuPDF
from PIL import Image


def resolve(src, outdir):
    if os.path.exists(src):
        return src
    m = re.search(r"(\d{4}\.\d{4,5})(v\d+)?", src)
    if "arxiv.org" in src or m:
        aid = m.group(0) if m else src
        url = f"https://arxiv.org/pdf/{aid}"
        dst = os.path.join(outdir, "paper.pdf")
        print(f"downloading {url}")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as r, open(dst, "wb") as f:
            f.write(r.read())
        return dst
    raise SystemExit(f"not found and not an arXiv id/url: {src}")


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    src, outdir = sys.argv[1], sys.argv[2]
    pages_dir = os.path.join(outdir, "pages")
    assets_dir = os.path.join(outdir, "assets")
    for d in (outdir, pages_dir, assets_dir):
        os.makedirs(d, exist_ok=True)

    pdf = resolve(src, outdir)
    doc = fitz.open(pdf)
    md, manifest, seen = [], [], set()
    for i, page in enumerate(doc, 1):
        md.append(f"\n\n## [page {i}]\n\n" + page.get_text("text"))
        page.get_pixmap(dpi=140).save(os.path.join(pages_dir, f"page-{i}.png"))
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                ext = doc.extract_image(xref)
            except Exception:
                continue
            w, h = ext.get("width", 0), ext.get("height", 0)
            if w * h < 90000:  # skip logos / icons / rules
                continue
            name = f"fig{len(manifest) + 1:02d}.png"
            try:
                Image.open(io.BytesIO(ext["image"])).convert("RGB").save(os.path.join(assets_dir, name))
            except Exception:
                continue
            manifest.append({"stem": name[:-4], "page": i, "w": w, "h": h})

    open(os.path.join(outdir, "content.md"), "w").write("".join(md))
    json.dump(manifest, open(os.path.join(outdir, "figures.json"), "w"), indent=2)
    print(f"pages: {len(doc)} -> {pages_dir}")
    print(f"embedded figures: {len(manifest)} -> {assets_dir}")
    print("text -> content.md ; manifest -> figures.json")
    print("NOTE: vector diagrams may not extract as embedded images — crop them from pages/page-N.png with crop.py")


if __name__ == "__main__":
    main()
