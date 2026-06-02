#!/usr/bin/env python
"""Crop a region from a page render (use for vector figures that didn't extract).

Usage: python crop.py <page.png> <x0> <y0> <x1> <y1> <out.png>
Coords are fractions of the page (0..1): top-left (x0,y0) to bottom-right (x1,y1).
Example: python crop.py work/pages/page-3.png 0.08 0.10 0.95 0.42 work/assets/fig_pipeline.png
"""
import sys
from PIL import Image

if len(sys.argv) != 7:
    raise SystemExit(__doc__)
src = sys.argv[1]
x0, y0, x1, y1 = map(float, sys.argv[2:6])
out = sys.argv[6]
im = Image.open(src)
W, H = im.size
im.crop((int(x0 * W), int(y0 * H), int(x1 * W), int(y1 * H))).save(out)
print("saved", out)
