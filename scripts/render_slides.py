#!/usr/bin/env python
"""Render a PPTX to per-slide PNGs via LibreOffice + pdftoppm.

Usage: python render_slides.py deck.pptx out_dir [dpi]
Produces out_dir/slide-1.png, slide-2.png, ...  and prints the count.
"""
import sys, os, subprocess, tempfile, glob, shutil

def main():
    pptx = os.path.abspath(sys.argv[1])
    out_dir = os.path.abspath(sys.argv[2])
    dpi = sys.argv[3] if len(sys.argv) > 3 else "110"
    os.makedirs(out_dir, exist_ok=True)
    for f in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(f)
    with tempfile.TemporaryDirectory() as td:
        env = os.environ.copy(); env["LC_ALL"] = "en_US.UTF-8"; env["LANG"] = "en_US.UTF-8"
        with tempfile.TemporaryDirectory() as userdir:
            subprocess.run(
                ["soffice", "--headless", "--norestore", "--nolockcheck",
                 f"-env:UserInstallation=file://{userdir}",
                 "--convert-to", "pdf", pptx, "--outdir", td],
                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env,
            )
        pdfs = glob.glob(os.path.join(td, "*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice produced no PDF")
        subprocess.run(["pdftoppm", "-png", "-r", dpi, pdfs[0], os.path.join(out_dir, "slide")],
                       check=True)
        # pdftoppm names slide-1.png or slide-01.png; normalize to slide-N.png
        pages = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
        for p in pages:
            base = os.path.basename(p)
            num = base.replace("slide-", "").replace(".png", "").lstrip("0") or "0"
            tgt = os.path.join(out_dir, f"slide-{int(num)}.png")
            if p != tgt:
                shutil.move(p, tgt)
        print(f"{len(pages)} slides -> {out_dir}")

if __name__ == "__main__":
    main()
