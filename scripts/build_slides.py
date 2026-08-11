#!/usr/bin/env python
"""Build a 16:9 editable PPTX from a deck.json spec.

Usage: python build_slides.py deck.json medfact_slides.pptx [assets_dir]

deck.json schema (see sample_deck.json):
{
  "meta": {title, subtitle, authors, affiliation, venue, presenter, date},
  "slides": [ { "layout": <type>, "title":..., "kicker":..., "number":...,
                "bullets": [str | {"text":str,"level":0}], "columns": [[..],[..]],
                "figure": "<manifest id or filename>", "figure_caption": str,
                "notes": str } ]
}
Layouts: title | section | bullets | bullets_figure | figure | table | two_column | takeaways
Bold spans inside text use **markdown** style: "**Over-criticism** hurts precision".
"""
import sys, json, os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

NAVY  = RGBColor(0x0B, 0x25, 0x45)
NAVY2 = RGBColor(0x16, 0x33, 0x5F)
GOLD  = RGBColor(0xE8, 0xA3, 0x3D)
GOLDD = RGBColor(0xB9, 0x78, 0x1B)
INK   = RGBColor(0x1A, 0x24, 0x38)
MUTED = RGBColor(0x5A, 0x66, 0x80)
LIGHT = RGBColor(0xF4, 0xF7, 0xFC)
GOLDSOFT = RGBColor(0xFB, 0xF1, 0xDE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT  = "Helvetica Neue"

SW, SH = Inches(13.333), Inches(7.5)

def _emu_in(x):  # Inches helper accepting float
    return Inches(x)

def parse_bold(s):
    """Split '**bold** normal' into [(text, is_bold), ...]."""
    parts, bold = [], False
    for chunk in s.split("**"):
        if chunk:
            parts.append((chunk, bold))
        bold = not bold
    return parts or [("", False)]

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def rect(slide, x, y, w, h, color, line=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp

def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(0); tf.margin_top = tf.margin_bottom = Pt(0)
    return tb, tf

def style_runs(para, text, size, color, bold=False, italic=False, font=FONT):
    """Add runs to a paragraph, honoring **bold** spans (bold+navy)."""
    for chunk, is_b in parse_bold(text):
        r = para.add_run(); r.text = chunk
        r.font.size = Pt(size); r.font.name = font
        r.font.bold = bold or is_b
        r.font.italic = italic
        r.font.color.rgb = NAVY if (is_b and not bold) else color

def add_para(tf, text, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT,
             space_after=6, space_before=0, line=1.04, first=False):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after); p.space_before = Pt(space_before)
    try: p.line_spacing = line
    except Exception: pass
    style_runs(p, text, size, color, bold, italic)
    return p

RULE_Y = 7.18          # gold hairline above the footer
FOOT_Y = RULE_Y + 0.08  # text sits clear of the rule (a 10pt line is ~0.17in tall)

def footer(slide, num, total):
    rect(slide, 0, RULE_Y, 13.333, 0.022, GOLD)
    tb, tf = textbox(slide, 11.0, FOOT_Y, 1.78, 0.24)
    add_para(tf, f"{num} / {total}", 10, MUTED, align=PP_ALIGN.RIGHT, first=True)

def add_notes(slide, notes):
    if notes:
        slide.notes_slide.notes_text_frame.text = notes

def fit(path, box_w, box_h):
    """Return (w,h) in inches fitting image inside box preserving aspect."""
    try:
        iw, ih = Image.open(path).size
    except Exception:
        return box_w, box_h
    ar = iw / ih
    w, h = box_w, box_w / ar
    if h > box_h:
        h, w = box_h, box_h * ar
    return w, h

def add_figure(slide, path, cx, cy, box_w, box_h, caption=None, cap_h=0.45, cap_size=11.5):
    """Place image centered in box (cx,cy = box top-left), caption below."""
    if not path or not os.path.exists(path):
        return
    reserve = cap_h if caption else 0
    w, h = fit(path, box_w, box_h - reserve)
    left = cx + (box_w - w) / 2
    top = cy + (box_h - reserve - h) / 2
    pic = slide.shapes.add_picture(path, Inches(left), Inches(top), Inches(w), Inches(h))
    pic.line.color.rgb = RGBColor(0xD7, 0xDE, 0xEA); pic.line.width = Pt(0.75)
    if caption:
        # keep the caption attached to the image instead of pinned to the bottom of the box
        cap_top = min(top + h + 0.06, cy + box_h - cap_h)
        tb, tf = textbox(slide, cx, cap_top, box_w, cap_h)
        add_para(tf, caption, cap_size, MUTED, italic=True, align=PP_ALIGN.CENTER, first=True, line=1.05)

def slide_title(slide, title, kicker=None):
    if kicker:
        tb, tf = textbox(slide, 0.55, 0.34, 12.2, 0.32)
        p = add_para(tf, kicker.upper(), 12.5, GOLDD, bold=True, first=True)
    tb, tf = textbox(slide, 0.55, 0.62 if kicker else 0.45, 12.2, 0.9)
    add_para(tf, title, 27, NAVY, bold=True, first=True, line=1.0)
    rect(slide, 0.57, 1.45 if kicker else 1.28, 1.6, 0.045, GOLD)
    return 1.75 if kicker else 1.58  # content top y

def hanging(para, indent_in):
    """Hanging indent: wrapped lines align under the text, not under the bullet glyph."""
    pPr = para._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent_in))))
    pPr.set("indent", str(-int(Inches(indent_in))))

def bullets_block(slide, bullets, x, y, w, h, base_size=19):
    n = len(bullets)
    size = base_size if n <= 5 else (17 if n <= 7 else 15)
    tb, tf = textbox(slide, x, y, w, h)
    first = True
    for b in bullets:
        if isinstance(b, dict):
            text, level = b.get("text", ""), b.get("level", 0)
        else:
            text, level = str(b), 0
        if level == 0:
            p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
            p.space_after = Pt(7); p.space_before = Pt(2); p.line_spacing = 1.06
            r = p.add_run(); r.text = "▪  "; r.font.size = Pt(size); r.font.color.rgb = GOLD; r.font.bold = True; r.font.name = FONT
            style_runs(p, text, size, INK)
            hanging(p, size / 68.0)  # ≈ width of "▪  " at this point size
        else:
            p = tf.add_paragraph(); p.space_after = Pt(4); p.line_spacing = 1.04
            r = p.add_run(); r.text = "        –  "; r.font.size = Pt(size-2); r.font.color.rgb = MUTED; r.font.name = FONT
            style_runs(p, text, size-2, INK)
            hanging(p, (size - 2) / 26.0)  # ≈ width of the indented dash prefix
        first = False
    return tb

def build(deck, out_pptx, assets):
    prs = Presentation(); prs.slide_width = SW; prs.slide_height = SH
    blank = prs.slide_layouts[6]
    meta = deck.get("meta", {})
    slides = deck["slides"]
    total = len(slides)

    def figpath(fid):
        if not fid: return None
        for cand in (fid, fid + ".png", os.path.join(assets, fid), os.path.join(assets, fid + ".png")):
            if os.path.exists(cand): return cand
        return None

    for i, s in enumerate(slides, 1):
        layout = s.get("layout", "bullets")
        sl = prs.slides.add_slide(blank)
        notes = s.get("notes", "")

        if layout == "title":
            set_bg(sl, NAVY)
            rect(sl, 0, 0, 13.333, 0.16, GOLD)
            tb, tf = textbox(sl, 0.9, 0.5, 11.5, 0.4)
            add_para(tf, (meta.get("venue") or "").upper(), 14, GOLD, bold=True, first=True)
            tb, tf = textbox(sl, 0.9, 2.0, 11.5, 2.6, anchor=MSO_ANCHOR.TOP)
            add_para(tf, meta.get("title", ""), 40, WHITE, bold=True, first=True, line=1.03)
            if meta.get("subtitle"):
                add_para(tf, meta["subtitle"], 20, RGBColor(0xC8,0xD4,0xEA), space_before=10)
            rect(sl, 0.92, 4.7, 2.2, 0.05, GOLD)
            tb, tf = textbox(sl, 0.9, 4.95, 11.5, 1.6)
            add_para(tf, meta.get("authors", ""), 17, RGBColor(0xDF,0xE7,0xF4), bold=True, first=True)
            if meta.get("affiliation"):
                add_para(tf, meta["affiliation"], 14, RGBColor(0xAA,0xB9,0xD4), italic=True, space_before=2)
            foot = "   ·   ".join(x for x in [meta.get("presenter",""), meta.get("date","")] if x)
            if foot:
                add_para(tf, foot, 13.5, GOLD, space_before=14)
            add_notes(sl, notes); continue

        if layout == "section":
            set_bg(sl, NAVY)
            rect(sl, 0, 0, 0.22, 7.5, GOLD)
            tb, tf = textbox(sl, 1.1, 2.4, 11, 2.6, anchor=MSO_ANCHOR.MIDDLE)
            if s.get("number"):
                add_para(tf, str(s["number"]), 60, GOLD, bold=True, first=True, line=1.0)
            add_para(tf, s.get("title",""), 36, WHITE, bold=True, space_before=4, line=1.02)
            if s.get("subtitle"):
                add_para(tf, s["subtitle"], 18, RGBColor(0xC8,0xD4,0xEA), space_before=8)
            add_notes(sl, notes); continue

        # ---- content slides (white) ----
        set_bg(sl, WHITE)
        cy = slide_title(sl, s.get("title",""), s.get("kicker"))

        if layout == "bullets":
            bullets_block(sl, s.get("bullets", []), 0.6, cy+0.05, 12.1, 5.1, base_size=20)

        elif layout == "bullets_figure":
            bullets_block(sl, s.get("bullets", []), 0.6, cy+0.05, 6.6, 5.0, base_size=18)
            add_figure(sl, figpath(s.get("figure")), 7.45, cy+0.0, 5.3, 5.05, s.get("figure_caption"))

        elif layout == "figure":
            add_figure(sl, figpath(s.get("figure")), 0.7, cy+0.0, 11.9, 5.05, s.get("figure_caption"))

        elif layout == "figure_bullets":
            # wide, short figures (tables / multi-panel strips) waste a full-slide `figure`:
            # give them a full-width band on top and put the takeaways underneath.
            fh = float(s.get("figure_height", 3.2))
            add_figure(sl, figpath(s.get("figure")), 0.6, cy+0.0, 12.13, fh,
                       s.get("figure_caption"), cap_h=0.32, cap_size=10.5)
            rest = 5.05 - fh - 0.15
            bullets_block(sl, s.get("bullets", []), 0.6, cy + fh + 0.15, 12.13, max(rest, 0.6),
                          base_size=float(s.get("bullet_size", 17)))

        elif layout == "table":
            if s.get("bullets"):
                bullets_block(sl, s["bullets"], 0.6, cy+0.05, 4.6, 5.0, base_size=17)
                add_figure(sl, figpath(s.get("figure")), 5.4, cy+0.0, 7.4, 5.05, s.get("figure_caption"))
            else:
                # full-bleed results table: maximize footprint so digits stay legible.
                box_top = 1.55  # start just under the title's gold rule
                box_h = 7.13 - box_top  # extend down to just above the footer rule
                add_figure(sl, figpath(s.get("figure")), 0.3, box_top, 12.73, box_h,
                           s.get("figure_caption"), cap_h=0.3, cap_size=10.5)

        elif layout == "matrix":
            # case-analysis grid (native shapes, still editable): row/col headers + cells,
            # cells listed in "dead" render muted (the branch that gets killed).
            m = s.get("matrix", {})
            mcols, mrows = m.get("cols", []), m.get("rows", [])
            mcells = m.get("cells", [])
            deadset = {tuple(d) for d in m.get("dead", [])}
            gx, gy, gw = 0.6, cy + 0.05, 12.13
            lab_w = float(m.get("label_width", 3.4))
            grid_h = float(m.get("height", 2.9))
            hh = 0.5
            cw = (gw - lab_w) / max(len(mcols), 1)
            rh = (grid_h - hh) / max(len(mrows), 1)
            for cj, ch in enumerate(mcols):  # NB: do not shadow the outer slide counter `i`
                tb, tf = textbox(sl, gx + lab_w + cj * cw + 0.1, gy, cw - 0.2, hh, anchor=MSO_ANCHOR.MIDDLE)
                add_para(tf, ch, 14, GOLDD, bold=True, align=PP_ALIGN.CENTER, first=True, line=1.05)
            for ri, rhd in enumerate(mrows):
                top = gy + hh + ri * rh
                tb, tf = textbox(sl, gx, top + 0.08, lab_w - 0.25, rh - 0.16, anchor=MSO_ANCHOR.MIDDLE)
                add_para(tf, rhd, 15, NAVY, bold=True, first=True, line=1.06)
                for cj in range(len(mcols)):
                    is_dead = (ri, cj) in deadset
                    rect(sl, gx + lab_w + cj * cw + 0.06, top + 0.06, cw - 0.12, rh - 0.12,
                         RGBColor(0xEF, 0xF1, 0xF5) if is_dead else GOLDSOFT,
                         line=RGBColor(0xD7, 0xDE, 0xEA) if is_dead else RGBColor(0xEC, 0xCF, 0x97))
                    txt = mcells[ri][cj] if ri < len(mcells) and cj < len(mcells[ri]) else ""
                    tb, tf = textbox(sl, gx + lab_w + cj * cw + 0.24, top + 0.14, cw - 0.48, rh - 0.28,
                                     anchor=MSO_ANCHOR.MIDDLE)
                    add_para(tf, txt, 15, MUTED if is_dead else INK, align=PP_ALIGN.CENTER,
                             first=True, line=1.08, italic=is_dead)
            if s.get("bullets"):
                bullets_block(sl, s["bullets"], 0.6, gy + grid_h + 0.22, 12.13,
                              max(5.05 - grid_h - 0.27, 0.6), base_size=17)

        elif layout == "two_column":
            cols = s.get("columns", [[], []])
            bullets_block(sl, cols[0] if len(cols)>0 else [], 0.6, cy+0.05, 5.9, 5.0, base_size=18)
            rect(sl, 6.66, cy+0.1, 0.012, 4.7, RGBColor(0xD7,0xDE,0xEA))
            bullets_block(sl, cols[1] if len(cols)>1 else [], 6.95, cy+0.05, 5.8, 5.0, base_size=18)

        elif layout == "takeaways":
            box = rect(sl, 0.6, cy+0.2, 12.13, 4.6, GOLDSOFT)
            box.line.color.rgb = RGBColor(0xEC,0xCF,0x97); box.line.width = Pt(1)
            rect(sl, 0.6, cy+0.2, 0.13, 4.6, GOLD)
            bullets_block(sl, s.get("bullets", []), 1.05, cy+0.45, 11.4, 4.1, base_size=21)

        else:
            bullets_block(sl, s.get("bullets", []), 0.6, cy+0.05, 12.1, 5.1, base_size=20)

        footer(sl, i, total)
        add_notes(sl, notes)

    prs.save(out_pptx)
    print(f"saved {out_pptx}  ({total} slides)")

if __name__ == "__main__":
    deck_path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "slides.pptx"
    assets = sys.argv[3] if len(sys.argv) > 3 else "assets"
    deck = json.load(open(deck_path))
    build(deck, out, assets)
