#!/usr/bin/env python
"""Check a deck.json against the house style before you build it.

Usage: python lint_deck.py deck.json [assets_dir]

Errors (exit 1) are things that render wrong or mislead: dashes, a dangling
[^N] with no matching reference, a missing figure, a slide with no notes.
Warnings (exit 0) are things worth a look: an overlong bullet, an unreferenced
reference, closing slides that are not named Conclusions / Future work.
"""
import sys, json, os, re

DASHES = "‐‑‒–—―−⁃﹘﹣－"          # every Unicode dash and minus variant
DASH_RE = re.compile(f"[{DASHES}]")
SUP_RE = re.compile(r"\[\^([\d,]+)\]")
BAD_SUP_RE = re.compile(r"\[\^(?![\d,]+\])")   # [^ not followed by digits and a close

LAYOUTS = {"title", "section", "bullets", "bullets_figure", "figure",
           "figure_bullets", "table", "two_column", "matrix", "takeaways"}
FIGURE_LAYOUTS = {"bullets_figure", "figure", "figure_bullets", "table"}

MAX_TOP_BULLETS = 9     # bullets_block stops scaling past this
MAX_BULLET_WORDS = 20
MAX_TITLE_CHARS = 75
MIN_NOTES_CHARS = 40

errors, warnings = [], []


def err(where, msg):
    errors.append(f"{where}: {msg}")


def warn(where, msg):
    warnings.append(f"{where}: {msg}")


def texts_of(slide):
    """Every author-written string on a slide, so no check can miss one."""
    out = []
    for k in ("title", "kicker", "subtitle", "figure_caption", "notes"):
        if slide.get(k):
            out.append(str(slide[k]))
    for b in slide.get("bullets", []):
        out.append(b.get("text", "") if isinstance(b, dict) else str(b))
    for col in slide.get("columns", []):
        for b in col:
            out.append(b.get("text", "") if isinstance(b, dict) else str(b))
    m = slide.get("matrix", {})
    out += [str(x) for x in m.get("cols", [])] + [str(x) for x in m.get("rows", [])]
    for row in m.get("cells", []):
        out += [str(x) for x in row]
    cite = slide.get("cite")
    out += [str(c) for c in cite] if isinstance(cite, (list, tuple)) else ([str(cite)] if cite else [])
    return out


def body_texts_of(slide):
    """Just the visible body text, i.e. where [^N] markers are allowed to live."""
    out = []
    for k in ("title", "kicker", "figure_caption"):
        if slide.get(k):
            out.append(str(slide[k]))
    for b in slide.get("bullets", []):
        out.append(b.get("text", "") if isinstance(b, dict) else str(b))
    for col in slide.get("columns", []):
        for b in col:
            out.append(b.get("text", "") if isinstance(b, dict) else str(b))
    return out


def bullets_of(slide):
    out = list(slide.get("bullets", []))
    for col in slide.get("columns", []):
        out += list(col)
    return out


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    path = sys.argv[1]
    assets = sys.argv[2] if len(sys.argv) > 2 else None
    deck = json.load(open(path, encoding="utf-8"))

    meta = deck.get("meta", {})
    if meta.get("subtitle"):
        err("meta", "title slide carries no subtitle; drop meta.subtitle")
    if meta.get("short_title"):
        warn("meta", "short_title is unused since the running title was removed")
    for s in texts_of({"notes": meta.get("title", "")}) + [str(v) for v in meta.values()]:
        if DASH_RE.search(s):
            err("meta", f"em/en dash in {s[:60]!r}")

    slides = deck.get("slides", [])
    if not slides:
        err("deck", "no slides")

    for i, sl in enumerate(slides, 1):
        w = f"slide {i}"
        layout = sl.get("layout", "bullets")
        if layout not in LAYOUTS:
            err(w, f"unknown layout {layout!r}")

        for t in texts_of(sl):
            if DASH_RE.search(t):
                err(w, f"em/en dash in {t[:70]!r}")
            if BAD_SUP_RE.search(t):
                err(w, f"malformed citation marker in {t[:70]!r}")

        # citations: every [^N] must resolve, and a numbered list is required to number them
        marks = set()
        for t in body_texts_of(sl):
            for m in SUP_RE.findall(t):
                marks.update(int(x) for x in m.split(",") if x)
        cite = sl.get("cite")
        if marks and not cite:
            err(w, f"body uses marker(s) {sorted(marks)} but the slide has no cite")
        elif marks and isinstance(cite, str):
            err(w, "cite must be a list for markers to be auto-numbered")
        elif marks and isinstance(cite, (list, tuple)):
            bad = [n for n in marks if n < 1 or n > len(cite)]
            if bad:
                err(w, f"marker(s) {bad} have no reference (cite has {len(cite)})")
            unused = [n for n in range(1, len(cite) + 1) if n not in marks]
            if unused:
                warn(w, f"reference(s) {unused} are never marked in the body")
        if isinstance(cite, (list, tuple)):
            for c in cite:
                if "(" not in str(c) or ")" not in str(c):
                    warn(w, f"cite entry lacks a source in parentheses: {str(c)[:50]!r}")

        # notes carry the talk, so an empty one is a real gap
        if layout != "section" and len(str(sl.get("notes", "")).strip()) < MIN_NOTES_CHARS:
            err(w, "missing or very short speaker notes")

        # figures must resolve, otherwise the slide silently renders empty
        fig = sl.get("figure")
        if fig and layout not in FIGURE_LAYOUTS:
            warn(w, f"layout {layout!r} ignores the figure field")
        if fig and assets:
            if not any(os.path.exists(os.path.join(assets, fig + ext)) for ext in ("", ".png")):
                err(w, f"figure {fig!r} not found in {assets}")

        bl = bullets_of(sl)
        top = [b for b in bl if not isinstance(b, dict) or b.get("level", 0) == 0]
        if len(top) > MAX_TOP_BULLETS:
            warn(w, f"{len(top)} top-level bullets, past where the type stops scaling")
        for b in bl:
            t = (b.get("text", "") if isinstance(b, dict) else str(b)).replace("**", "")
            if len(t.split()) > MAX_BULLET_WORDS:
                warn(w, f"{len(t.split())}-word bullet: {t[:55]!r}")
        if len(str(sl.get("title", ""))) > MAX_TITLE_CHARS:
            warn(w, f"title is {len(sl['title'])} chars and may wrap past two lines")

    # closing convention
    closing = [str(s.get("title", "")).lower() for s in slides[-2:]]
    if any(s.get("layout") == "takeaways" for s in slides) and "conclusions" not in " ".join(closing):
        warn("deck", "no slide titled Conclusions near the end")
    if closing and "caveat" in closing[-1]:
        warn("deck", "closing slide reads as caveats; prefer Future work")

    for e in errors:
        print(f"error   {e}")
    for x in warnings:
        print(f"warning {x}")
    print(f"\n{len(slides)} slides, {len(errors)} error(s), {len(warnings)} warning(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
