#!/usr/bin/env python3
"""Second pass over a deck built by build_from_outline.py.

build_from_outline.py clones template slides and fills text. This adds the
things a paper talk actually needs: a red claim line under the title, a
figure with its caption, annotations drawn on that figure, a comparison
matrix, an equation band, and the bottom callout.

    python scripts/build_from_outline.py outline.json -o talk.pptx
    python scripts/compose.py outline.json talk.pptx

Regions are computed, never hand-written: say `"layout": "figure-right"`
and the body box shrinks to make room. Overlap is the failure mode every
other slide generator ships with; this is how it is avoided.
"""
from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette
RED = RGBColor(0xC0, 0x00, 0x00)      # claim line, callout frame, annotations
GREY = RGBColor(0x59, 0x59, 0x59)     # captions
HEAD_BG = RGBColor(0x3E, 0x4A, 0x5B)  # matrix header
HEAD_FG = RGBColor(0xFF, 0xFF, 0xFF)
ROW_HI = RGBColor(0xE8, 0xF2, 0xE1)   # the "Ours" row
BAND = RGBColor(0xF2, 0xF2, 0xF2)     # equation band

# ---------------------------------------------------------------- geometry
# template: 13.333 x 7.5in, title 0.40-1.85, body 2.00-6.75, page no. at 6.95
TITLE = (0.92, 0.33, 11.50, 0.92)
SUBTITLE = (0.95, 1.18, 11.50, 0.44)
CONTENT_TOP_PLAIN = 1.72
CONTENT_TOP_SUB = 1.78
CONTENT_BOTTOM = 6.80
CALLOUT_H = 0.78
CAPTION_H = 0.34
GUTTER = 0.30
LEFT, RIGHT = 0.92, 12.42          # body text column
BLEED_L, BLEED_R = 0.62, 12.72     # figures may run wider than text


def regions(spec: dict) -> dict:
    """Body box and figure box, in inches, for this slide's layout."""
    top = CONTENT_TOP_SUB if spec.get("subtitle") else CONTENT_TOP_PLAIN
    bottom = CONTENT_BOTTOM - (CALLOUT_H + 0.18 if spec.get("callout") else 0)
    h = bottom - top
    w = RIGHT - LEFT
    layout = spec.get("layout") or ("figure-bottom" if spec.get("figure") else "text-only")

    if layout == "text-only":
        return {"layout": layout, "body": (LEFT, top, w, h), "fig": None}
    if layout == "figure-right":
        bw = w * 0.50
        return {"layout": layout,
                "body": (LEFT, top, bw, h),
                "fig": (LEFT + bw + GUTTER, top, BLEED_R - (LEFT + bw + GUTTER), h)}
    if layout == "figure-left":
        fw = w * 0.48
        return {"layout": layout,
                "body": (BLEED_L + fw + GUTTER, top, RIGHT - (BLEED_L + fw + GUTTER), h),
                "fig": (BLEED_L, top, fw, h)}
    if layout == "figure-full":
        return {"layout": layout, "body": None,
                "fig": (BLEED_L, top, BLEED_R - BLEED_L, h)}
    if layout == "figure-bottom":
        th = min(1.85, h * 0.38)
        return {"layout": layout,
                "body": (LEFT, top, w, th),
                "fig": (BLEED_L, top + th + 0.12, BLEED_R - BLEED_L, h - th - 0.12)}
    raise SystemExit(f"unknown layout {layout!r}")


# ---------------------------------------------------------------- helpers

def _is_page_number(shape) -> bool:
    return shape.has_text_frame and Emu(shape.top).inches > 6.6 and Emu(shape.width).inches < 4.5


def _text_shapes(slide):
    return [s for s in slide.shapes if s.has_text_frame and not _is_page_number(s)]


def find_title(slide):
    cands = [s for s in _text_shapes(slide) if Emu(s.top).inches < 1.9]
    return max(cands, key=lambda s: s.width * s.height) if cands else None


def find_body(slide, title):
    cands = [s for s in _text_shapes(slide)
             if s is not title and Emu(s.top).inches >= 1.5]
    return max(cands, key=lambda s: s.width * s.height) if cands else None


def strip_style(shape) -> None:
    """Drop the theme's <p:style>; it re-applies a text shadow we never asked for."""
    sp = shape._element
    for st in sp.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}style'):
        sp.remove(st)


def place(shape, box) -> None:
    x, y, w, h = box
    shape.left, shape.top, shape.width, shape.height = (
        Inches(x), Inches(y), Inches(w), Inches(h))


def textbox(slide, box, text, size, *, color=None, bold=False, italic=False,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    x, y, w, h = box
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    return tb


def fit(img_box, iw, ih, *, vcenter=True):
    """Aspect-preserving fit. Side layouts centre; stacked layouts hug the top."""
    x, y, w, h = img_box
    scale = min(w / iw, h / ih)
    fw, fh = iw * scale, ih * scale
    dy = (h - fh) / 2 if vcenter else 0.0
    return (x + (w - fw) / 2, y + dy, fw, fh)


# ---------------------------------------------------------------- pieces

def add_subtitle(slide, spec) -> None:
    title = find_title(slide)
    if title is not None:
        place(title, TITLE)
        title.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    textbox(slide, SUBTITLE, spec["subtitle"], 18, color=RED)


def add_figure(slide, spec, box, *, vcenter=True) -> tuple[float, float, float, float] | None:
    from PIL import Image

    fig = spec["figure"]
    src = Path(fig["src"] if isinstance(fig, dict) else fig)
    if not src.exists():
        raise SystemExit(f"figure not found: {src}")
    fig = fig if isinstance(fig, dict) else {"src": str(src)}

    x, y, w, h = box
    if fig.get("caption"):
        h -= CAPTION_H
    iw, ih = Image.open(src).size
    rect = fit((x, y, w, h), iw, ih, vcenter=vcenter)
    slide.shapes.add_picture(str(src), Inches(rect[0]), Inches(rect[1]),
                             Inches(rect[2]), Inches(rect[3]))
    if fig.get("caption"):
        textbox(slide, (x, rect[1] + rect[3] + 0.06, w, CAPTION_H),
                fig["caption"], 11, color=GREY, italic=True, align=PP_ALIGN.CENTER)
    return rect


def add_annotations(slide, spec, rect) -> None:
    """Coordinates are fractions of the placed picture, so they survive resizing."""
    fx, fy, fw, fh = rect
    for a in spec.get("annotations", []):
        kind = a.get("type", "box")
        colour = RGBColor.from_string(a.get("color", "C00000"))
        at = a["at"]
        if kind in ("box", "circle"):
            x, y, w, h = at
            shp = slide.shapes.add_shape(
                MSO_SHAPE.OVAL if kind == "circle" else MSO_SHAPE.RECTANGLE,
                Inches(fx + x * fw), Inches(fy + y * fh),
                Inches(w * fw), Inches(h * fh))
            strip_style(shp)
            shp.fill.background()
            shp.line.color.rgb = colour
            shp.line.width = Pt(2.25)
            if a.get("dash"):
                shp.line.dash_style = 4  # MSO_LINE_DASH_STYLE.DASH
            shp.shadow.inherit = False
            if a.get("text"):
                shp.text_frame.text = a["text"]
                r = shp.text_frame.paragraphs[0].runs[0]
                r.font.size, r.font.bold, r.font.color.rgb = Pt(12), True, colour
        elif kind == "arrow":
            x0, y0, x1, y1 = at
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(fx + x0 * fw), Inches(fy + y0 * fh),
                Inches(fx + x1 * fw), Inches(fy + y1 * fh))
            conn.line.color.rgb = colour
            conn.line.width = Pt(2.25)
            _arrowhead(conn)
        elif kind == "label":
            x, y = at[0], at[1]
            textbox(slide, (fx + x * fw, fy + y * fh, 3.4, 0.36),
                    a.get("text", ""), a.get("size", 13), color=colour,
                    bold=True, wrap=False)


def _arrowhead(connector) -> None:
    from pptx.oxml.ns import qn
    ln = connector.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)


def add_callout(slide, spec) -> None:
    c = spec["callout"]
    label, text = (c.get("label"), c["text"]) if isinstance(c, dict) else (None, c)
    y = CONTENT_BOTTOM - CALLOUT_H
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(BLEED_L), Inches(y),
                                 Inches(BLEED_R - BLEED_L), Inches(CALLOUT_H))
    strip_style(shp)
    shp.fill.background()
    shp.line.color.rgb = RED
    shp.line.width = Pt(1.5)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.16)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if label:
        r = p.add_run()
        r.text = f"{label}:  "
        r.font.size, r.font.bold, r.font.color.rgb = Pt(15), True, RED
    r = p.add_run()
    r.text = text
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)


def add_matrix(slide, spec, box) -> None:
    m = spec["matrix"]
    header, rows = m["header"], m["rows"]
    hi = m.get("highlight_row")
    x, y, w, h = box
    n_rows, n_cols = len(rows) + 1, len(header)
    height = min(h, 0.42 + 0.40 * len(rows))
    gf = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y),
                                Inches(w), Inches(height))
    table = gf.table
    table.first_row = True
    for c, txt in enumerate(header):
        cell = table.cell(0, c)
        cell.text = str(txt)
        cell.fill.solid()
        cell.fill.fore_color.rgb = HEAD_BG
        _style_cell(cell, 13, bold=True, color=HEAD_FG,
                    align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    for r, row in enumerate(rows, start=1):
        for c, txt in enumerate(row[:n_cols]):
            cell = table.cell(r, c)
            cell.text = str(txt)
            if hi is not None and r == hi:
                cell.fill.solid()
                cell.fill.fore_color.rgb = ROW_HI
            else:
                cell.fill.background()
            _style_cell(cell, 12.5, bold=(hi is not None and r == hi),
                        align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)


def _style_cell(cell, size, *, bold=False, color=None, align=PP_ALIGN.LEFT):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.08)
    for p in cell.text_frame.paragraphs:
        p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold
            if color is not None:
                r.font.color.rgb = color


def add_equation(slide, spec, box) -> None:
    eq = spec["equation"]
    text, label = (eq.get("text"), eq.get("label")) if isinstance(eq, dict) else (eq, None)
    x, y, w, _ = box
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                  Inches(w), Inches(0.78))
    strip_style(band)
    band.fill.solid()
    band.fill.fore_color.rgb = BAND
    band.line.fill.background()
    band.shadow.inherit = False
    tf = band.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size, r.font.italic = Pt(20), True
    r.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    if label:
        textbox(slide, (x, y + 0.82, w, 0.3), label, 11,
                color=GREY, italic=True, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------- driver

def flatten(outline: dict) -> list[dict]:
    flat = []
    if outline.get("meta"):
        flat.append({"role": "cover"})
    if outline.get("summary"):
        flat.append({"role": "summary"})
    flat.extend(outline.get("slides", []))
    return flat


VISUAL_KEYS = ("subtitle", "figure", "annotations", "callout", "matrix", "equation")


def compose(outline_path: Path, deck: Path, out: Path | None = None) -> Path:
    outline = json.loads(outline_path.read_text(encoding="utf-8"))
    flat = flatten(outline)
    prs = Presentation(str(deck))
    if len(prs.slides._sldIdLst) != len(flat):
        raise SystemExit(
            f"{deck} has {len(prs.slides._sldIdLst)} slides but the outline "
            f"describes {len(flat)} - rebuild with build_from_outline.py first")

    touched = 0
    for slide, spec in zip(prs.slides, flat):
        if not any(spec.get(k) for k in VISUAL_KEYS):
            continue
        touched += 1
        reg = regions(spec)
        if spec.get("subtitle"):
            add_subtitle(slide, spec)

        body = find_body(slide, find_title(slide))
        if body is not None and not spec.get("bullets"):
            # no bullets means the template's own placeholder prose is still there
            body.text_frame.clear()
        if body is not None:
            if reg["body"] is None:
                body.text_frame.clear()
                place(body, (LEFT, CONTENT_TOP_PLAIN, 0.4, 0.3))
            else:
                place(body, reg["body"])
                if reg["layout"] in ("figure-right", "figure-left"):
                    # justified text in a half-width column opens rivers
                    for para in body.text_frame.paragraphs:
                        para.alignment = PP_ALIGN.LEFT

        if spec.get("equation"):
            bx, by, bw, bh = reg["body"] or (LEFT, CONTENT_TOP_PLAIN, RIGHT - LEFT, 1.0)
            add_equation(slide, spec, (bx, by + bh - 1.1, bw, 1.1))
        if spec.get("matrix"):
            add_matrix(slide, spec, reg["body"] or reg["fig"])
        if spec.get("figure") and reg["fig"]:
            rect = add_figure(slide, spec, reg["fig"],
                              vcenter=reg["layout"] in ("figure-right", "figure-left"))
            if spec.get("annotations"):
                add_annotations(slide, spec, rect)
        if spec.get("callout"):
            add_callout(slide, spec)

    out = out or deck
    prs.save(str(out))
    print(f"composed {touched}/{len(flat)} slides -> {out}")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline")
    ap.add_argument("deck")
    ap.add_argument("-o", "--output", help="default: edit the deck in place")
    a = ap.parse_args()
    compose(Path(a.outline), Path(a.deck), Path(a.output) if a.output else None)


if __name__ == "__main__":
    main()
