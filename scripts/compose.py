#!/usr/bin/env python3
"""Second pass over a deck built by build_from_outline.py.

build_from_outline.py clones template slides and fills text. This adds the
things a paper talk actually needs: a red claim line under the title, a
figure with its caption, annotations drawn on that figure, a comparison
matrix, a native editable diagram, an equation band, an embedded video, the
bottom callout -- and the graphic exhibits a paper talk is otherwise missing:
cards, a flow, big numbers, a positioning map, or your own draw function
(scripts/exhibits.py).

    python scripts/build_from_outline.py outline.json -o talk.pptx
    python scripts/compose.py outline.json talk.pptx

Regions are computed, never hand-written: say `"layout": "figure-right"`
and the body box shrinks to make room. Overlap is the failure mode every
other slide generator ships with; this is how it is avoided.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

import cjk
import exhibits

# ---------------------------------------------------------------- palette
RED = RGBColor(0xC0, 0x00, 0x00)      # claim line, callout frame, annotations
GREY = RGBColor(0x59, 0x59, 0x59)     # captions
HEAD_FG = RGBColor(0xFF, 0xFF, 0xFF)
# matrix head, banding and the decisive column come from the section colour
# (exhibits.tones); the "Ours" row is pale red, the lab's one focus colour
ROW_HI = RGBColor.from_string(exhibits.RED_PALE)
BAND = RGBColor(0xF2, 0xF2, 0xF2)     # equation band
DIAGRAM_FILL = RGBColor(0xFA, 0xFA, 0xFA)
DIAGRAM_LINE = RGBColor(0x5B, 0x67, 0x78)
DIAGRAM_TEXT = RGBColor(0x1A, 0x1A, 0x1A)
DIAGRAM_GROUP = RGBColor(0xA6, 0xAD, 0xB8)

# ---------------------------------------------------------------- geometry
# template: 13.333 x 7.5in, title 0.40-1.85, body 2.00-6.75, page no. at 6.95
# The title band has to stop above the subtitle: the title is bottom-anchored,
# so its box bottom is where the glyphs sit. 0.28 + 0.86 = 1.14, clear of 1.16.
TITLE = (0.92, 0.28, 11.50, 0.86)
SUBTITLE = (0.95, 1.16, 11.50, 0.44)
CONTENT_TOP_PLAIN = 1.72
CONTENT_TOP_SUB = 1.78
CONTENT_BOTTOM = 6.80
CALLOUT_H = 0.78
CAPTION_H = 0.34
GUTTER = 0.30
LEFT, RIGHT = 0.92, 12.42          # body text column
BLEED_L, BLEED_R = 0.62, 12.72     # figures may run wider than text
SETTLE = 0.40     # share of the left-over height that goes above a short group
STACK_GAP = 0.28  # between the bullets and the exhibit under them
ASSERT = (0.92, 1.16, 11.50, 0.88)  # assertion-evidence: the sentence headline
ASSERT_TOP = 2.10                   # evidence starts below it
DIVIDER = RGBColor(0xC8, 0xCC, 0xD2)


WIDE_GRAPHICS = ("cards", "flow", "bignum", "matrix", "draw")
GRAPHICS = exhibits.GRAPHIC_KEYS + ("matrix", "diagram")   # one of these per slide


def default_layout(spec: dict) -> str:
    """Where the exhibit goes when the outline does not say.

    A wide graphic (cards, flow, numbers, a table) with bullets sits under
    them; alone it takes the whole content region. A quadrant is square, so
    it shares the slide side by side. A figure defaults to the bottom, which
    is the one place a paper's wide architecture diagram stays legible.
    """
    has_bullets = bool(spec.get("bullets"))
    if spec.get("video") or spec.get("diagram"):
        return "figure-right" if has_bullets else "figure-full"
    if spec.get("quadrant"):
        return "figure-right" if has_bullets else "figure-full"
    if any(spec.get(k) for k in WIDE_GRAPHICS):
        return "figure-bottom" if has_bullets else "figure-full"
    if spec.get("figure"):
        return "figure-bottom" if has_bullets else "figure-full"
    if spec.get("stage"):
        return "figure-right"   # the pipeline strip lives where a figure would
    return "text-only"


BODY_PT = 20                 # the template's body size
LINE_H = BODY_PT * 1.25 / 72  # inches per line


def body_lines(spec: dict, width: float, pt: float = BODY_PT) -> int:
    """How many lines the bullets take at `pt` (the template's 20pt) in `width` inches.

    Measured with the deck's own face (measure.py); the old flat 0.55em guess
    called a 44-character bullet two lines when it is one, and every such
    miss left a hole under the text block it was used to place. The body box
    loses 0.2in to its insets and 0.44in to the hanging bullet.
    """
    import measure
    n = 0
    for b in spec.get("bullets") or []:
        text = b["text"] if isinstance(b, dict) else str(b)
        level = int(b.get("level", 0)) if isinstance(b, dict) else 0
        bold = bool(b.get("bold")) if isinstance(b, dict) else False
        n += max(1, measure.lines(text, max(0.5, width - 0.64 - 0.40 * level), pt, bold))
    return n


def body_strip_height(spec: dict, width: float) -> float:
    """The strip a `figure-bottom` slide gives its bullets: as tall as the
    lines need and no taller (feedback B1/E3), so a one-line lead-in leaves
    the whole width *and* most of the height to the figure."""
    n = body_lines(spec, width)
    return 0.0 if n == 0 else 0.22 + LINE_H * n


def regions(spec: dict) -> dict:
    """Body box and figure box, in inches, for this slide's layout."""
    top = CONTENT_TOP_SUB if spec.get("subtitle") else CONTENT_TOP_PLAIN
    bottom = CONTENT_BOTTOM - (CALLOUT_H + 0.18 if spec.get("callout") else 0)
    h = bottom - top
    w = RIGHT - LEFT
    layout = spec.get("layout")
    if not layout:
        layout = default_layout(spec)
    elif layout == "text-only" and spec.get("bullets") and any(
            spec.get(k) for k in WIDE_GRAPHICS + ("quadrant", "diagram")):
        # bullets and a graphic cannot share one column; stack them
        layout = "figure-bottom"

    span = (top, bottom)
    if layout in ("assertion-evidence", "ae"):
        # Alley's structure: a sentence headline, then visual evidence, no bullets
        return {"layout": "assertion-evidence", "span": (ASSERT_TOP, bottom),
                "body": None,
                "fig": (BLEED_L, ASSERT_TOP, BLEED_R - BLEED_L, bottom - ASSERT_TOP)}
    if layout == "text-only":
        return {"layout": layout, "span": span, "body": (LEFT, top, w, h), "fig": None}
    if layout == "figure-right":
        bw = w * 0.50
        return {"layout": layout, "span": span,
                "body": (LEFT, top, bw, h),
                "fig": (LEFT + bw + GUTTER, top, BLEED_R - (LEFT + bw + GUTTER), h)}
    if layout == "figure-left":
        fw = w * 0.48
        return {"layout": layout, "span": span,
                "body": (BLEED_L + fw + GUTTER, top, RIGHT - (BLEED_L + fw + GUTTER), h),
                "fig": (BLEED_L, top, fw, h)}
    if layout == "figure-full":
        return {"layout": layout, "span": span, "body": None,
                "fig": (BLEED_L, top, BLEED_R - BLEED_L, h)}
    if layout == "figure-bottom":
        th = min(1.85, h * 0.38, body_strip_height(spec, w))
        return {"layout": layout, "span": span,
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


MIN_BODY_AREA = 2.0     # in^2; below this it is a label, not the body placeholder


def find_body(slide, title):
    """The bullet region -- never a small fixed label such as the Todolist header.

    Without the area floor, a conclusion slide (whose only other text frame is
    the template's `Todolist & Suggestion from Prof.` caption) hands that label
    back as the body, and the caller then clears it and stretches it across the
    content area.
    """
    cands = [s for s in _text_shapes(slide)
             if s is not title and Emu(s.top).inches >= 1.5
             and Emu(s.width).inches * Emu(s.height).inches >= MIN_BODY_AREA]
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
    run.font.name = exhibits.FONT
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    return tb


def _set_shape_name(shape, name: str) -> None:
    """Give generated objects stable names in PowerPoint's Selection Pane."""
    nodes = shape._element.xpath(".//p:cNvPr")
    if nodes:
        nodes[0].set("name", name)


def caption(slide, box, text, size=11, *, align=PP_ALIGN.CENTER, color=None):
    """Grey italic for Latin captions; 中文 has no true italic, so it is set
    upright and never below 12pt (feedback A4)."""
    if cjk.has_cjk(text):
        return textbox(slide, box, text, max(12, size), color=color or GREY,
                       italic=False, align=align)
    return textbox(slide, box, text, size, color=color or GREY, italic=True, align=align)


def fit(img_box, iw, ih, *, vcenter=True):
    """Aspect-preserving fit. Side layouts centre; stacked layouts hug the top."""
    x, y, w, h = img_box
    scale = min(w / iw, h / ih)
    fw, fh = iw * scale, ih * scale
    dy = (h - fh) / 2 if vcenter else 0.0
    return (x + (w - fw) / 2, y + dy, fw, fh)


# ---------------------------------------------------------------- pieces

SUBTITLE_NAME = "acm:subtitle"


def subtitle_box(slide, text: str):
    """The red claim line. Named, so a later pass finds it instead of adding
    another (feedback B4/F7: build_from_outline draws it too)."""
    for shp in slide.shapes:
        if shp.name == SUBTITLE_NAME:
            shp.text_frame.paragraphs[0].runs[0].text = text
            place(shp, SUBTITLE)
            return shp
    tb = textbox(slide, SUBTITLE, text, 18, color=RED)
    tb.name = SUBTITLE_NAME
    return tb


def add_subtitle(slide, spec) -> None:
    title = find_title(slide)
    if title is not None:
        place(title, TITLE)
        title.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    subtitle_box(slide, spec["subtitle"])


def add_assertion(slide, spec) -> None:
    """The claim, at reading size, in place of the small red subtitle."""
    title = find_title(slide)
    if title is not None:
        place(title, TITLE)
        title.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    for shp in list(slide.shapes):               # build's small red line, if any
        if shp.name == SUBTITLE_NAME:
            shp._element.getparent().remove(shp._element)
    textbox(slide, ASSERT, spec["subtitle"], 24,
            color=RGBColor(0x1A, 0x1A, 0x1A), bold=True, anchor=MSO_ANCHOR.TOP)


def add_figure(slide, spec, box, *, vcenter=True) -> tuple[float, float, float, float] | None:
    from PIL import Image

    fig = spec["figure"]
    fig = fig if isinstance(fig, dict) else {"src": fig}
    srcs = fig["src"] if isinstance(fig["src"], list) else [fig["src"]]
    if not 1 <= len(srcs) <= 2:
        raise SystemExit("figure.src takes one path, or two to stack (feedback E2)")
    for s_ in srcs:
        if not Path(s_).exists():
            raise SystemExit(f"figure not found: {s_}")

    x, y, w, h = box
    if fig.get("caption"):
        h -= CAPTION_H
    if len(srcs) == 2:
        # two parts of one exhibit -- the paper's equation and the figure it
        # describes -- stacked ("column", default) or side by side ("row"),
        # each fitted to its half
        gap = 0.14
        sizes = [Image.open(s_).size for s_ in srcs]
        if fig.get("stack") == "row":
            halves = [(x, y, (w - gap) / 2, h), (x + (w + gap) / 2, y, (w - gap) / 2, h)]
        else:
            need = [ih_ / iw_ * w for iw_, ih_ in sizes]   # height each wants at full width
            frac = need[0] / sum(need)
            h0 = (h - gap) * frac
            halves = [(x, y, w, h0), (x, y + h0 + gap, w, h - gap - h0)]
        rects = [fit(hb, iw_, ih_, vcenter=True) for hb, (iw_, ih_) in zip(halves, sizes)]
        for s_, r in zip(srcs, rects):
            slide.shapes.add_picture(s_, Inches(r[0]), Inches(r[1]), Inches(r[2]), Inches(r[3]))
        left, top = min(r[0] for r in rects), min(r[1] for r in rects)
        right = max(r[0] + r[2] for r in rects)
        bottom = max(r[1] + r[3] for r in rects)
        rect = (left, top, right - left, bottom - top)
        if fig.get("caption"):
            caption(slide, (x, rect[1] + rect[3] + 0.06, w, CAPTION_H), fig["caption"])
        return rect
    src = srcs[0]
    iw, ih = Image.open(src).size
    rect = fit((x, y, w, h), iw, ih, vcenter=vcenter)
    slide.shapes.add_picture(str(src), Inches(rect[0]), Inches(rect[1]),
                             Inches(rect[2]), Inches(rect[3]))
    if fig.get("caption"):
        caption(slide, (x, rect[1] + rect[3] + 0.06, w, CAPTION_H), fig["caption"])
    return rect


BADGE_BG = RGBColor(0x1A, 0x1A, 0x1A)   # the play badge on a video poster
BADGE_H = 0.30


def _play_badge(slide, rect, label: str) -> None:
    """A small marker so nobody mistakes the poster for a still figure."""
    x, y, w, h = rect
    bw = 0.36 + 0.11 * len(label)
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x + 0.10), Inches(y + h - BADGE_H - 0.10),
                                 Inches(bw), Inches(BADGE_H))
    strip_style(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = BADGE_BG
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size, r.font.bold = Pt(11), True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _set_playback(movie, opts: dict) -> None:
    """Rewrite the `p:timing` node python-pptx wrote, for autoplay and loop.

    `add_movie` always emits `<p:cond delay="indefinite"/>` - play on click.
    Flipping that one attribute to `delay="0"` is what PowerPoint itself writes
    for "Start: Automatically", and `repeatCount="indefinite"` on the same time
    node is "Loop until Stopped". Both are plain ISO-29500 attributes, so the
    result still passes `office/validate.py`.
    """
    from pptx.oxml.ns import qn

    pic = movie._element
    spid = pic.find(qn("p:nvPicPr")).find(qn("p:cNvPr")).get("id")
    sld = pic.getroottree().getroot()
    tgt = sld.xpath('.//p:timing//p:video//p:spTgt[@spid="%s"]' % spid)
    if not tgt:
        raise SystemExit(
            "python-pptx did not write a p:timing node for this movie. That "
            "happens when the slide already has one wrapped in "
            "<mc:AlternateContent> (a non-OOXML animation, python-pptx issue "
            "#954) - remove the animation from that slide and rebuild.")
    media = tgt[0].getparent().getparent()          # spTgt -> tgtEl -> cMediaNode
    cTn = media.find(qn("p:cTn"))
    if opts.get("autoplay"):
        cTn.find(qn("p:stCondLst")).find(qn("p:cond")).set("delay", "0")
    if opts.get("loop"):
        cTn.set("repeatCount", "indefinite")
    if opts.get("mute"):
        media.set("mute", "1")
    if opts.get("volume") is not None:
        pct = max(0.0, min(100.0, float(opts["volume"])))
        media.set("vol", str(int(round(pct * 1000))))   # 1000ths of a percent
    if opts.get("fullscreen"):
        media.getparent().set("fullScrn", "1")


def add_video(slide, spec, box, *, vcenter=True) -> tuple[float, float, float, float]:
    """Embed a movie the way PowerPoint expects to find one.

    Three things this gets right that a bare `add_movie` call does not: the
    media part is typed by extension instead of `video/unknown`, the file is
    refused up front when PowerPoint cannot decode it, and a real frame is
    pulled from the video as the poster so the slide is not a grey speaker icon.
    """
    import video as video_mod
    from PIL import Image

    v = spec["video"]
    v = {"src": v} if isinstance(v, str) else dict(v)
    src = Path(v["src"])
    if not src.exists():
        raise SystemExit(f"video not found: {src}")

    info = video_mod.probe(src)
    blocking = [m for lv, m in video_mod.problems(src, info) if lv == "error"]
    if blocking:
        raise SystemExit(
            f"{src} will not play in PowerPoint:\n  - "
            + "\n  - ".join(blocking)
            + f"\nFix it once with:\n  {video_mod.fix_command(src)}")

    still = Path(v["poster"]) if v.get("poster") else \
        video_mod.auto_poster(src, v.get("poster_at"))
    if not still.exists():
        raise SystemExit(f"poster image not found: {still}")

    x, y, w, h = box
    if v.get("caption"):
        h -= CAPTION_H
    iw, ih = Image.open(still).size
    rect = fit((x, y, w, h), iw, ih, vcenter=vcenter)
    movie = slide.shapes.add_movie(
        str(src), Inches(rect[0]), Inches(rect[1]), Inches(rect[2]), Inches(rect[3]),
        poster_frame_image=str(still), mime_type=video_mod.mime_for(src))
    _set_playback(movie, v)

    if v.get("badge", True):
        secs = int(info.get("duration") or 0)
        _play_badge(slide, rect,
                    f"\u25b6 {secs // 60}:{secs % 60:02d}" if secs else "\u25b6")
    if v.get("caption"):
        caption(slide, (x, rect[1] + rect[3] + 0.06, w, CAPTION_H), v["caption"])
    return rect


def add_annotations(slide, spec, rect) -> None:
    """Coordinates are fractions of the placed picture, so they survive resizing."""
    fx, fy, fw, fh = rect
    for a in spec.get("annotations", []):
        kind = a.get("type", "box")
        colour = RGBColor.from_string(a.get("color", "C00000"))
        at = a["at"] if "at" in a else xyxy_to_xywh(a["xyxy"])
        if kind in ("box", "circle"):
            x, y, w, h = at
            if x + w > 1.02 or y + h > 1.02:
                raise SystemExit(
                    f"annotation {kind} at {at} runs past the picture edge. `at` is "
                    f"[x, y, w, h] as fractions of the picture; if you have corners "
                    f"(x0, y0, x1, y1), the way figure.py crop takes them, write "
                    f"`\"xyxy\": [...]` instead.")
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


def xyxy_to_xywh(b) -> list[float]:
    """Corners (x0, y0, x1, y1) -> (x, y, w, h). Both are fractions of the picture."""
    x0, y0, x1, y1 = b
    if x1 < x0 or y1 < y0:
        raise SystemExit(f"xyxy box {b} has its corners the wrong way round")
    return [x0, y0, x1 - x0, y1 - y0]


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


def add_matrix(slide, spec, box, accent: str = exhibits.SECTION["others"]) -> None:
    """A comparison table sized to the room it has, not to a fixed 12.5pt.

        "matrix": {
          "header": ["Method", "Duration", "PSNR", "CSIM"],
          "rows": [...],
          "highlight_row": 4,           // 1-based; also accepts a list
          "highlight_col": 2,           // 1-based; the decisive metric
          "col_widths": [2.2, 1, 1, 1], // relative; first column wider by default
          "caption": "Self re-enactment on INSTA; higher is better except LPIPS",
          "size": 16                    // override the computed font size
        }

    Type scales with the number of rows and the box height, rows grow to
    fill up to ~70% of the region, and a short table is centred in it, so
    a four-row comparison no longer sits at 12.5pt above half a page of
    white space.
    """
    m = spec["matrix"]
    header, rows = m["header"], m["rows"]
    hi_rows = m.get("highlight_row")
    hi_rows = set(hi_rows if isinstance(hi_rows, list) else ([hi_rows] if hi_rows else []))
    hi_col = m.get("highlight_col")
    t = exhibits.tones(accent)
    head_bg = RGBColor.from_string(t["dark"])
    band = RGBColor.from_string(t["wash"])
    col_hi = RGBColor.from_string(exhibits.tint(accent, 0.80))
    x, y, w, h = box
    n_rows, n_cols = len(rows) + 1, len(header)

    # row height: enough for the type, capped so the table does not balloon
    size = m.get("size")
    if size is None:
        size = 18 if n_rows <= 4 else 16 if n_rows <= 6 else 14 if n_rows <= 9 else 12
    row_h = max(0.36, min(0.70, size * 2.5 / 72))
    cap_h = 0.32 if m.get("caption") else 0
    height = row_h * n_rows
    avail = h - cap_h
    if height > avail:                        # too many rows: shrink to fit
        row_h = avail / n_rows
        size = max(10, min(size, row_h * 72 / 2.3))
        height = avail
    top = y + max(0.0, (avail - height) * 0.35) if height < avail * 0.6 else y

    widths = m.get("col_widths")
    if not widths:
        # equal columns are a generated-deck tell and waste width on short
        # numeric metrics: weight each column by its longest cell, label
        # column slightly favoured, no column allowed to dominate
        widths = []
        for c in range(n_cols):
            values = [str(header[c])] + [str(row[c]) for row in rows if c < len(row)]
            widths.append(max(4, min(28, max(len(v) for v in values))))
        widths[0] *= 1.22
    if len(widths) != n_cols:
        raise SystemExit(f"matrix: col_widths has {len(widths)} entries for {n_cols} columns")
    unit = w / sum(widths)

    gf = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(top),
                                Inches(w), Inches(height))
    table = gf.table
    table.first_row = True
    for c, frac in enumerate(widths):
        table.columns[c].width = Inches(frac * unit)
    for r in range(n_rows):
        table.rows[r].height = Inches(row_h)

    for c, txt in enumerate(header):
        cell = table.cell(0, c)
        cell.text = str(txt)
        cell.fill.solid()
        cell.fill.fore_color.rgb = head_bg
        _style_cell(cell, size, bold=True, color=HEAD_FG,
                    align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    for r, row in enumerate(rows, start=1):
        table.rows[r].height = Inches(max(0.34, (height - 0.42) / max(1, len(rows))))
        for c, txt in enumerate(row[:n_cols]):
            cell = table.cell(r, c)
            cell.text = str(txt)
            lit_row = r in hi_rows
            lit_col = hi_col is not None and c + 1 == hi_col
            if lit_row and lit_col:
                cell.fill.solid(); cell.fill.fore_color.rgb = ROW_HI
                _style_cell(cell, size, bold=True, color=RED, align=PP_ALIGN.CENTER)
            elif lit_row:
                cell.fill.solid(); cell.fill.fore_color.rgb = ROW_HI
                _style_cell(cell, size, bold=True, color=RED if c == 0 else None,
                            align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
            elif lit_col:
                cell.fill.solid(); cell.fill.fore_color.rgb = col_hi
                _style_cell(cell, size, bold=False,
                            align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
            else:
                # banded rows in the section's wash keep a wide row readable
                # across the slide without a grid of rules
                if r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = band
                else:
                    cell.fill.background()
                _style_cell(cell, size, bold=(c == 0),
                            align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    if m.get("caption"):
        caption(slide, (x, top + height + 0.06, w, 0.28), m["caption"])


# ---------------------------------------------------------------- native diagrams

DIAGRAM_SHAPES = {
    # Shape carries semantics. Plain processes are square rectangles; rounded
    # boxes are reserved for explicit use, and terminators stay terminators.
    # This avoids the generic "every idea is a pastel pill" AI aesthetic.
    "process": MSO_SHAPE.RECTANGLE,
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded": MSO_SHAPE.ROUNDED_RECTANGLE,
    "decision": MSO_SHAPE.DIAMOND,
    "terminator": MSO_SHAPE.FLOWCHART_TERMINATOR,
    "circle": MSO_SHAPE.OVAL,
    "database": MSO_SHAPE.CAN,
}


def _diagram_levels(node_ids: list[str], edges: list[dict]) -> dict[str, int]:
    """Longest-path layers for a DAG, with deterministic cycle fallback.

    Most slide diagrams are small DAGs.  Kahn layering gives branches the
    expected shared column/row.  A feedback loop leaves nodes behind; those
    are placed in declaration order after their deepest resolved predecessor,
    and the back edge is still drawn.  This keeps a loop editable without
    pretending to be a general graph-layout engine.
    """
    order = {node_id: i for i, node_id in enumerate(node_ids)}
    incoming = {node_id: 0 for node_id in node_ids}
    outgoing = {node_id: [] for node_id in node_ids}
    predecessors = {node_id: [] for node_id in node_ids}
    for edge in edges:
        source, target = edge["from"], edge["to"]
        if source == target:
            continue
        outgoing[source].append(target)
        predecessors[target].append(source)
        incoming[target] += 1

    queue = sorted((n for n in node_ids if incoming[n] == 0), key=order.get)
    levels = {n: 0 for n in queue}
    done = set()
    while queue:
        node_id = queue.pop(0)
        done.add(node_id)
        for target in outgoing[node_id]:
            levels[target] = max(levels.get(target, 0), levels[node_id] + 1)
            incoming[target] -= 1
            if incoming[target] == 0:
                queue.append(target)
                queue.sort(key=order.get)

    next_level = max(levels.values(), default=-1) + 1
    for node_id in node_ids:
        if node_id in done:
            continue
        resolved = [levels[p] + 1 for p in predecessors[node_id] if p in levels]
        levels[node_id] = max(resolved, default=next_level)
        next_level = max(next_level, levels[node_id] + 1)
    return levels


def _diagram_boxes(diagram: dict, box) -> dict[str, tuple[float, float, float, float]]:
    nodes = diagram["nodes"]
    node_ids = [str(n["id"]) for n in nodes]
    direction = str(diagram.get("direction", "LR")).upper()
    if direction not in {"LR", "RL", "TB", "BT"}:
        raise SystemExit("diagram.direction must be LR, RL, TB or BT")
    levels = _diagram_levels(node_ids, diagram.get("edges") or [])
    columns: dict[int, list[str]] = {}
    for node_id in node_ids:
        columns.setdefault(levels[node_id], []).append(node_id)
    level_values = sorted(columns)
    level_index = {value: i for i, value in enumerate(level_values)}
    x, y, w, h = box
    horizontal = direction in {"LR", "RL"}
    primary_count = max(1, len(level_values))
    secondary_count = max((len(v) for v in columns.values()), default=1)

    if horizontal:
        cell_w, cell_h = w / primary_count, h / secondary_count
        node_w = max(0.72, min(2.30, cell_w * 0.74))
        node_h = max(0.42, min(0.78, cell_h * 0.56))
    else:
        cell_w, cell_h = w / secondary_count, h / primary_count
        node_w = max(0.82, min(2.40, cell_w * 0.70))
        node_h = max(0.42, min(0.78, cell_h * 0.56))
    # one height for every node, tall enough for the wordiest label: a fixed
    # 0.78in box let a three-line label spill over its frame
    need = max(_node_text_height(n, node_w) for n in nodes)
    node_h = max(node_h, min(need, cell_h * 0.92))

    out = {}
    for node in nodes:
        node_id = str(node["id"])
        if node.get("at") is not None:
            at = node["at"]
            if not isinstance(at, list) or len(at) != 4:
                raise SystemExit(f"diagram node {node_id!r}: `at` must be [x,y,w,h]")
            nx, ny, nw, nh = (float(v) for v in at)
            if min(nx, ny, nw, nh) < 0 or nx + nw > 1 or ny + nh > 1:
                raise SystemExit(f"diagram node {node_id!r}: `at` must stay inside 0..1")
            out[node_id] = (x + nx * w, y + ny * h, nw * w, nh * h)
            continue

        layer = level_index[levels[node_id]]
        peers = columns[levels[node_id]]
        slot = peers.index(node_id)
        if horizontal:
            px = layer if direction == "LR" else primary_count - 1 - layer
            cx = x + (px + 0.5) * cell_w
            cy = y + (slot + 0.5) * (h / len(peers))
        else:
            py = layer if direction == "TB" else primary_count - 1 - layer
            cx = x + (slot + 0.5) * (w / len(peers))
            cy = y + (py + 0.5) * cell_h
        nw, nh = node_w, node_h
        if node.get("kind") == "decision":
            nw, nh = min(cell_w * 0.78, nw * 1.12), min(cell_h * 0.78, nh * 1.28)
        if node.get("kind") == "circle":
            side = min(nw, nh)
            nw = nh = side
        out[node_id] = (cx - nw / 2, cy - nh / 2, nw, nh)
    return out


def _node_lines(node) -> list[tuple[str, float, bool]]:
    """`text` split on newlines: the first line is the node's name, the rest
    a smaller second line -- what it does, what it guarantees."""
    size = float(node.get("size", 16))
    parts = str(node.get("text") or node.get("id")).split("\n")
    bold = bool(node.get("bold", node.get("accent") or len(parts) > 1))
    return [(parts[0], size, bold)] + [(t, max(10.0, size - 3), False) for t in parts[1:]]


def _node_text_height(node, width: float) -> float:
    import measure
    inner = width * (0.78 if node.get("kind") in ("terminator", "decision", "circle") else 0.94) - 0.12
    return 0.22 + sum(max(1, measure.lines(t, inner, pt, b)) * pt * 1.18 / 72
                      for t, pt, b in _node_lines(node))


def _edge_points(source, target, direction: str):
    sx, sy, sw, sh = source
    tx, ty, tw, th = target
    if direction == "LR":
        return (sx + sw, sy + sh / 2, tx, ty + th / 2)
    if direction == "RL":
        return (sx, sy + sh / 2, tx + tw, ty + th / 2)
    if direction == "TB":
        return (sx + sw / 2, sy + sh, tx + tw / 2, ty)
    return (sx + sw / 2, sy, tx + tw / 2, ty + th)


def _rgb(value, default: RGBColor) -> RGBColor:
    if value is None:
        return default
    return RGBColor.from_string(str(value).lstrip("#").upper())


def _validate_diagram(diagram: dict) -> None:
    if not isinstance(diagram, dict):
        raise SystemExit("diagram must be an object with nodes and edges")
    nodes = diagram.get("nodes") or []
    edges = diagram.get("edges") or []
    if len(nodes) < 2:
        raise SystemExit("diagram needs at least two nodes")
    if len(nodes) > 20 or len(edges) > 32:
        raise SystemExit("diagram is too dense for one slide (max 20 nodes / 32 edges)")
    ids = [str(n.get("id", "")) for n in nodes]
    if any(not node_id for node_id in ids) or len(ids) != len(set(ids)):
        raise SystemExit("diagram node ids must be present and unique")
    known = set(ids)
    for edge in edges:
        if edge.get("from") not in known or edge.get("to") not in known:
            raise SystemExit(f"diagram edge references an unknown node: {edge}")
    for group in diagram.get("groups") or []:
        missing = set(group.get("nodes") or []) - known
        if missing:
            raise SystemExit(f"diagram group {group.get('id')!r} has unknown nodes: {sorted(missing)}")


def add_diagram(slide, spec, box) -> None:
    """Draw an editable flow/architecture diagram with native PPT objects."""
    diagram = spec["diagram"]
    _validate_diagram(diagram)
    x, y, w, h = box
    cap_text = diagram.get("caption")
    if cap_text:
        h -= CAPTION_H
    canvas = (x, y, w, h)
    boxes = _diagram_boxes(diagram, canvas)
    direction = str(diagram.get("direction", "LR")).upper()

    # Frames and connectors go behind nodes. They remain separate native
    # objects so PowerPoint users can recolour, reroute, relabel or delete them.
    for index, group in enumerate(diagram.get("groups") or [], start=1):
        members = [boxes[node_id] for node_id in group.get("nodes") or []]
        if not members:
            continue
        pad, title_h = 0.12, 0.25
        gx = min(b[0] for b in members) - pad
        gy = min(b[1] for b in members) - pad - title_h
        gr = max(b[0] + b[2] for b in members) + pad
        gb = max(b[1] + b[3] for b in members) + pad
        frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Inches(gx), Inches(gy), Inches(gr - gx), Inches(gb - gy))
        strip_style(frame)
        frame.fill.background()
        frame.line.color.rgb = _rgb(group.get("color"), DIAGRAM_GROUP)
        frame.line.width = Pt(1.0)
        frame.line.dash_style = 4
        frame.shadow.inherit = False
        _set_shape_name(frame, f"diagram:group:{group.get('id', index)}")
        tf = frame.text_frame
        tf.clear()
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = Inches(0.10)
        tf.margin_top = Inches(0.02)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = str(group.get("label") or group.get("id") or "")
        r.font.size, r.font.bold, r.font.color.rgb = Pt(11), True, DIAGRAM_GROUP

    connector_records = []
    for index, edge in enumerate(diagram.get("edges") or [], start=1):
        points = _edge_points(boxes[edge["from"]], boxes[edge["to"]], direction)
        x0, y0, x1, y1 = points
        offset = abs(y1 - y0) if direction in {"LR", "RL"} else abs(x1 - x0)
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.ELBOW if offset > 0.08 else MSO_CONNECTOR.STRAIGHT,
            Inches(x0), Inches(y0), Inches(x1), Inches(y1))
        connector.line.color.rgb = _rgb(edge.get("color"), DIAGRAM_LINE)
        connector.line.width = Pt(float(edge.get("width", 1.6)))
        if edge.get("dashed"):
            connector.line.dash_style = 4
        if edge.get("arrow", True):
            _arrowhead(connector)
        edge_id = edge.get("id") or f"e{index}"
        _set_shape_name(connector, f"diagram:edge:{edge_id}")
        connector_records.append((connector, edge))
        if edge.get("label"):
            mx, my = (x0 + x1) / 2, (y0 + y1) / 2
            label = textbox(slide, (mx - 0.48, my - 0.16, 0.96, 0.30),
                            str(edge["label"]), 10.5, color=DIAGRAM_TEXT,
                            bold=True, align=PP_ALIGN.CENTER,
                            anchor=MSO_ANCHOR.MIDDLE)
            label.fill.solid()
            label.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            label.line.fill.background()
            _set_shape_name(label, f"diagram:label:{edge_id}")

    node_shapes = {}
    for node in diagram["nodes"]:
        node_id = str(node["id"])
        nx, ny, nw, nh = boxes[node_id]
        kind = str(node.get("kind", "process"))
        if kind not in DIAGRAM_SHAPES:
            raise SystemExit(f"diagram node {node_id!r}: unknown kind {kind!r}")
        shape = slide.shapes.add_shape(DIAGRAM_SHAPES[kind],
                                       Inches(nx), Inches(ny), Inches(nw), Inches(nh))
        strip_style(shape)
        accent = bool(node.get("accent"))
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(node.get("fill"),
                                         RGBColor(0xFD, 0xEE, 0xEE) if accent else DIAGRAM_FILL)
        shape.line.color.rgb = _rgb(node.get("line"), RED if accent else DIAGRAM_LINE)
        shape.line.width = Pt(2.0 if accent else 1.25)
        shape.shadow.inherit = False
        _set_shape_name(shape, f"diagram:node:{node_id}")
        node_shapes[node_id] = shape
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Inches(0.06)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        for li, (text, pt, bold) in enumerate(_node_lines(node)):
            p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = text
            r.font.size = Pt(pt)
            r.font.bold = bold
            r.font.color.rgb = _rgb(node.get("text_color"),
                                    DIAGRAM_TEXT if li == 0 else GREY)

    # Attach endpoints after every node exists. The connector was inserted
    # first, so it stays visually behind the nodes while still following them
    # when a user moves a box in PowerPoint.
    points = {
        "LR": (3, 1),   # right -> left
        "RL": (1, 3),   # left -> right
        "TB": (2, 0),   # bottom -> top
        "BT": (0, 2),   # top -> bottom
    }
    begin_idx, end_idx = points[direction]
    for connector, edge in connector_records:
        if edge.get("attached", True):
            connector.begin_connect(node_shapes[edge["from"]], begin_idx)
            connector.end_connect(node_shapes[edge["to"]], end_idx)

    if cap_text:
        # under the drawing, not under the empty canvas it was laid out in
        low = max(b[1] + b[3] for b in boxes.values())
        caption(slide, (x, min(low + 0.30, y + h + 0.05), w, CAPTION_H), str(cap_text))


def _style_cell(cell, size, *, bold=False, color=None, align=PP_ALIGN.LEFT):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.08)
    for p in cell.text_frame.paragraphs:
        p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.name = exhibits.FONT
            r.font.bold = bold
            if color is not None:
                r.font.color.rgb = color


def equation_height(eq: dict) -> float:
    """Vertical band an equation needs, including its where-list and captions."""
    h = 1.15
    h += 0.24 * len(eq.get("where") or [])
    h += 0.28 * len(eq.get("captions") or {})
    if eq.get("label"):
        h += 0.26
    return min(h, 3.4)


def _equation_image(eq: dict, outline_dir: Path, tag: str) -> Path:
    """A cropped equation is used as given; a rebuilt one is rendered now."""
    if eq.get("src"):
        src = Path(eq["src"])
        if not src.exists():
            raise SystemExit(f"equation image not found: {src}")
        return src
    parts = eq.get("parts")
    if not parts:
        raise SystemExit("equation needs either `src` (quote it) or `parts` (explain it)")
    import equation as equation_mod
    pairs = [(p[0], (p[1] if len(p) > 1 else "111111").lstrip("#").upper())
             for p in parts]
    # scratch, like the render PNGs: the deck is the only thing that lands in
    # the user's directory
    import tempfile
    out = Path(tempfile.gettempdir()) / "acm-eq" / f"{outline_dir.name}-{tag}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    return equation_mod.render(pairs, out)


def add_equation(slide, spec, box, outline_dir: Path, tag: str) -> None:
    eq = spec["equation"]
    if isinstance(eq, str):                      # tolerate the old plain-text form
        eq = {"parts": [[eq]]}
    from PIL import Image

    x, y, w, h = box
    where = eq.get("where") or []
    captions = eq.get("captions") or {}
    below = 0.24 * len(where) + 0.28 * len(captions) + (0.26 if eq.get("label") else 0)

    img = _equation_image(eq, outline_dir, tag)
    iw, ih = Image.open(img).size
    rect = fit((x, y, w, max(0.5, h - below)), iw, ih, vcenter=True)
    slide.shapes.add_picture(str(img), Inches(rect[0]), Inches(rect[1]),
                             Inches(rect[2]), Inches(rect[3]))
    if eq.get("annotations"):
        add_annotations(slide, {"annotations": eq["annotations"]}, rect)

    cy = rect[1] + rect[3] + 0.05
    if eq.get("label"):
        caption(slide, (x, cy, w, 0.24), eq["label"])
        cy += 0.26
    for colour, text in captions.items():
        textbox(slide, (x, cy, w, 0.26), text, 12,
                color=RGBColor.from_string(colour.lstrip("#").upper()),
                bold=True, align=PP_ALIGN.CENTER)
        cy += 0.28
    for line in where:
        textbox(slide, (x + 0.10, cy, w - 0.10, 0.22), line, 12, color=GREY)
        cy += 0.24


# ---------------------------------------------------------------- recurring strip

def add_stage(slide, spec, stage_figure, box) -> None:
    """The method-overview figure, repeated, with a box on the stage being told.

    Both reference decks do this on every method slide; naming the figure once
    at the top of the outline keeps it out of every slide entry.
    """
    from PIL import Image

    src = Path(stage_figure["src"])
    if not src.exists():
        raise SystemExit(f"stage_figure not found: {src}")
    name = spec["stage"]
    at = (stage_figure.get("stages") or {}).get(name)
    if at is None:
        raise SystemExit(f"stage {name!r} is not in stage_figure.stages")
    if isinstance(at, dict):
        at = at["at"] if "at" in at else xyxy_to_xywh(at["xyxy"])
    elif at[0] + at[2] > 1.02 or at[1] + at[3] > 1.02:
        raise SystemExit(
            f"stage {name!r} box {at} runs past the picture edge: stages are "
            f"[x, y, w, h] fractions; for corners write {{\"xyxy\": [x0, y0, x1, y1]}}")

    x, y, w, h = box
    iw, ih = Image.open(src).size
    cap = 0.30 if stage_figure.get("caption") else 0.0
    rect = fit((x, y, w, h - cap), iw, ih, vcenter=False)
    slide.shapes.add_picture(str(src), Inches(rect[0]), Inches(rect[1]),
                             Inches(rect[2]), Inches(rect[3]))
    add_annotations(slide, {"annotations": [
        {"type": "box", "at": at, "color": stage_figure.get("color", "C00000")}]}, rect)
    if stage_figure.get("caption"):
        caption(slide, (x, rect[1] + rect[3] + 0.04, w, 0.26), stage_figure["caption"], 11)


OURS_BOX = (0.92, 6.90, 1.9, 0.26)   # bottom-left, clear of the page number


def add_ours(slide, spec) -> None:
    """A small tag that says this slide is the presenter's, not the paper's.

    The lab rule keeps author claim, evidence and presenter interpretation
    visibly separate; this is the visible part. `"ours": true` draws
    `OUR TAKE`; a string draws that string (`OUR EXPERIMENT`, `我們的實驗`).
    """
    o = spec["ours"]
    label = o if isinstance(o, str) else "OUR TAKE"
    x, y, w, h = OURS_BOX
    w = max(w, 0.3 + 0.1 * len(label))
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    strip_style(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = RED
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size, r.font.bold = Pt(10), True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def add_divider(slide, reg) -> None:
    """The thin rule between a figure column and its text, as in the lab decks."""
    if reg["fig"] is None or reg["body"] is None:
        return
    bx, bw = reg["body"][0], reg["body"][2]
    fx = reg["fig"][0]
    x = (bx + bw + fx) / 2 if fx > bx else (fx + reg["fig"][2] + bx) / 2
    top, bottom = reg["span"]
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x), Inches(top + 0.05),
                                      Inches(x), Inches(bottom - 0.05))
    line.line.color.rgb = DIVIDER
    line.line.width = Pt(1.0)


# ---------------------------------------------------------------- balance

SPARSE_LINES = 6        # a text-only slide this short gets larger, airier type
SPARSE_PT = 22
SPARSE_GAP = 14         # pt before each paragraph


def _airy_text(body, spec, reg) -> float | None:
    """A text-only slide with three short bullets used to be three 20pt lines
    under the title and a blank lower half. Give them 22pt and paragraph air;
    `settle` then centres them. Returns the point size used, or None."""
    if (body is None or reg["layout"] != "text-only" or reg["fig"] is not None
            or not spec.get("bullets") or any(spec.get(k) for k in GRAPHICS + ("equation",))):
        return None
    if body_lines(spec, reg["body"][2], SPARSE_PT) > SPARSE_LINES:
        return None
    for i, para in enumerate(body.text_frame.paragraphs):
        if i:
            para.space_before = Pt(SPARSE_GAP)
        for r in para.runs:
            r.font.size = Pt(SPARSE_PT)
    return SPARSE_PT


def _text_height(spec, width: float, pt: float | None) -> float:
    if not spec.get("bullets"):
        return 0.0
    n = body_lines(spec, width, pt or BODY_PT)
    gaps = (len(spec["bullets"]) - 1) * SPARSE_GAP / 72 if pt else 0.0
    return 0.22 + n * (pt or BODY_PT) * 1.25 / 72 + gaps


def _extent(shapes) -> tuple[float, float] | None:
    if not shapes:
        return None
    tops = [Emu(s_.top).inches for s_ in shapes]
    return min(tops), max(t + Emu(s_.height).inches for t, s_ in zip(tops, shapes))


def _shift(shapes, dy: float) -> None:
    if abs(dy) < 0.01:
        return
    for s_ in shapes:
        s_.top = Emu(int(s_.top + Inches(dy)))


def _shift_body(body, dy: float, floor: float) -> None:
    """Move the bullets down by `dy`. The box keeps its height unless that
    would cross `floor` (the region's bottom); then it ends there, which is
    still taller than the text it was measured for."""
    if body is None or abs(dy) < 0.01:
        return
    body.top = Emu(int(body.top + Inches(dy)))
    body.height = Emu(max(int(Inches(0.3)),
                          min(int(body.height), int(Inches(floor)) - int(body.top))))


def settle(spec, reg, body, drawn, sparse_pt=None) -> None:
    """Treat the bullets and the exhibit as one group and place the group.

    Each piece used to be positioned on its own: the text pinned to the top,
    a short exhibit centred in what was left, so a slide came out as three
    lines, a two-inch gap, a strip of boxes and another gap. Now a stacked
    slide closes the gap between text and exhibit, then sits the whole group
    a little above the centre of the content region; a side-by-side slide
    gives both columns one shared top. Nothing grows or shrinks -- only moves.
    `"valign": "top"` on a slide opts out.
    """
    span_top, span_bottom = reg["span"]
    if reg["layout"] == "assertion-evidence":
        span_top = ASSERT_TOP
    has_text = body is not None and reg["body"] is not None and bool(spec.get("bullets"))
    body_top = Emu(body.top).inches if has_text else span_top

    if reg["layout"] in ("figure-right", "figure-left") and reg["fig"] and reg["body"]:
        fx, _, fw, _ = reg["fig"]
        bw = reg["body"][2]

        def in_fig(s_):
            mid = Emu(s_.left).inches + Emu(s_.width).inches / 2
            return fx - 0.05 <= mid <= fx + fw + 0.05

        fig_side = [s_ for s_ in drawn if in_fig(s_)]
        text_side = [s_ for s_ in drawn if not in_fig(s_)]
        t_ext = _extent(text_side)
        t_top = body_top if has_text else (t_ext[0] if t_ext else span_top)
        t_bot = max(body_top + _text_height(spec, bw, sparse_pt) if has_text else t_top,
                    t_ext[1] if t_ext else t_top)
        f_ext = _extent(fig_side)
        tallest = max(t_bot - t_top, (f_ext[1] - f_ext[0]) if f_ext else 0.0)
        top = span_top + max(0.0, (span_bottom - span_top - tallest) * SETTLE)
        _shift(text_side, top - t_top)
        if has_text:
            _shift_body(body, top - t_top, span_bottom)
        if f_ext:
            _shift(fig_side, top - f_ext[0])
        return

    ext = _extent(drawn)
    width = reg["body"][2] if reg["body"] else RIGHT - LEFT
    text_bot = body_top + _text_height(spec, width, sparse_pt) if has_text else span_top
    if ext:
        gap = STACK_GAP if has_text else 0.0
        if ext[0] > text_bot + gap:                 # close the hole under the text
            _shift(drawn, text_bot + gap - ext[0])
        ext = _extent(drawn)
    g_top = body_top if has_text else (ext[0] if ext else span_top)
    g_bot = max(text_bot if has_text else g_top, ext[1] if ext else g_top)
    free = (span_bottom - span_top) - (g_bot - g_top)
    if free <= 0.05:
        return
    dy = span_top + free * SETTLE - g_top
    _shift(drawn, dy)
    if has_text:
        _shift_body(body, dy, span_bottom)


_RPR_AFTER_LATIN = ("ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")


def latin_font(prs, font: str = exhibits.FONT) -> int:
    """Give every run this script drew an explicit Latin face.

    The template's layouts say Calibri and its theme says Arial, so a shape
    python-pptx adds without a typeface renders in Arial beside Calibri
    bullets. Template shapes (Google Slides export names) and placeholders
    already inherit Calibri and are left alone.
    """
    from pptx.oxml.ns import qn

    def runs(shape):
        if shape.shape_type == 6:                                   # group
            for sub in shape.shapes:
                yield from runs(sub)
            return
        if shape.is_placeholder or shape.name.startswith(("Google Shape", "Shape ")):
            return
        if shape.has_text_frame:
            yield from shape.text_frame._txBody.iter(qn("a:r"))
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    yield from cell.text_frame._txBody.iter(qn("a:r"))

    after = {qn("a:" + n) for n in _RPR_AFTER_LATIN}
    n = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            for r_el in runs(shape):
                rPr = r_el.find(qn("a:rPr"))
                if rPr is None:
                    rPr = r_el.makeelement(qn("a:rPr"), {})
                    r_el.insert(0, rPr)
                if rPr.find(qn("a:latin")) is not None:
                    continue
                latin = rPr.makeelement(qn("a:latin"), {"typeface": font})
                anchor = next((c for c in rPr if c.tag in after), None)
                if anchor is not None:
                    anchor.addprevious(latin)
                else:
                    rPr.append(latin)
                n += 1
    return n


# ---------------------------------------------------------------- driver

def flatten(outline: dict) -> list[dict]:
    flat = []
    if outline.get("meta"):
        flat.append({"role": "cover"})
    if outline.get("summary"):
        flat.append({"role": "summary"})
    flat.extend(outline.get("slides", []))
    return flat


VISUAL_KEYS = ("subtitle", "figure", "video", "diagram", "annotations",
               "callout", "matrix", "equation", "stage", "divider",
               "ours") + exhibits.GRAPHIC_KEYS
# `"custom": true` says the user draws this slide's exhibit with their own
# script after the build; compose leaves the region alone (feedback F8)


STAMP = "acm-composed"          # so a second pass cannot silently double every shape


def compose(outline_path: Path, deck: Path, out: Path | None = None,
            force: bool = False) -> Path:
    outline = json.loads(outline_path.read_text(encoding="utf-8"))
    flat = flatten(outline)
    prs = Presentation(str(deck))
    if prs.core_properties.content_status == STAMP and not force:
        raise SystemExit(
            f"{deck} has already been composed - running this again stacks a "
            f"second figure, caption and callout on top of the first. Rebuild "
            f"with build_from_outline.py, then compose. (--force overrides.)")
    if len(prs.slides._sldIdLst) != len(flat):
        raise SystemExit(
            f"{deck} has {len(prs.slides._sldIdLst)} slides but the outline "
            f"describes {len(flat)} - rebuild with build_from_outline.py first")

    touched = 0
    stage_figure = outline.get("stage_figure")
    outline_dir = outline_path.resolve().parent

    for idx, (slide, spec) in enumerate(zip(prs.slides, flat), start=1):
        if not any(spec.get(k) for k in VISUAL_KEYS):
            continue
        touched += 1
        reg = regions(spec)
        if spec.get("subtitle"):
            if reg["layout"] == "assertion-evidence":
                add_assertion(slide, spec)
            else:
                add_subtitle(slide, spec)

        # an equation shares the body column with the bullets, so claim its band
        # before the body placeholder is resized -- that is what stops overlap
        eq_box = None
        if spec.get("equation") and reg["body"]:
            bx, by, bw, bh = reg["body"]
            eq = spec["equation"]
            band = equation_height(eq if isinstance(eq, dict) else {})
            if spec.get("bullets"):
                band = min(band, bh - 0.9)
                eq_y = max(by + body_strip_height(spec, bw) + 0.15, by + 0.9)
                eq_box = (bx, min(eq_y, by + bh - band), bw, band)
                reg["body"] = (bx, by, bw, eq_box[1] - by - 0.05)
            else:
                eq_box = (bx, by, bw, bh)
                reg["body"] = (bx, by, 0.4, 0.3)

        # a table slide's content is the table build_from_outline already filled;
        # there is no bullet region to move, clear or shrink
        body = None if spec.get("table") else find_body(slide, find_title(slide))
        if body is not None and not spec.get("bullets"):
            body.text_frame.clear()
        if body is not None:
            if reg["body"] is None:
                body.text_frame.clear()
                place(body, (LEFT, CONTENT_TOP_PLAIN, 0.4, 0.3))
            else:
                place(body, reg["body"])
                if reg["layout"] in ("figure-right", "figure-left"):
                    for para in body.text_frame.paragraphs:
                        para.alignment = PP_ALIGN.LEFT

        sparse = _airy_text(body, spec, reg)
        before = {shp.shape_id for shp in slide.shapes}
        if eq_box:
            add_equation(slide, spec, eq_box, outline_dir, str(idx))

        # graphic exhibits take the figure slot; without one (text-only) they
        # take the body, which the bullets have already given up
        graphic = [k for k in GRAPHICS if spec.get(k)]
        if len(graphic) > 1:
            raise SystemExit(f"slide {idx} carries {' + '.join(graphic)} - one exhibit "
                             f"per slide; split it")
        if graphic and (spec.get("figure") or spec.get("video")):
            raise SystemExit(f"slide {idx} has {graphic[0]} and a figure/video - one "
                             f"exhibit per slide; split it")
        if graphic and graphic[0] == "diagram" and (spec.get("stage") or spec.get("equation")):
            raise SystemExit(f"slide {idx} has `diagram` plus stage/equation - one "
                             f"exhibit per slide; split them")
        if graphic:
            kind = graphic[0]
            gbox = reg["fig"] or reg["body"]
            if gbox is None:
                raise SystemExit(f"slide {idx}: no room for {kind} in layout {reg['layout']}")
            accent = exhibits.accent_for(spec.get("role", ""))
            if kind == "matrix":
                add_matrix(slide, spec, gbox, accent)
            elif kind == "diagram":
                add_diagram(slide, spec, gbox)
            elif kind == "draw":
                exhibits.run_custom(slide, spec, gbox, accent, outline_dir, idx)
            else:
                exhibits.DRAW[kind](slide, spec, gbox, accent)
        if spec.get("stage") and reg["fig"]:
            if not stage_figure:
                raise SystemExit(f"slide {idx} sets `stage` but the outline has no "
                                 f"top-level `stage_figure`")
            add_stage(slide, spec, stage_figure, reg["fig"])
        elif spec.get("video") and reg["fig"]:
            if spec.get("figure"):
                raise SystemExit(f"slide {idx} has both `figure` and `video` - "
                                 f"one exhibit per slide; split them")
            rect = add_video(slide, spec, reg["fig"],
                             vcenter=reg["layout"] in ("figure-right", "figure-left",
                                                       "assertion-evidence"))
            if spec.get("annotations"):
                add_annotations(slide, spec, rect)
        elif spec.get("figure") and reg["fig"]:
            rect = add_figure(slide, spec, reg["fig"],
                              vcenter=reg["layout"] in ("figure-right", "figure-left",
                                                        "assertion-evidence"))
            if spec.get("annotations"):
                add_annotations(slide, spec, rect)
        if not spec.get("custom") and spec.get("valign") != "top" and not spec.get("table"):
            settle(spec, reg, body,
                   [shp for shp in slide.shapes if shp.shape_id not in before], sparse)
        if spec.get("divider"):
            add_divider(slide, reg)
        if spec.get("callout"):
            add_callout(slide, spec)
        if spec.get("ours"):
            add_ours(slide, spec)

    out = out or deck
    latin_font(prs)
    cjk.tag_deck(prs, cjk.font_from_outline(outline))
    prs.core_properties.content_status = STAMP
    prs.save(str(out))
    print(f"composed {touched}/{len(flat)} slides -> {out}")
    return out


def print_sizes() -> None:
    """Draw a figure to these sizes and its labels stay at the size you set.

    A plot made 6in wide for a 5.75in slot is placed at scale 0.96, so 12pt
    axis labels stay 12pt. Made 12in wide for the same slot, they land at
    6pt (feedback B3).
    """
    print(f"{'layout':20} {'bullets':8} {'exhibit box (w x h in)':24} with subtitle")
    for layout in ("figure-right", "figure-left", "figure-bottom", "figure-full",
                   "assertion-evidence"):
        for bullets in (False, True):
            if layout in ("figure-full", "assertion-evidence") and bullets:
                continue
            spec = {"layout": layout, "figure": "x.png"}
            if bullets:
                spec["bullets"] = ["one line", "two lines"]
            a = regions(spec)["fig"]
            b = regions(dict(spec, subtitle="s"))["fig"]
            print(f"{layout:20} {('2 lines' if bullets else '-'):8} "
                  f"{a[2]:.2f} x {a[3] - CAPTION_H:.2f}{'':13} {b[2]:.2f} x {b[3] - CAPTION_H:.2f}")
    print("(heights already leave room for a caption; a callout takes 0.96in more)")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline", nargs="?")
    ap.add_argument("deck", nargs="?")
    ap.add_argument("-o", "--output", help="default: edit the deck in place")
    ap.add_argument("--force", action="store_true",
                    help="compose a deck that was already composed once")
    ap.add_argument("--sizes", action="store_true",
                    help="print the exhibit box each layout gives, in inches, and exit")
    a = ap.parse_args()
    if a.sizes:
        print_sizes()
        return
    if not a.outline or not a.deck:
        ap.error("outline and deck are required (or --sizes)")
    compose(Path(a.outline), Path(a.deck),
            Path(a.output) if a.output else None, force=a.force)


if __name__ == "__main__":
    main()
