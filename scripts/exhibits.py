#!/usr/bin/env python3
"""Graphic exhibits for compose.py: cards, flow, big numbers, a positioning
map, and the custom-draw hook.

Before these existed a slide could carry bullets, a table or one picture, so
every pipeline and every comparison was typed into a table -- which is why a
sixteen-slide paper talk came out as seven tables and two bullet lists. Each
function here takes the exhibit box compose.py computed for the layout and
draws inside it, so nothing overlaps anything else on the slide.

Every function has the same shape:

    draw_<kind>(slide, spec, box, accent) -> None

`box` is (x, y, w, h) in inches. `accent` is the section colour of the slide's
role, so the shapes match the template's tab instead of fighting it.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- palette
# the template's own section tabs, read from its layouts (theme accent6,
# accent5, and two literal fills); anything drawn on a slide takes the tab
# colour of that slide's section so the deck still reads as one template
SECTION = {
    "project":  "70AD47",
    "research": "5B9BD5",
    "paper":    "C198E0",
    "others":   "F89AB5",
}
RED = "C00000"
INK = "1A1A1A"
GREY = "595959"
LINE = "BFC5CC"
PANEL = "F4F6F8"
WHITE = "FFFFFF"


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.lstrip("#").upper())


def accent_for(role: str) -> str:
    for key, col in SECTION.items():
        if role.startswith(key):
            return col
    return SECTION["others"]


def tint(hexstr: str, amount: float = 0.82) -> str:
    """Mix a colour toward white; 0.82 is a pale panel, 0.5 a mid tint."""
    h = hexstr.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    mix = lambda c: int(round(c + (255 - c) * amount))
    return f"{mix(r):02X}{mix(g):02X}{mix(b):02X}"


# ---------------------------------------------------------------- primitives

def _strip_style(shape) -> None:
    sp = shape._element
    for st in sp.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}style'):
        sp.remove(st)


def rect(slide, box, *, fill=None, line=None, width=1.0, rounded=False, shape=None):
    x, y, w, h = box
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE)
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    _strip_style(shp)
    if rounded:
        shp.adjustments[0] = 0.08
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(width)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def write(shape_or_slide, text_lines, *, box=None, size=14, color=INK, bold=False,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, margin=0.08, italic=False):
    """Fill a shape's text frame, or add a textbox when given a slide and box.

    `text_lines` is a str or a list; a list item may be a dict
    {text, size, bold, color} to vary one line.
    """
    if box is not None:
        x, y, w, h = box
        shp = shape_or_slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        shp = shape_or_slide
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    lines = [text_lines] if isinstance(text_lines, str) else list(text_lines)
    first = True
    for item in lines:
        d = item if isinstance(item, dict) else {"text": str(item)}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        r = p.add_run()
        r.text = d.get("text", "")
        r.font.size = Pt(d.get("size", size))
        r.font.bold = d.get("bold", bold)
        r.font.italic = d.get("italic", italic)
        r.font.color.rgb = rgb(d.get("color", color))
    return shp


def arrow(slide, x0, y0, x1, y1, *, color=GREY, width=1.75):
    from pptx.oxml.ns import qn
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    conn.line.color.rgb = rgb(color)
    conn.line.width = Pt(width)
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return conn


def _fit_size(text: str, width_in: float, base: float, floor: float = 11) -> float:
    """Shrink a one-line label until it fits its width, down to `floor` pt."""
    size = base
    while size > floor and len(text) * size * 0.55 / 72 > width_in - 0.2:
        size -= 1
    return size


# ---------------------------------------------------------------- cards

def draw_cards(slide, spec, box, accent) -> None:
    """2-6 panels in a row (or two rows), each with a heading and a few lines.

        "cards": {
          "items": [
            {"title": "Overfitting priors", "text": "Train per identity; 10-400 min",
             "tag": "camp 1", "bullets": ["..."], "highlight": false, "color": "70AD47"},
            ...
          ],
          "columns": 3,            // optional; default fits <=4 in one row
          "style": "panel"         // "panel" (tinted) | "outline" | "header"
        }

    Use for: the camps in related work, the contributions, strengths vs
    weaknesses, before/after, the four questions about a method component.
    """
    c = spec["cards"]
    items = c["items"] if isinstance(c, dict) else list(c)
    style = (c.get("style") if isinstance(c, dict) else None) or "header"
    n = len(items)
    if not 1 <= n <= 6:
        raise SystemExit(f"cards: {n} items; use 2-6 (one idea per card)")
    cols = (c.get("columns") if isinstance(c, dict) else None) or (n if n <= 4 else 3)
    rows = -(-n // cols)
    x, y, w, h = box
    gap = 0.22
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    head_h = 0.52 if ch > 1.6 else 0.42
    body_pt = 15 if cw >= 3.6 else 13
    if rows > 1 and ch < 2.2:
        body_pt = min(body_pt, 13)

    for i, it in enumerate(items):
        it = {"title": it} if isinstance(it, str) else it
        r, col = divmod(i, cols)
        cx, cy = x + col * (cw + gap), y + r * (ch + gap)
        colour = it.get("color") or (RED if it.get("highlight") else accent)
        if style == "panel":
            rect(slide, (cx, cy, cw, ch), fill=tint(colour, 0.86),
                 line=colour if it.get("highlight") else None, width=2.0, rounded=True)
        else:
            rect(slide, (cx, cy, cw, ch), fill=WHITE, line=colour,
                 width=2.25 if it.get("highlight") else 1.25, rounded=True)
        if style == "header":
            rect(slide, (cx, cy, cw, head_h), fill=colour, rounded=False)
            head_color = WHITE
        else:
            head_color = colour
        title_pt = _fit_size(it.get("title", ""), cw, 17 if ch > 1.6 else 15, 12)
        write(slide, it.get("title", ""), box=(cx, cy, cw, head_h), size=title_pt,
              bold=True, color=head_color, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
        if it.get("tag"):
            write(slide, it["tag"], box=(cx + cw - 1.3, cy + head_h + 0.04, 1.26, 0.28),
                  size=10, color=colour, bold=True, align=PP_ALIGN.RIGHT, margin=0.06)
        lines = []
        if it.get("text"):
            lines.append({"text": it["text"], "size": body_pt})
        for b in it.get("bullets") or []:
            lines.append({"text": "• " + str(b), "size": body_pt})
        if it.get("foot"):
            lines.append({"text": it["foot"], "size": body_pt - 2, "color": GREY, "italic": True})
        if lines:
            write(slide, lines, box=(cx, cy + head_h + 0.06, cw, ch - head_h - 0.10),
                  size=body_pt, margin=0.14)


# ---------------------------------------------------------------- flow

def draw_flow(slide, spec, box, accent) -> None:
    """A pipeline: steps left-to-right joined by arrows, one step lit.

        "flow": {
          "steps": [
            {"label": "FLAME tracking", "sub": "UV maps", "highlight": false},
            "MGPM", {"label": "Enhancer", "sub": "3 real frames", "highlight": true},
            ...
          ],
          "highlight": 3,          // 1-based, alternative to per-step flags
          "direction": "row",      // "row" (default) | "column"
          "style": "box"           // "box" (rounded) | "chevron"
        }

    Two to seven steps. More than that is a figure, not a flow.
    """
    f = spec["flow"]
    steps = f["steps"] if isinstance(f, dict) else list(f)
    hi = f.get("highlight") if isinstance(f, dict) else None
    direction = (f.get("direction") if isinstance(f, dict) else None) or "row"
    style = (f.get("style") if isinstance(f, dict) else None) or "box"
    n = len(steps)
    if not 2 <= n <= 7:
        raise SystemExit(f"flow: {n} steps; use 2-7 (crop the paper's figure past that)")
    x, y, w, h = box
    norm = [{"label": s} if isinstance(s, str) else dict(s) for s in steps]
    for i, s in enumerate(norm, start=1):
        s["highlight"] = s.get("highlight") or (hi == i)
    has_sub = any(s.get("sub") for s in norm)

    if direction == "column":
        gap = 0.28
        sh = min(0.95, (h - gap * (n - 1)) / n)
        sw = min(w, 5.2)
        sx = x + (w - sw) / 2
        top = y + (h - (sh * n + gap * (n - 1))) / 2
        for i, s in enumerate(norm):
            sy = top + i * (sh + gap)
            _flow_step(slide, s, (sx, sy, sw, sh), accent, has_sub, style)
            if i < n - 1:
                arrow(slide, sx + sw / 2, sy + sh + 0.02, sx + sw / 2, sy + sh + gap - 0.02)
        return

    gap = 0.34 if style == "box" else 0.08
    sw = (w - gap * (n - 1)) / n
    sh = min(h, 1.45 if has_sub else 1.05)
    sy = y + (h - sh) / 2
    for i, s in enumerate(norm):
        sx = x + i * (sw + gap)
        _flow_step(slide, s, (sx, sy, sw, sh), accent, has_sub, style)
        if style == "box" and i < n - 1:
            arrow(slide, sx + sw + 0.03, sy + sh / 2, sx + sw + gap - 0.03, sy + sh / 2)
    if isinstance(f, dict) and f.get("caption"):
        write(slide, f["caption"], box=(x, sy + sh + 0.08, w, 0.3), size=11,
              color=GREY, italic=True, align=PP_ALIGN.CENTER)


def _flow_step(slide, s, box, accent, has_sub, style) -> None:
    x, y, w, h = box
    colour = s.get("color") or (RED if s.get("highlight") else accent)
    if style == "chevron":
        shp = rect(slide, box, fill=colour if s.get("highlight") else tint(accent, 0.55),
                   shape=MSO_SHAPE.CHEVRON)
        fg = WHITE
    else:
        shp = rect(slide, box, fill=tint(colour, 0.84) if s.get("highlight") else WHITE,
                   line=colour, width=2.5 if s.get("highlight") else 1.5, rounded=True)
        fg = colour if s.get("highlight") else INK
    label_pt = _fit_size(s["label"], w - (0.4 if style == "chevron" else 0), 16, 11)
    lines = [{"text": s["label"], "size": label_pt, "bold": True, "color": fg}]
    if s.get("sub"):
        lines.append({"text": s["sub"], "size": max(10, label_pt - 4),
                      "color": fg if style == "chevron" else GREY})
    write(shp, lines, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
          margin=0.22 if style == "chevron" else 0.08)


# ---------------------------------------------------------------- big numbers

def draw_bignum(slide, spec, box, accent) -> None:
    """One to four numbers that carry the result, at a size you can read
    from the back of the room.

        "bignum": {
          "items": [
            {"value": "20 min", "label": "per-identity fitting",
             "sub": "vs 400 min CAP4D", "highlight": true},
            {"value": "+0.38", "label": "PSNR over SplattingAvatar"}
          ]
        }

    Every value must come from the paper or your own run; the slide that
    shows a number also names its baseline.
    """
    b = spec["bignum"]
    items = b["items"] if isinstance(b, dict) else list(b)
    n = len(items)
    if not 1 <= n <= 4:
        raise SystemExit(f"bignum: {n} items; one to four, or it is a table")
    x, y, w, h = box
    gap = 0.3
    cw = (w - gap * (n - 1)) / n
    value_pt = 54 if n <= 2 else (44 if n == 3 else 36)
    for i, it in enumerate(items):
        cx = x + i * (cw + gap)
        colour = it.get("color") or (RED if it.get("highlight") else accent)
        rect(slide, (cx, y, cw, h), fill=PANEL, rounded=True)
        rect(slide, (cx, y, 0.09, h), fill=colour)
        v_pt = _fit_size(str(it["value"]), cw - 0.3, value_pt, 24)
        lines = [{"text": str(it["value"]), "size": v_pt, "bold": True, "color": colour}]
        if it.get("label"):
            lines.append({"text": it["label"], "size": 16, "bold": True})
        if it.get("sub"):
            lines.append({"text": it["sub"], "size": 13, "color": GREY})
        write(slide, lines, box=(cx + 0.12, y, cw - 0.2, h), align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE, margin=0.1)


# ---------------------------------------------------------------- quadrant

def draw_quadrant(slide, spec, box, accent) -> None:
    """A positioning map: two labelled axes, prior work as dots, ours in red.

        "quadrant": {
          "x": ["slow", "fast"],               // axis label pair (left, right)
          "y": ["low identity", "high identity"],
          "points": [
            {"label": "CAP4D", "at": [0.15, 0.80]},
            {"label": "FlashAvatar", "at": [0.85, 0.20]},
            {"label": "ELITE (ours)", "at": [0.80, 0.85], "highlight": true}
          ],
          "quadrants": ["", "goal", "", ""]    // optional corner labels: TL TR BL BR
        }

    `at` is a fraction of the plot: [0,0] bottom-left, [1,1] top-right, like
    any chart. Place by argument, not by data -- it is a map of the field,
    and the paper's own position is the point of the slide.
    """
    q = spec["quadrant"]
    x, y, w, h = box
    lab_h, lab_w = 0.40, 1.10
    px, py = x + lab_w, y + 0.15
    pw, ph = w - lab_w - 0.2, h - lab_h - 0.25
    # frame and midlines
    rect(slide, (px, py, pw, ph), fill="FAFBFC", line=LINE, width=1.0)
    for frac in (0.5,):
        _dash(slide, px + pw * frac, py, px + pw * frac, py + ph)
        _dash(slide, px, py + ph * frac, px + pw, py + ph * frac)
    # axes with arrows and their labels
    arrow(slide, px, py + ph + 0.02, px + pw, py + ph + 0.02, color=INK, width=1.5)
    arrow(slide, px - 0.02, py + ph, px - 0.02, py, color=INK, width=1.5)
    xl = q.get("x") or ["", ""]
    yl = q.get("y") or ["", ""]
    write(slide, xl[0], box=(px, py + ph + 0.06, pw / 2, lab_h), size=12, color=GREY)
    write(slide, xl[1], box=(px + pw / 2, py + ph + 0.06, pw / 2, lab_h), size=12,
          color=GREY, align=PP_ALIGN.RIGHT)
    write(slide, yl[1], box=(x, py - 0.05, lab_w - 0.1, 0.6), size=12, color=GREY,
          align=PP_ALIGN.RIGHT)
    write(slide, yl[0], box=(x, py + ph - 0.55, lab_w - 0.1, 0.6), size=12, color=GREY,
          align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM)
    if q.get("title_x"):
        write(slide, q["title_x"], box=(px, py + ph + 0.06, pw, lab_h), size=12,
              bold=True, align=PP_ALIGN.CENTER)
    corners = q.get("quadrants") or []
    for txt, (cx, cy, al) in zip(corners, [
            (px + 0.08, py + 0.06, PP_ALIGN.LEFT), (px + pw / 2, py + 0.06, PP_ALIGN.RIGHT),
            (px + 0.08, py + ph - 0.42, PP_ALIGN.LEFT), (px + pw / 2, py + ph - 0.42, PP_ALIGN.RIGHT)]):
        if txt:
            write(slide, txt, box=(cx, cy, pw / 2 - 0.12, 0.36), size=11, bold=True,
                  color=accent, align=al, italic=True)
    # points
    d = 0.26
    for p in q.get("points") or []:
        fx, fy = p["at"]
        if not (0 <= fx <= 1 and 0 <= fy <= 1):
            raise SystemExit(f"quadrant point {p.get('label')!r} is outside [0,1]: {p['at']}")
        cx = px + fx * pw - d / 2
        cy = py + (1 - fy) * ph - d / 2
        colour = p.get("color") or (RED if p.get("highlight") else accent)
        rect(slide, (cx, cy, d, d), fill=colour, line=WHITE, width=1.0, shape=MSO_SHAPE.OVAL)
        lw = min(2.4, max(1.2, 0.09 * len(p.get("label", "")) + 0.3))
        fits_right = cx + d + 0.04 + lw <= x + w
        side = p.get("side") or ("right" if fits_right else "left")
        if side == "right":
            write(slide, p.get("label", ""), box=(cx + d + 0.04, cy - 0.08, lw, 0.42),
                  size=12, bold=bool(p.get("highlight")),
                  color=colour if p.get("highlight") else INK, margin=0.02)
        else:
            write(slide, p.get("label", ""), box=(cx - lw - 0.04, cy - 0.08, lw, 0.42),
                  size=12, bold=bool(p.get("highlight")), align=PP_ALIGN.RIGHT,
                  color=colour if p.get("highlight") else INK, margin=0.02)


def _dash(slide, x0, y0, x1, y1) -> None:
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    conn.line.color.rgb = rgb(LINE)
    conn.line.width = Pt(0.75)
    conn.line.dash_style = 4


# ---------------------------------------------------------------- custom hook

def run_custom(slide, spec, box, accent, outline_dir: Path, index: int) -> None:
    """`"draw": "design.py:draw_timeline"` -- your own function on this slide.

    The function is called as

        fn(slide, box, spec, kit)

    where `box` is the exhibit region (x, y, w, h in inches) that the layout
    left for it, `spec` the slide's outline entry, and `kit` a namespace with
    the helpers this module uses (rect, write, arrow, rgb, tint, accent,
    Inches, Pt, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE). Draw inside `box`; the
    title, subtitle, bullets and callout are already placed around it.
    `assets/examples/design.py` is the reference implementation.
    """
    target = spec["draw"]
    if ":" not in target:
        raise SystemExit(f"slide {index}: `draw` must be 'file.py:function', got {target!r}")
    mod_path, fn_name = target.rsplit(":", 1)
    path = Path(mod_path)
    if not path.is_absolute():
        path = (outline_dir / path) if not path.exists() else path
    if not path.exists():
        raise SystemExit(f"slide {index}: draw module not found: {path}")
    name = f"acm_custom_{path.stem}_{abs(hash(str(path.resolve()))) % 10**6}"
    if name in sys.modules:
        mod = sys.modules[name]
    else:
        spec_ = importlib.util.spec_from_file_location(name, path)
        mod = importlib.util.module_from_spec(spec_)
        sys.modules[name] = mod
        spec_.loader.exec_module(mod)
    fn = getattr(mod, fn_name, None)
    if fn is None:
        raise SystemExit(f"slide {index}: {path.name} has no function {fn_name!r}")
    kit = _Kit(accent)
    fn(slide, box, spec, kit)


class _Kit:
    """What a custom draw function gets handed, so it needs no imports."""
    rect = staticmethod(rect)
    write = staticmethod(write)
    arrow = staticmethod(arrow)
    rgb = staticmethod(rgb)
    tint = staticmethod(tint)
    Inches, Pt = Inches, Pt
    PP_ALIGN, MSO_ANCHOR, MSO_SHAPE = PP_ALIGN, MSO_ANCHOR, MSO_SHAPE
    RED, INK, GREY, LINE, PANEL, WHITE = RED, INK, GREY, LINE, PANEL, WHITE
    SECTION = SECTION

    def __init__(self, accent: str):
        self.accent = accent


GRAPHIC_KEYS = ("cards", "flow", "bignum", "quadrant", "draw")

DRAW = {
    "cards": draw_cards,
    "flow": draw_flow,
    "bignum": draw_bignum,
    "quadrant": draw_quadrant,
}
