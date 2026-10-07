#!/usr/bin/env python3
"""Gate a deck before you show it. Exits 1 on ERROR, 0 on WARN only.

    python scripts/qa_check.py outline.json talk.pptx

Every ERROR here traces to lab-rules.md or to what the two reference decks
actually do (assets/examples/reference-decks.md). Everything softer is a
WARN and is your call -- the reference decks themselves trip some of them,
which is exactly why they are not errors.
"""
from __future__ import annotations

import argparse
import functools
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).resolve().parent))

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "acm_template.pptx"

BODY_WORD_CAP = 40          # lab-rules.md: "~40 words of body text per slide, hard"
NOTES_MIN_SENTENCES = 3     # WARN only: SKILL.md asks for three to six
TEXT_ONLY_MAX = 0.25          # WARN only: SliderEdit sits at 24% (4/17)
CALLOUT_MAX = 0.15            # WARN only: both reference decks use zero
CLAIM_MIN_WORDS = 4

# roles that are structural, not argument-bearing
EXEMPT = {"cover", "summary", "paper_list", "project_summary", "research_summary"}

# What each functional slide type must carry. PPTAgent's Stage I calls this a
# content schema; here it is the machine-checkable half of slide-patterns.md.
EXHIBIT_KEYS = ("figure", "video", "diagram", "matrix", "equation", "stage",
                "table", "cards", "flow", "bignum", "quadrant", "draw", "custom")
CAPTION_MAX = 60              # WARN: characters (CJK counts double) before a caption shrinks
TABLE_KEYS = ("matrix", "table")
TABLE_MAX = 0.35              # WARN: past this the deck is a spreadsheet
RUN_MAX = 3                   # WARN: this many table-only or text-only slides in a row
SQUEEZED_W = 7.0              # WARN: a wide figure placed narrower than this
FILL_MIN = 0.50               # WARN: share of the content height that carries anything
GAP_MAX = 1.6                 # WARN: an empty band across the slide taller than this (in)
ROLE_SCHEMA = {
    "paper_method":     {"exhibit": "error"},
    "paper_results":    {"exhibit": "error"},
    "paper_related":    {"exhibit": "warn"},
    "paper_intro":      {"exhibit": "warn"},
    "paper_conclusion": {"bullets": "warn"},
    "project_results":  {"exhibit": "error"},
    "research_results": {"exhibit": "error"},
}

NON_CLAIM = re.compile(
    r"^(results?|method|methods|introduction|conclusion|related work|"
    r"experiments?|ablation|overview|background|motivation)$", re.I)


CJK_PER_WORD = 1.8          # one English word of slide text is ~1.8 中文字


def cjk_count(text: str) -> int:
    return len(re.findall(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff]", text or ""))


def words(text: str) -> int:
    """Word-equivalents: latin per token, CJK at 1.8 characters per word.

    Counting every 中文字 as a word made the 40-word cap two lines of Chinese
    (feedback A2); 1.8 puts a Chinese body at the same visual density as an
    English one, about 70 characters.
    """
    cjk = len(re.findall(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z0-9][A-Za-z0-9\-/%.+]*", text))
    return latin + int(round(cjk / CJK_PER_WORD))


def flatten(outline: dict) -> list[dict]:
    flat = []
    if outline.get("meta"):
        flat.append({"role": "cover"})
    if outline.get("summary"):
        flat.append({"role": "summary"})
    flat.extend(outline.get("slides", []))
    return flat


def bullet_text(spec: dict) -> str:
    out = []
    for b in spec.get("bullets") or []:
        out.append(b["text"] if isinstance(b, dict) else str(b))
    return " ".join(out)


@functools.lru_cache(maxsize=32)
def _video_issues(src: str) -> tuple[tuple[str, str], ...]:
    """Cached: probing a file costs an ffprobe subprocess."""
    try:
        import video as video_mod
    except ImportError:                      # video.py travels with this script
        return ()
    path = Path(src)
    if not path.exists():
        return (("error", f"video not found: {path}"),)
    return tuple(video_mod.problems(path))


def check_outline(flat: list[dict]) -> tuple[list[str], list[str]]:
    errors, warns = [], []
    content = [s for s in flat if s.get("role") not in EXEMPT]
    text_only = 0
    n_callout = 0

    for i, spec in enumerate(flat, start=1):
        role = spec.get("role", "?")
        tag = f"slide {i} ({role})"
        if role in EXEMPT:
            continue

        # 1. the slide must say something -- a specific title is enough on its
        # own. Both reference decks carry the argument in the title
        # (`Time-Aware Encoder (TAE)`, `Continuous Control by Scaling LoRA`);
        # SliderEdit uses no red line and no box on any of its 17 slides.
        title = (spec.get("title") or "").strip()
        claim = spec.get("subtitle") or ""
        if not claim and isinstance(spec.get("callout"), dict):
            claim = spec["callout"].get("text", "")
        elif not claim and isinstance(spec.get("callout"), str):
            claim = spec["callout"]
        claim = claim.strip()

        generic = bool(NON_CLAIM.match(title)) or not title
        if generic and not claim:
            errors.append(f"{tag}: generic title {title!r} says nothing and there "
                          f"is no `subtitle` to carry the point - name the method, "
                          f"or state the finding in a subtitle")
        elif claim and words(claim) < CLAIM_MIN_WORDS:
            warns.append(f"{tag}: claim {claim!r} reads as a label, not a statement")

        if spec.get("callout"):
            n_callout += 1

        # 2. body word ceiling -- slide body only, never captions or notes
        n = words(bullet_text(spec))
        if n > BODY_WORD_CAP:
            errors.append(f"{tag}: {n} words of body text (cap {BODY_WORD_CAP}) "
                          f"- split the slide or move detail into notes")

        # 3. borrowed figures need a source
        fig = spec.get("figure")
        if isinstance(fig, dict):
            if not fig.get("caption"):
                warns.append(f"{tag}: figure has no caption")
            elif len(fig["caption"]) + cjk_count(fig["caption"]) > CAPTION_MAX:
                warns.append(f"{tag}: caption is {len(fig['caption'])} characters - past "
                             f"~{CAPTION_MAX} it renders small; move the detail into "
                             f"bullets or notes (feedback E4)")
            if isinstance(fig.get("src"), list) and len(fig["src"]) != 2:
                errors.append(f"{tag}: figure.src as a list takes exactly two paths")
            if not (fig.get("source") or "Fig" in (fig.get("caption") or "")):
                errors.append(f"{tag}: borrowed figure with no source "
                              f"- lab-rules.md requires citing borrowed visuals")
        # 3b. an embedded video is an exhibit with a second failure mode: it
        # can be present, correctly captioned, and still be a black rectangle
        # in the meeting because PowerPoint cannot decode it.
        vid = spec.get("video")
        if vid:
            v = {"src": vid} if isinstance(vid, str) else vid
            for level, msg in _video_issues(str(v.get("src", ""))):
                (errors if level == "error" else warns).append(f"{tag}: {msg}")
            if not v.get("caption"):
                warns.append(f"{tag}: video has no caption")
            if not v.get("source"):
                errors.append(f"{tag}: video with no source - say whose it is "
                              f"(\"ours\", or the paper it came from)")

        # 3c. diagrams are native PowerPoint objects, not figure bitmaps.  The
        # topology is explicit so an unknown endpoint cannot silently become a
        # floating arrow in the deck.
        diagram = spec.get("diagram")
        if diagram:
            if not isinstance(diagram, dict):
                errors.append(f"{tag}: `diagram` must be an object")
            else:
                nodes = diagram.get("nodes") or []
                edges = diagram.get("edges") or []
                ids = [str(n.get("id", "")) for n in nodes if isinstance(n, dict)]
                if len(nodes) < 2:
                    errors.append(f"{tag}: diagram needs at least two nodes")
                if len(nodes) > 20 or len(edges) > 32:
                    errors.append(f"{tag}: diagram is too dense for one slide "
                                  f"(max 20 nodes / 32 edges)")
                if len(ids) != len(nodes) or any(not node_id for node_id in ids) \
                        or len(ids) != len(set(ids)):
                    errors.append(f"{tag}: diagram node ids must be present and unique")
                known = set(ids)
                for edge in edges:
                    if not isinstance(edge, dict) or edge.get("from") not in known \
                            or edge.get("to") not in known:
                        errors.append(f"{tag}: diagram edge references an unknown node: {edge}")
                other = [k for k in ("figure", "video", "stage", "matrix", "equation")
                         if spec.get(k)]
                if other:
                    errors.append(f"{tag}: native `diagram` cannot share a slide with "
                                  f"{', '.join(other)} - split the exhibits")

        if not any(spec.get(k) for k in EXHIBIT_KEYS):
            text_only += 1

        # 3c. annotation boxes: `at` is [x, y, w, h]; `xyxy` is corners. A box
        # that runs past the picture edge is the signature of corners written
        # into `at`, which is the mistake figure.py's crop syntax invites.
        for a in spec.get("annotations") or []:
            kind = a.get("type", "box")
            if "xyxy" in a:
                x0, y0, x1, y1 = a["xyxy"]
                if x1 <= x0 or y1 <= y0:
                    errors.append(f"{tag}: annotation xyxy {a['xyxy']} has its corners "
                                  f"the wrong way round")
            elif kind in ("box", "circle") and "at" in a:
                x, y, w, h = a["at"]
                if x + w > 1.02 or y + h > 1.02:
                    errors.append(f"{tag}: annotation `at` {a['at']} runs past the "
                                  f"picture edge - `at` is [x, y, w, h]; for corners "
                                  f"(as figure.py crop takes them) write `xyxy`")
            elif "at" not in a and "xyxy" not in a:
                errors.append(f"{tag}: annotation has neither `at` nor `xyxy`")

        # 3d. a wide figure in a half column is unreadable: compute where it
        # lands and say so, instead of leaving it for the render
        single = isinstance(fig, dict) and isinstance(fig.get("src"), str)
        if single and Path(fig["src"]).exists():
            placed = _placed_figure(spec, fig["src"])
            if placed is not None:
                pw, ph, aspect = placed
                if aspect >= 1.8 and pw < SQUEEZED_W:
                    warns.append(f"{tag}: figure is {aspect:.1f}:1 but lands {pw:.1f}in "
                                 f"wide in layout {spec.get('layout')!r} - its labels "
                                 f"will not read; use figure-bottom / figure-full, or "
                                 f"crop to the component this slide is about")

        # 3e. a custom draw function must exist before the build is attempted
        if spec.get("draw"):
            msg = _draw_target_issue(spec["draw"])
            if msg:
                errors.append(f"{tag}: {msg}")

        # 4. what this functional slide type must carry
        schema = ROLE_SCHEMA.get(role, {})
        for key, severity in schema.items():
            if key == "exhibit":
                ok = any(spec.get(k) for k in EXHIBIT_KEYS)
                msg = (f"{tag}: a {role} slide with no exhibit - use a figure, "
                       f"native diagram, matrix or equation")
            else:
                # the takeaways may be cards instead of bullets
                ok = bool(spec.get(key)) or (key == "bullets" and bool(spec.get("cards")))
                msg = f"{tag}: {role} slide is missing `{key}`"
            if not ok:
                (errors if severity == "error" else warns).append(msg)

        # assertion-evidence slides carry their evidence visually, never as bullets
        if spec.get("layout") in ("assertion-evidence", "ae"):
            if spec.get("bullets"):
                errors.append(f"{tag}: assertion-evidence layout cannot take bullets "
                              f"- the claim is the headline, the visual is the evidence")
            if not any(spec.get(k) for k in EXHIBIT_KEYS):
                errors.append(f"{tag}: assertion-evidence layout with no evidence")
            if not spec.get("subtitle"):
                errors.append(f"{tag}: assertion-evidence layout needs `subtitle` "
                              f"- that is the assertion")

        # 5. speaker notes
        notes = (spec.get("notes") or "").strip()
        n_sent = len([x for x in re.split(r"[。．.!?！？;；\n]+", notes) if x.strip()])
        if not notes:
            warns.append(f"{tag}: no speaker notes")
        elif n_sent < NOTES_MIN_SENTENCES:
            warns.append(f"{tag}: notes are {n_sent} sentence(s) - the rule is "
                         f"three to six, in 中文, full sentences")

    if content and n_callout > max(2, round(len(content) * CALLOUT_MAX)):
        warns.append(f"{n_callout} callout boxes across {len(content)} content slides "
                     f"- TADSR uses 0 and SliderEdit uses 0. A box on every slide "
                     f"reads as a tic; keep it for the one or two that pivot.")

    if content:
        ratio = text_only / len(content)
        if ratio > TEXT_ONLY_MAX:
            warns.append(f"{text_only}/{len(content)} content slides carry no exhibit "
                         f"({ratio:.0%}); the reference decks sit at 8-12%")
    warns.extend(texture_warnings(flat))
    return errors, warns


def _placed_figure(spec: dict, src: str):
    """(width, height, aspect) the figure gets under this slide's layout."""
    try:
        import compose
        from PIL import Image
    except ImportError:
        return None
    try:
        iw, ih = Image.open(src).size
    except Exception:
        return None
    reg = compose.regions(spec)
    if not reg.get("fig"):
        return None
    x, y, w, h = reg["fig"]
    if isinstance(spec.get("figure"), dict) and spec["figure"].get("caption"):
        h -= compose.CAPTION_H
    _, _, fw, fh = compose.fit((x, y, w, h), iw, ih)
    return fw, fh, iw / ih


def _draw_target_issue(target: str) -> str | None:
    if ":" not in str(target):
        return f"`draw` must be 'file.py:function', got {target!r}"
    mod, fn = str(target).rsplit(":", 1)
    if not Path(mod).exists():
        return f"draw module {mod!r} not found (resolved from the working directory)"
    import re as _re
    if not _re.search(rf"^def\s+{_re.escape(fn)}\s*\(", Path(mod).read_text(encoding="utf-8"), _re.M):
        return f"{mod} defines no function {fn!r}"
    return None


def texture_warnings(flat: list[dict]) -> list[str]:
    """The deck-level failure the per-slide checks cannot see: monotony.

    Sixteen slides of which seven are tables and two are bare bullets is a
    spreadsheet with a title band, and it passed every per-slide rule. So:
    cap the table share, break runs of the same shape, and ask a paper talk
    to carry at least one slide of the presenter's own.
    """
    out = []
    content = [s for s in flat if s.get("role") not in EXEMPT]
    if not content:
        return out

    def shape(spec):
        if any(spec.get(k) for k in TABLE_KEYS):
            return "table"
        if any(spec.get(k) for k in EXHIBIT_KEYS):
            return "graphic"
        return "text"

    shapes = [shape(s) for s in content]
    n_table = shapes.count("table")
    if n_table / len(content) > TABLE_MAX and n_table >= 3:
        out.append(f"{n_table}/{len(content)} content slides are tables - a pipeline "
                   f"is a `flow` (or a `diagram` when it branches), camps or contributions are `cards`, the decisive "
                   f"number is a `bignum`, the field is a `quadrant`; keep `matrix` "
                   f"for the one comparison where the cells are the argument")
    run_kind, run_len, run_start = None, 0, 0
    for i, k in enumerate(shapes + [None]):
        if k == run_kind:
            run_len += 1
            continue
        if run_kind in ("table", "text") and run_len >= RUN_MAX:
            first = flat.index(content[run_start]) + 1
            what = "table" if run_kind == "table" else "text-only"
            out.append(f"slides {first}-{first + run_len - 1}: {run_len} {what} slides in "
                       f"a row - the audience stops reading by the third; turn one "
                       f"into a figure, flow, cards or big numbers")
        run_kind, run_len, run_start = k, 1, i

    is_paper = any(str(s.get("role", "")).startswith("paper_") for s in content)
    if is_paper and not any(s.get("ours") for s in content):
        out.append("a paper talk with no slide marked `\"ours\": true` - the lab wants "
                   "to hear what you tried, what it means for our work, or your own "
                   "critique; put that on its own slide and mark it")
    return out


# Tokens the template ships as "fill me in". Any of these surviving into the
# deck is the single most embarrassing thing that can happen in a lab meeting,
# and it is the one thing a render is usually opened to look for -- so look for
# it here instead, for free.
PLACEHOLDER = re.compile(
    r"\bXXX\b|\b20XX\b|\bUr Name\b|Conf\.Name|Sponsor\.Name|Proj\.Name"
    r"|\bPaper Title\b|lorem ipsum", re.I)

MARGIN_BOTTOM, MARGIN_RIGHT = 7.2, 13.1
AUTOFIT_HARD = 1.8              # past this even PowerPoint's autofit clips


def _needed_height(text: str, width_in: float, pt: float, n_paras: int) -> float:
    """Inches of box this text wants. Latin glyphs ~0.52em, CJK 1.0em, 1.22 leading."""
    cjk = len(re.findall(r"[\u4e00-\u9fff\u3040-\u30ff]", text))
    ems = (len(text) - cjk) * 0.52 + cjk * 1.0
    per_line = max(1.0, (width_in * 72) / pt)
    lines = max(n_paras, ems / per_line)
    return lines * pt * 1.22 / 72


A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _effective_pt(tf) -> float:
    """The size this text actually renders at, not the one python-pptx exposes.

    A template placeholder carries no `sz` on its runs -- the size comes down
    from the layout, which python-pptx will not resolve. This template is a
    Google Slides export, and those write `a:buSzPts` alongside every paragraph
    with the same value as the run size (54pt title, 20pt body, verified
    against every slide that does carry an explicit `sz`). So: explicit run
    size, then the paragraph default, then the bullet size, then give up.
    """
    el = tf._txBody
    sizes = [int(r.get("sz")) for r in el.iter(A_NS + "rPr") if r.get("sz")]
    if sizes:
        return max(sizes) / 100.0
    sizes = [int(r.get("sz")) for r in el.iter(A_NS + "defRPr") if r.get("sz")]
    if sizes:
        return max(sizes) / 100.0
    sizes = [int(b.get("val")) for b in el.iter(A_NS + "buSzPts") if b.get("val")]
    if sizes:
        return max(sizes) / 100.0
    return 18.0


def _text_need(tf, width_in: float, fallback_pt: float) -> float:
    """Inches the text frame's paragraphs want, each at its own size.

    Judging every paragraph at the frame's largest size called a big-number
    tile (72pt value, 20pt label, 15pt baseline) three 72pt lines tall and
    reported an overflow that is not there.
    """
    h = 0.0
    for para in tf.paragraphs:
        sizes = [int(r.get("sz")) / 100 for r in para._p.iter(A_NS + "rPr") if r.get("sz")]
        pt = max(sizes) if sizes else fallback_pt
        text = "".join(r.text for r in para.runs)
        h += _needed_height(text, width_in, pt, 1)
        spc = para._p.find(f"{A_NS}pPr/{A_NS}spcBef/{A_NS}spcPts")
        if spc is not None and spc.get("val"):
            h += int(spc.get("val")) / 100 / 72
    return h


def _autofits(tf) -> bool:
    """PowerPoint shrinks text in a `normAutofit` box instead of overflowing it."""
    body = tf._txBody.find(A_NS + "bodyPr")
    return body is not None and body.find(A_NS + "normAutofit") is not None


def _boxes(slide):
    """Every text frame worth judging, with its geometry already in inches."""
    out = []
    for shp in slide.shapes:
        if not shp.has_text_frame or not shp.text_frame.text.strip():
            continue
        w, h = Emu(shp.width).inches, Emu(shp.height).inches
        x, y = Emu(shp.left).inches, Emu(shp.top).inches
        if w < 0.5 or h < 0.3:
            continue
        if y > 6.6 and w < 4.5:
            continue                       # the template's page-number placeholder
        tf = shp.text_frame
        pt = _effective_pt(tf)
        out.append({
            "name": shp.name, "x": x, "y": y, "w": w, "h": h,
            "text": tf.text,
            "pt": pt,
            "paras": len(tf.paragraphs),
            "need": _text_need(tf, w, pt),
            "autofit": _autofits(tf),
        })
    return out


def _deck_issues(deck: Path) -> list[tuple[str, int, str, str]]:
    """(level, slide number, dedup key, message) for everything geometry can see."""
    out = []
    prs = Presentation(str(deck))
    for i, slide in enumerate(prs.slides, start=1):
        boxes = _boxes(slide)
        for b in boxes:
            need = b["need"]
            ratio = need / b["h"]
            # An autofit box shrinks its own text, so a mild overrun is cosmetic
            # (smaller type than the lab's 24pt floor) rather than a clipped
            # slide. Past ~1.8x even autofit gives up.
            if ratio > (AUTOFIT_HARD if b["autofit"] else 1.15):
                out.append(("error", i, f"overflow|{b['name']}|{b['text']}",
                            f"slide {i}: text overflows {b['name']!r} "
                            f"(needs ~{need:.1f}in, box is {b['h']:.1f}in)"))
            elif b["autofit"] and ratio > 1.15:
                out.append(("warn", i, f"shrink|{b['name']}|{b['text']}",
                            f"slide {i}: {b['name']!r} will be auto-shrunk to fit "
                            f"(~{need:.1f}in of text in a {b['h']:.1f}in box) "
                            f"- shorten it or it drops below the 24pt floor"))
            if b["y"] + b["h"] > MARGIN_BOTTOM or b["x"] + b["w"] > MARGIN_RIGHT:
                out.append(("warn", i, f"margin|{b['name']}",
                            f"slide {i}: {b['name']!r} runs past the safe margin"))

        # A title that grew a second line and landed on the subtitle is the
        # classic "you had to look at it" bug. It is arithmetic, not eyesight.
        for a in boxes:
            # autofit keeps the text inside its own box, so it cannot collide
            a_need = min(a["need"], a["h"]) if a["autofit"] else a["need"]
            for c in boxes:
                if c is a or c["y"] <= a["y"]:
                    continue
                overlap = min(a["x"] + a["w"], c["x"] + c["w"]) - max(a["x"], c["x"])
                if overlap < 0.3 * min(a["w"], c["w"]):
                    continue
                if a["y"] + a_need > c["y"] + 0.04:
                    out.append(("error", i, f"collide|{a['name']}|{c['name']}",
                                f"slide {i}: {a['name']!r} wraps down into "
                                f"{c['name']!r} - shorten it or move the box"))

        # leftover fill-me tokens, in shapes and in table cells alike
        for shp in slide.shapes:
            texts = []
            if shp.has_text_frame:
                texts.append(shp.text_frame.text)
            if shp.has_table:
                texts.extend(c.text for r in shp.table.rows for c in r.cells)
            seen = set()
            for t in texts:
                for m in PLACEHOLDER.finditer(t or ""):
                    tok = m.group(0)
                    if tok.lower() in seen:
                        continue
                    seen.add(tok.lower())
                    out.append(("error", i, f"placeholder|{shp.name}|{tok}",
                                f"slide {i}: template placeholder {tok!r} survived "
                                f"into {shp.name!r} - fill it or blank it"))
    return out


@functools.lru_cache(maxsize=4)
def _template_keys(template: str) -> frozenset:
    """What the lab template already trips on its own.

    The template ships real defects -- `Todolist & Suggestion from Prof.`
    overflows its box on every conclusion slide, and slide 2 is wall-to-wall
    `XXX`. Reporting those on every run trains you to ignore the gate, and
    keeps `render_qa.py` refusing to render. Baseline them: a finding is only
    yours if the same shape with the same text is not already broken upstream.
    """
    if not Path(template).exists():
        return frozenset()
    # `placeholder|...` is deliberately excluded: every one of these tokens is
    # in the template by design, so baselining them would suppress exactly the
    # finding they exist to make.
    return frozenset(key for _, _, key, _ in _deck_issues(Path(template))
                     if not key.startswith("placeholder|"))


def check_deck(deck: Path, template: Path | None = TEMPLATE) -> tuple[list[str], list[str]]:
    """Everything a render would have been opened to check, minus the render.

    Overflow, margin breaches, a title colliding with the line under it, and
    template placeholders that were never filled. What is left for your eyes is
    genuinely visual: crop quality, annotation placement, colour.
    """
    inherited = _template_keys(str(template)) if template else frozenset()
    errors, warns = [], []
    for level, _, key, msg in _deck_issues(deck):
        if key in inherited:
            continue
        (errors if level == "error" else warns).append(msg)
    return errors, warns


def _occupied(slide, top: float, bottom: float) -> list[tuple[float, float]]:
    """Vertical extents (inches) of everything in the content region: a
    filled shape or a picture by its frame, a bare text box by its text."""
    spans = []
    for shp in slide.shapes:
        y, h = Emu(shp.top).inches, Emu(shp.height).inches
        w = Emu(shp.width).inches
        if y < top - 0.15 or y > bottom or shp.name == "acm:subtitle":
            continue                       # title band, subtitle, page number, tags
        filled = False
        try:
            filled = shp.fill.type is not None and shp.fill.type != 5   # 5: background
        except Exception:
            pass
        if shp.has_text_frame and not filled and shp.shape_type != 13:
            tf = shp.text_frame
            if not tf.text.strip():
                continue
            h = min(h, _text_need(tf, max(0.5, w - 0.2), _effective_pt(tf)) + 0.1)
        spans.append((max(top, y), min(bottom, y + max(h, 0.05))))
    return sorted(s_ for s_ in spans if s_[1] > s_[0])


def _fill_hint(spec: dict) -> str:
    if spec.get("flow"):
        return "give the steps `detail` lines, or move the flow onto a figure slide"
    if spec.get("matrix") or spec.get("table"):
        return "a short table leaves the slide half empty; lead with the bignum or pair it with the figure it summarises"
    if spec.get("equation") and not spec.get("figure"):
        return "put the figure the equation describes beside it (figure.src takes two)"
    if any(spec.get(k) for k in EXHIBIT_KEYS):
        return "enlarge the exhibit or use a layout that gives it the space"
    return "add an exhibit - cards, a flow, a figure, big numbers"


def check_fill(flat: list[dict], deck: Path) -> tuple[list[str], list[str]]:
    """The emptiness a rule-abiding outline still ships.

    Everything else in this file is about too much: too many words, too many
    callouts, too many tables. Nothing caught three bullets under a title and
    a blank lower half, or a strip of boxes floating in the middle of five
    empty inches -- and those passed every check. The reference decks give
    their figures 35-72% of the slide. This measures how much of the content
    region's height carries anything and the tallest empty band across it.
    """
    warns = []
    prs = Presentation(str(deck))
    for i, (slide, spec) in enumerate(zip(prs.slides, flat), start=1):
        if spec.get("role") in EXEMPT or spec.get("role") == "cover" or spec.get("table"):
            continue
        top = 1.78 if spec.get("subtitle") else 1.72
        if spec.get("layout") in ("assertion-evidence", "ae"):
            top = 2.10
        bottom = 6.80
        spans = _occupied(slide, top, bottom)
        covered, edge, gap, gap_at = 0.0, top, 0.0, (top, bottom)
        for a, b in spans + [(bottom, bottom)]:
            if a > edge:
                if a - edge > gap:
                    gap, gap_at = a - edge, (edge, a)
                covered += 0.0
            if b > edge:
                covered += b - max(a, edge)
                edge = b
        share = covered / (bottom - top)
        if share < FILL_MIN or gap > GAP_MAX:
            hint = _fill_hint(spec)
            warns.append(f"slide {i} ({spec.get('role')}): {share:.0%} of the content area "
                         f"is used; {gap:.1f}in stands empty between y={gap_at[0]:.1f} and "
                         f"{gap_at[1]:.1f}in - {hint}")
    return [], warns


CUSTOM_MIN_SHAPES = 2        # beyond title, body and page number


def check_custom(flat: list[dict], deck: Path) -> tuple[list[str], list[str]]:
    """`"custom": true` promises an exhibit drawn by the user's own script.
    Count what is actually on the slide, so the promise is checked against
    the deck rather than taken on faith (feedback F8)."""
    errors, warns = [], []
    prs = Presentation(str(deck))
    for i, (slide, spec) in enumerate(zip(prs.slides, flat), start=1):
        if not spec.get("custom"):
            continue
        extra = 0
        for shp in slide.shapes:
            if shp.has_text_frame and (shp.is_placeholder or shp.name == "acm:subtitle"):
                continue
            if shp.has_text_frame and shp.text_frame.text.strip() in ("", "\u2039#\u203a"):
                continue
            extra += 1
        if extra < CUSTOM_MIN_SHAPES:
            errors.append(f"slide {i}: marked `custom` but carries {extra} drawn shape(s) "
                          f"- run your drawing script before QA, or drop the flag")
    return errors, warns


def check_diagram_editability(flat: list[dict], deck: Path) -> tuple[list[str], list[str]]:
    """Prove diagram semantics landed as named native objects, not a bitmap."""
    errors, warns = [], []
    prs = Presentation(str(deck))
    for i, (slide, spec) in enumerate(zip(prs.slides, flat), start=1):
        diagram = spec.get("diagram")
        if not isinstance(diagram, dict):
            continue
        by_name = {shape.name: shape for shape in slide.shapes}
        names = set(by_name)
        for node in diagram.get("nodes") or []:
            expected = f"diagram:node:{node.get('id')}"
            if expected not in names:
                errors.append(f"slide {i}: native diagram node {expected!r} is missing")
            elif not by_name[expected].has_text_frame:
                errors.append(f"slide {i}: diagram node {expected!r} is not an "
                              f"editable text-bearing PowerPoint shape")
        for index, edge in enumerate(diagram.get("edges") or [], start=1):
            edge_id = edge.get("id") or f"e{index}"
            expected = f"diagram:edge:{edge_id}"
            if expected not in names:
                errors.append(f"slide {i}: native diagram edge {expected!r} is missing")
            elif edge.get("attached", True):
                element = by_name[expected]._element
                if not element.xpath(".//a:stCxn") or not element.xpath(".//a:endCxn"):
                    errors.append(f"slide {i}: diagram edge {expected!r} is not "
                                  f"attached to both endpoint nodes")
        for index, group in enumerate(diagram.get("groups") or [], start=1):
            group_id = group.get("id") or index
            expected = f"diagram:group:{group_id}"
            if expected not in names:
                errors.append(f"slide {i}: native diagram group {expected!r} is missing")
        if not any(name.startswith("diagram:node:") for name in names):
            errors.append(f"slide {i}: diagram has no editable PowerPoint node shapes")
    return errors, warns


def ghost_deck(flat: list[dict]) -> str:
    """Minto's horizontal logic: the claims alone must tell the whole story."""
    lines = []
    for i, spec in enumerate(flat, start=1):
        if spec.get("role") in EXEMPT:
            continue
        claim = spec.get("subtitle") or ""
        if not claim and spec.get("callout"):
            c = spec["callout"]
            claim = c["text"] if isinstance(c, dict) else c
        # the title is the primary claim carrier; fall back to it, and mark a
        # bare section label so a topic list is visible at a glance
        if not claim:
            title = (spec.get("title") or "").strip()
            claim = f"[title] {title}" if title and not NON_CLAIM.match(title) \
                else f"[label] {title or '(none)'}"
        lines.append(f"  {i:2}. {claim}")
    return "\n".join(lines)


def inventory(flat: list[dict]) -> str:
    rows = []
    for i, spec in enumerate(flat, start=1):
        role = spec.get("role", "?")
        if role in EXEMPT:
            continue
        ex = [k for k in EXHIBIT_KEYS if spec.get(k)] or ["-"]
        rows.append(f"  {i:2}. {role:17} {spec.get('layout') or 'auto':18} "
                    f"{'+'.join(ex):16} {words(bullet_text(spec)):3}w "
                    f"{len((spec.get('notes') or '').strip()):4}c notes")
    return "\n".join(rows)


REVIEW_RUBRIC = """
Now score the deck yourself on the three dimensions PPTEval uses, because none
of the checks above can see any of them:

  Content    Is each claim actually supported by the exhibit on its own slide?
             Any number, baseline or dataset stated that the source does not?
             Does a results slide show an actual observation, not just status?
             Can borrowed panels and replotted values be traced to their source?
  Design     Render the deck and look. Crowding, unreadable axis labels, a
             figure that needed a tighter crop, an annotation covering the
             thing it points at. Does each matrix need row/column lookup,
             or would a diagram, plot or paired result image explain more?
             Check figure labels at slide size, not only on a contact sheet.
  Coherence  Read the claim sequence above as a single paragraph. Does each
             claim follow from the one before? Where does it jump?

Name the three weakest slides and what you would change. Coherence is where
the ablation in that work showed the largest gap, and it is the one thing a
word count can never catch.
"""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline")
    ap.add_argument("deck", nargs="?")
    ap.add_argument("--review", action="store_true",
                    help="also print the claim sequence, a slide inventory, and "
                         "the rubric to judge content, design and coherence")
    a = ap.parse_args()

    outline = json.loads(Path(a.outline).read_text(encoding="utf-8"))
    errors, warns = check_outline(flatten(outline))
    if a.deck:
        e2, w2 = check_deck(Path(a.deck))
        errors += e2
        warns += w2
        e3, w3 = check_diagram_editability(flatten(outline), Path(a.deck))
        errors += e3
        warns += w3
        e4, w4 = check_custom(flatten(outline), Path(a.deck))
        errors += e4
        warns += w4
        e5, w5 = check_fill(flatten(outline), Path(a.deck))
        errors += e5
        warns += w5

    flat = flatten(outline)
    if a.review:
        print("CLAIM SEQUENCE (read this as one paragraph)")
        print(ghost_deck(flat))
        print("\nSLIDE INVENTORY")
        print(inventory(flat))
        print(REVIEW_RUBRIC)

    for w in warns:
        print(f"WARN  {w}")
    for e in errors:
        print(f"ERROR {e}")
    print(f"\n{len(errors)} error(s), {len(warns)} warning(s)")
    if errors:
        print("Fix the errors, rebuild, re-run. Warnings are a judgement call.")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
