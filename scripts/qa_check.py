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
EXHIBIT_KEYS = ("figure", "video", "matrix", "equation", "stage", "table")
ROLE_SCHEMA = {
    "paper_method":     {"exhibit": "error"},
    "paper_results":    {"exhibit": "error"},
    "paper_related":    {"exhibit": "warn", "callout": "warn"},
    "paper_intro":      {"exhibit": "warn"},
    "paper_conclusion": {"bullets": "warn"},
    "project_results":  {"exhibit": "error"},
    "research_results": {"exhibit": "error"},
}

NON_CLAIM = re.compile(
    r"^(results?|method|methods|introduction|conclusion|related work|"
    r"experiments?|ablation|overview|background|motivation)$", re.I)


def words(text: str) -> int:
    """CJK counts per character; latin per whitespace token."""
    cjk = len(re.findall(r"[\u4e00-\u9fff\u3040-\u30ff]", text))
    latin = len(re.findall(r"[A-Za-z0-9][A-Za-z0-9\-/%.+]*", text))
    return cjk + latin


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
        elif claim and len(claim.split()) < CLAIM_MIN_WORDS:
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

        if not any(spec.get(k) for k in ("figure", "video", "matrix", "equation",
                                         "table")):
            text_only += 1

        # 4. what this functional slide type must carry
        schema = ROLE_SCHEMA.get(role, {})
        for key, severity in schema.items():
            if key == "exhibit":
                ok = any(spec.get(k) for k in EXHIBIT_KEYS)
                msg = (f"{tag}: a {role} slide with no exhibit - one figure, "
                       f"matrix or equation per slide")
            else:
                ok = bool(spec.get(key))
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
    return errors, warns


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
        out.append({
            "name": shp.name, "x": x, "y": y, "w": w, "h": h,
            "text": tf.text,
            "pt": _effective_pt(tf),
            "paras": len(tf.paragraphs),
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
            need = _needed_height(b["text"], b["w"], b["pt"], b["paras"])
            b["need"] = need
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
             Any number, baseline or dataset stated that the paper does not?
  Design     Render the deck and look. Crowding, unreadable axis labels, a
             figure that needed a tighter crop, an annotation covering the
             thing it points at.
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
