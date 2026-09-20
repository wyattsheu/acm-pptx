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
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt

BODY_WORD_CAP = 40          # lab-rules.md: "~40 words of body text per slide, hard"
NOTES_MIN = 80              # WARN only: TADSR leaves results-slide notes empty
TEXT_ONLY_MAX = 0.25          # WARN only: SliderEdit sits at 24% (4/17)
CALLOUT_MAX = 0.15            # WARN only: both reference decks use zero
CLAIM_MIN_WORDS = 4

# roles that are structural, not argument-bearing
EXEMPT = {"cover", "summary", "paper_list", "project_summary", "research_summary"}

# What each functional slide type must carry. PPTAgent's Stage I calls this a
# content schema; here it is the machine-checkable half of slide-patterns.md.
EXHIBIT_KEYS = ("figure", "matrix", "equation", "stage", "table")
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
        if not any(spec.get(k) for k in ("figure", "matrix", "equation", "table")):
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
        if not notes:
            warns.append(f"{tag}: no speaker notes")
        elif len(notes) < NOTES_MIN:
            warns.append(f"{tag}: notes are {len(notes)} chars - thin")

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


def check_deck(deck: Path) -> tuple[list[str], list[str]]:
    """Cheap overflow estimate. The render-and-look pass still matters."""
    errors, warns = [], []
    prs = Presentation(str(deck))
    for i, slide in enumerate(prs.slides, start=1):
        for shp in slide.shapes:
            if not shp.has_text_frame or not shp.text_frame.text.strip():
                continue
            w_in = Emu(shp.width).inches
            h_in = Emu(shp.height).inches
            if w_in < 0.5 or h_in < 0.3:
                continue
            if Emu(shp.top).inches > 6.6 and w_in < 4.5:
                continue           # the template's own page-number placeholder
            sizes = [r.font.size.pt for p in shp.text_frame.paragraphs
                     for r in p.runs if r.font.size]
            pt = max(sizes) if sizes else 20.0
            # latin glyphs run ~0.52em wide, CJK 1.0em; 1.22 line spacing
            txt = shp.text_frame.text
            cjk = len(re.findall(r"[\u4e00-\u9fff\u3040-\u30ff]", txt))
            ems = (len(txt) - cjk) * 0.52 + cjk * 1.0
            per_line = max(1.0, (w_in * 72) / pt)
            lines = max(len(shp.text_frame.paragraphs), ems / per_line)
            need = lines * pt * 1.22 / 72
            if need > h_in * 1.15:
                errors.append(f"slide {i}: text likely overflows {shp.name!r} "
                              f"(needs ~{need:.1f}in, box is {h_in:.1f}in)")
            if Emu(shp.top).inches + h_in > 7.2 or Emu(shp.left).inches + w_in > 13.1:
                warns.append(f"slide {i}: {shp.name!r} runs past the safe margin")
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
