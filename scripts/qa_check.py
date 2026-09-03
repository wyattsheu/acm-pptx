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
TEXT_ONLY_MAX = 0.15        # WARN only: measured 8% (ELITE) and 12% (TADSR)
CLAIM_MIN_WORDS = 4

# roles that are structural, not argument-bearing
EXEMPT = {"cover", "summary", "paper_list", "project_summary", "research_summary"}

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

    for i, spec in enumerate(flat, start=1):
        role = spec.get("role", "?")
        tag = f"slide {i} ({role})"
        if role in EXEMPT:
            continue

        # 1. the claim must exist somewhere -- subtitle or callout, not the title
        claim = spec.get("subtitle") or ""
        if not claim and isinstance(spec.get("callout"), dict):
            claim = spec["callout"].get("text", "")
        elif not claim and isinstance(spec.get("callout"), str):
            claim = spec["callout"]
        if not claim.strip():
            errors.append(f"{tag}: no claim line - add `subtitle` or `callout`. "
                          f"The template's titles are labels; the argument lives "
                          f"in the red line and the bottom box.")
        elif len(claim.split()) < CLAIM_MIN_WORDS:
            warns.append(f"{tag}: claim {claim!r} reads as a label, not a statement")

        if NON_CLAIM.match((spec.get("title") or "").strip()) and not claim:
            errors.append(f"{tag}: generic title {spec.get('title')!r} with no claim line")

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

        # 4. speaker notes
        notes = (spec.get("notes") or "").strip()
        if not notes:
            warns.append(f"{tag}: no speaker notes")
        elif len(notes) < NOTES_MIN:
            warns.append(f"{tag}: notes are {len(notes)} chars - thin")

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


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline")
    ap.add_argument("deck", nargs="?")
    a = ap.parse_args()

    outline = json.loads(Path(a.outline).read_text(encoding="utf-8"))
    errors, warns = check_outline(flatten(outline))
    if a.deck:
        e2, w2 = check_deck(Path(a.deck))
        errors += e2
        warns += w2

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
