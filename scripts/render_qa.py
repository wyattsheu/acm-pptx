#!/usr/bin/env python3
"""Render only the slides worth looking at, into a scratch dir, not the user's.

    python scripts/render_qa.py outline.json 2026-09-18-GM.pptx

Two things this exists to stop.

One: the working directory filling up with `slide-01.jpg` ... `slide-25.jpg`
and a `.pdf` nobody asked for. The deck is the deliverable; everything here is
scaffolding, so it goes to a temp dir and is deleted when you are done. Pass
`--keep` to stack runs up instead of replacing the last one.

Two: the token bill. A slide reaches the model downscaled to 1568px on its long
edge -- ~1600 tokens each, no matter whether you rendered at 110 dpi or 300.
A 25-slide talk eyeballed in full is ~40k, and again on every rebuild.
qa_check.py already finds overflow and margin breaches geometrically, for free,
so render what it flagged plus a small sample and stop there.

    --pages 3,7-9   render exactly these, ignore the flags
    --sample N      extra evenly-spaced slides beyond the flagged ones (default 2)
    --all           render everything (use once, on the final pass)
    --dpi N         default 110; below ~90 small captions start to break up
    --keep          add to the scratch dir instead of clearing it first
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import qa_check  # noqa: E402

SLIDE_W_IN, SLIDE_H_IN = 13.333, 7.5
LONG_EDGE_CAP = 1568          # the vision downscale; dpi above this buys nothing


def tokens_per_page(dpi: int) -> int:
    w, h = SLIDE_W_IN * dpi, SLIDE_H_IN * dpi
    scale = min(1.0, LONG_EDGE_CAP / max(w, h))
    return int(w * scale * h * scale / 750)


def parse_pages(spec: str) -> set[int]:
    out: set[int] = set()
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(chunk))
    return out


def flagged_pages(outline: Path, deck: Path) -> tuple[set[int], int, int]:
    """Slide numbers qa_check complained about, plus its error/warn counts."""
    data = json.loads(outline.read_text(encoding="utf-8"))
    errors, warns = qa_check.check_outline(qa_check.flatten(data))
    e2, w2 = qa_check.check_deck(deck)
    errors, warns = errors + e2, warns + w2

    pages = set()
    for msg in errors + warns:
        m = re.match(r"slide (\d+)", msg)
        if m:
            pages.add(int(m.group(1)))
    return pages, len(errors), len(warns)


def render(deck: Path, pages: list[int], dpi: int, scratch: Path) -> list[Path]:
    """PPTX -> PDF -> PNG, entirely inside scratch. Nothing touches deck.parent."""
    if not shutil.which("pdftoppm"):
        raise SystemExit("pdftoppm not found - install Poppler")

    soffice = Path(__file__).resolve().parent / "office" / "soffice.py"
    subprocess.run([sys.executable, str(soffice), "--headless", "--convert-to",
                    "pdf", "--outdir", str(scratch), str(deck)],
                   check=True, capture_output=True)
    pdf = scratch / (deck.stem + ".pdf")
    if not pdf.exists():
        raise SystemExit(f"conversion produced no {pdf}")

    out = []
    for p in pages:
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi),
                        "-f", str(p), "-l", str(p),
                        str(pdf), str(scratch / "slide")],
                       check=True, capture_output=True)
        hits = sorted(scratch.glob(f"slide-*{p}.png"))
        out.extend(h for h in hits if h not in out)
    pdf.unlink(missing_ok=True)          # nobody asked for a PDF
    return sorted(set(out))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline")
    ap.add_argument("deck")
    ap.add_argument("--pages", help="explicit list, e.g. 3,7-9")
    ap.add_argument("--sample", type=int, default=2)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--keep", action="store_true",
                    help="add to the scratch dir instead of clearing it first")
    a = ap.parse_args()

    outline, deck = Path(a.outline), Path(a.deck)
    flags, n_err, n_warn = flagged_pages(outline, deck)

    from pptx import Presentation
    total = len(Presentation(str(deck)).slides._sldIdLst)

    if a.pages:
        pages = sorted(parse_pages(a.pages))
        why = "explicit"
    elif a.all:
        pages = list(range(1, total + 1))
        why = "all"
    else:
        picked = set(flags)
        step = max(1, total // (a.sample + 1))
        picked.update(range(step, total + 1, step) if a.sample else [])
        picked.add(1)                    # the cover carries the XXX / 20XX leftovers
        pages = sorted(p for p in picked if 1 <= p <= total)
        why = f"{len(flags)} flagged + {len(pages) - len(flags)} sampled"

    # Scratch lives outside the user's working directory, always. The deck is
    # the only artifact they asked for; renders, PDFs and stray PNGs are ours.
    scratch = Path(tempfile.gettempdir()) / "acm-qa" / deck.stem
    if scratch.exists() and not a.keep:
        shutil.rmtree(scratch, ignore_errors=True)
    scratch.mkdir(parents=True, exist_ok=True)

    imgs = render(deck, pages, a.dpi, scratch)

    per = tokens_per_page(a.dpi)
    print(f"qa_check: {n_err} error(s), {n_warn} warning(s)")
    print(f"rendered {len(pages)}/{total} slides ({why}) at {a.dpi} dpi -> {scratch}")
    for p, img in zip(pages, imgs):
        mark = "FLAG" if p in flags else "    "
        print(f"  {mark} slide {p:>2}  {img}")
    print(f"\n~{len(imgs) * per} tokens to view these; "
          f"~{(total - len(pages)) * per} saved by skipping the other "
          f"{total - len(pages)}")
    print("nothing was written to the working directory")
    if n_err:
        print("qa_check has errors - fix the outline before spending tokens on images")
    sys.exit(1 if n_err else 0)


if __name__ == "__main__":
    main()
