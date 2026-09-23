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
    --changed       only slides whose XML moved since the last run
    --contact N     tile N slides per image instead of one image per slide
    --dpi N         default 110; below ~90 small captions start to break up
    --keep          add to the scratch dir instead of clearing it first

`--contact` is the one that changes the arithmetic. Every image costs the same
~1600 tokens whatever it contains, so four slides tiled into one sheet cost a
quarter of four separate slides. A sheet is deck-level triage -- balance, a
figure that swamps its slide, a breadcrumb on the wrong section -- and captions
are not readable on it past two columns. Everything a caption would have told
you (leftover `XXX`, a title colliding with the subtitle, an overflowing box)
`qa_check.py` already finds without an image at all.

    --all --contact 6      a 25-slide deck for ~7k tokens instead of ~40k
    --changed              after a fix: only what you actually touched
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
LABEL_H = 18                  # the "slide 7" strip above each cell in a sheet


def tokens_per_page(dpi: int) -> int:
    w, h = SLIDE_W_IN * dpi, SLIDE_H_IN * dpi
    scale = min(1.0, LONG_EDGE_CAP / max(w, h))
    return int(w * scale * h * scale / 750)


def tokens_for(img: Path) -> int:
    """What this particular image costs once the vision pipeline downscales it."""
    from PIL import Image

    with Image.open(img) as im:
        w, h = im.size
    scale = min(1.0, LONG_EDGE_CAP / max(w, h))
    return int(w * scale * h * scale / 750)


def slide_digests(deck: Path) -> dict[int, str]:
    """A hash per slide, so a rebuild can be diffed instead of re-read."""
    import hashlib
    import zipfile

    from pptx import Presentation

    prs = Presentation(str(deck))
    out = {}
    with zipfile.ZipFile(deck) as zf:
        for i, slide in enumerate(prs.slides, start=1):
            part = slide.part.partname[1:]          # strip the leading "/"
            blob = zf.read(part)
            rels = f"{Path(part).parent}/_rels/{Path(part).name}.rels"
            try:
                blob += zf.read(rels)
            except KeyError:
                pass
            out[i] = hashlib.sha1(blob).hexdigest()
    return out


def changed_pages(deck: Path, scratch: Path) -> tuple[set[int], bool]:
    """Slides whose XML moved since the last run in this scratch dir."""
    state = scratch / "digests.json"
    now = slide_digests(deck)
    if not state.exists():
        return set(now), True                        # first run: everything is new
    was = {int(k): v for k, v in json.loads(state.read_text()).items()}
    return {i for i, h in now.items() if was.get(i) != h}, False


def write_digests(deck: Path, scratch: Path) -> None:
    (scratch / "digests.json").write_text(json.dumps(slide_digests(deck)))


def contact_sheets(imgs: list[Path], pages: list[int], per_sheet: int,
                   scratch: Path) -> list[Path]:
    """Tile rendered slides into sheets. One image, one price, many slides."""
    from PIL import Image, ImageDraw

    long_edge = LONG_EDGE_CAP
    out = []
    for n, start in enumerate(range(0, len(imgs), per_sheet), start=1):
        chunk = list(zip(pages[start:start + per_sheet], imgs[start:start + per_sheet]))
        # never leave half a sheet blank: a short last chunk narrows the grid
        cols = min(2 if per_sheet <= 6 else 3, len(chunk))
        rows = -(-len(chunk) // cols)
        cell_w = long_edge // cols
        with Image.open(chunk[0][1]) as probe:
            cell_h = int(cell_w * probe.height / probe.width)
        sheet = Image.new("RGB", (cell_w * cols, (cell_h + LABEL_H) * rows),
                          (0xFF, 0xFF, 0xFF))
        draw = ImageDraw.Draw(sheet)
        for k, (page, img) in enumerate(chunk):
            cx, cy = (k % cols) * cell_w, (k // cols) * (cell_h + LABEL_H)
            with Image.open(img) as im:
                sheet.paste(im.convert("RGB").resize((cell_w, cell_h)), (cx, cy + LABEL_H))
            draw.text((cx + 6, cy + 4), f"slide {page}", fill=(0x20, 0x20, 0x20))
            draw.rectangle([cx, cy + LABEL_H, cx + cell_w - 1, cy + LABEL_H + cell_h - 1],
                           outline=(0xCC, 0xCC, 0xCC))
        dest = scratch / f"contact-{n:02d}.png"
        sheet.save(dest)
        out.append(dest)
    for img in imgs:
        img.unlink(missing_ok=True)                  # the sheets replace them
    return out


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
    ap.add_argument("--changed", action="store_true",
                    help="only slides whose XML moved since the last run")
    ap.add_argument("--contact", type=int, default=1, metavar="N",
                    help="tile N slides per image (2-6); one image costs the "
                         "same whatever it holds")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--keep", action="store_true",
                    help="add to the scratch dir instead of clearing it first")
    a = ap.parse_args()

    outline, deck = Path(a.outline), Path(a.deck)
    flags, n_err, n_warn = flagged_pages(outline, deck)

    from pptx import Presentation
    total = len(Presentation(str(deck)).slides._sldIdLst)

    # Scratch lives outside the user's working directory, always. The deck is
    # the only artifact they asked for; renders, PDFs and stray PNGs are ours.
    scratch = Path(tempfile.gettempdir()) / "acm-qa" / deck.stem
    scratch.mkdir(parents=True, exist_ok=True)
    moved, first_run = changed_pages(deck, scratch)

    if a.pages:
        pages = sorted(parse_pages(a.pages))
        why = "explicit"
    elif a.changed:
        pages = sorted(moved | ({1} if first_run else set()))
        why = ("first run, everything is new" if first_run
               else f"{len(pages)} changed since the last run")
        if not pages:
            print(f"qa_check: {n_err} error(s), {n_warn} warning(s)")
            print("nothing changed since the last render - no tokens spent")
            write_digests(deck, scratch)
            sys.exit(1 if n_err else 0)
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

    if not a.keep:
        for stale in list(scratch.glob("slide-*.png")) + list(scratch.glob("contact-*.png")):
            stale.unlink(missing_ok=True)

    imgs = render(deck, pages, a.dpi, scratch)
    per_sheet = max(1, min(a.contact, 6))
    if per_sheet > 1 and imgs:
        imgs = contact_sheets(imgs, pages, per_sheet, scratch)
    write_digests(deck, scratch)

    per = tokens_per_page(a.dpi)
    print(f"qa_check: {n_err} error(s), {n_warn} warning(s)")
    print(f"rendered {len(pages)}/{total} slides ({why}) at {a.dpi} dpi -> {scratch}")
    if per_sheet > 1:
        print(f"tiled {per_sheet} per sheet - read these for balance and layout, "
              f"not for caption text")
        for img in imgs:
            print(f"       {img}")
    else:
        for p, img in zip(pages, imgs):
            mark = "FLAG" if p in flags else "    "
            print(f"  {mark} slide {p:>2}  {img}")
    spent, full = sum(tokens_for(i) for i in imgs), total * per
    print(f"\n~{spent} tokens to view these; ~{max(0, full - spent)} saved "
          f"against rendering all {total} one by one")
    if any((s_.get("video") if isinstance(s_, dict) else None)
           for s_ in json.loads(outline.read_text(encoding="utf-8")).get("slides", [])):
        print("note: a video renders as its poster frame here - playback itself "
              "can only be checked by opening the deck in PowerPoint")
    print("nothing was written to the working directory")
    if n_err:
        print("qa_check has errors - fix the outline before spending tokens on images")
    sys.exit(1 if n_err else 0)


if __name__ == "__main__":
    main()
