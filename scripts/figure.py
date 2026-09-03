#!/usr/bin/env python3
"""Pull presentation-resolution figures out of a paper PDF.

Two entry points, both re-render the ORIGINAL PDF at high dpi and crop --
never reuse a bitmap somebody else already downsampled.

    # list what MinerU found
    python scripts/figure.py list --mineru out/paper/auto

    # crop MinerU block #3 at 300 dpi
    python scripts/figure.py crop --pdf paper.pdf --mineru out/paper/auto \
        --index 3 -o figs/fig2.png

    # no MinerU: crop by fraction of the page (x0,y0,x1,y1 in 0..1)
    python scripts/figure.py crop --pdf paper.pdf --page 3 \
        --box 0.08,0.05,0.95,0.42 -o figs/teaser.png

MinerU bbox is [x0,y0,x1,y1] normalised to 0-1000 and page_idx is 0-based,
so both forms reduce to the same fractional crop.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

DEFAULT_DPI = 300
PAD = 0.006          # MinerU boxes clip captions/axis labels; bleed a little
MIN_PX = 900         # below this a figure will look soft on a projector


# --------------------------------------------------------------- MinerU

def _content_list(mineru: Path) -> Path:
    """MinerU writes <stem>_content_list.json (3.x may add _v2)."""
    mineru = Path(mineru)
    if mineru.is_file():
        return mineru
    cands = sorted(mineru.glob("*content_list*.json"))
    if not cands:
        raise SystemExit(f"no *_content_list*.json under {mineru}")
    # prefer v2 when both exist
    return sorted(cands, key=lambda p: ("_v2" not in p.name, p.name))[0]


def blocks(mineru: Path) -> list[dict]:
    """Image/table blocks in reading order, with caption already paired."""
    data = json.loads(_content_list(mineru).read_text(encoding="utf-8"))
    out = []
    for b in data:
        if b.get("type") not in ("image", "table", "chart"):
            continue
        cap = b.get("img_caption") or b.get("table_caption") or []
        foot = b.get("img_footnote") or b.get("table_footnote") or []
        out.append({
            "type": b["type"],
            "page": b.get("page_idx", 0),
            "bbox": b.get("bbox"),
            "caption": " ".join(cap).strip(),
            "footnote": " ".join(foot).strip(),
            "path": b.get("img_path") or b.get("table_img_path"),
        })
    return out


# --------------------------------------------------------------- cropping

def render_page(pdf: Path, page_idx: int, dpi: int, workdir: Path) -> Path:
    """pdftoppm renders 1-based pages; MinerU counts from 0."""
    prefix = workdir / "page"
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi),
         "-f", str(page_idx + 1), "-l", str(page_idx + 1),
         str(pdf), str(prefix)],
        check=True, capture_output=True,
    )
    hits = sorted(workdir.glob("page*.png"))
    if not hits:
        raise SystemExit(f"pdftoppm produced nothing for page {page_idx + 1}")
    return hits[0]


def crop(pdf: Path, page_idx: int, box: tuple[float, float, float, float],
         out: Path, dpi: int = DEFAULT_DPI, pad: float = PAD) -> Path:
    """box is (x0,y0,x1,y1) as fractions of the page."""
    from PIL import Image

    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        png = render_page(pdf, page_idx, dpi, Path(td))
        im = Image.open(png)
        W, H = im.size
        x0, y0, x1, y1 = box
        rect = (
            max(0, int((x0 - pad) * W)), max(0, int((y0 - pad) * H)),
            min(W, int((x1 + pad) * W)), min(H, int((y1 + pad) * H)),
        )
        if rect[2] - rect[0] < 10 or rect[3] - rect[1] < 10:
            raise SystemExit(f"crop box degenerate: {box}")
        im.crop(rect).save(out)

    w, h = Image.open(out).size
    if max(w, h) < MIN_PX:
        print(f"  warning: {out.name} is {w}x{h}px - re-run with a higher --dpi",
              file=sys.stderr)
    print(f"{out}  {w}x{h}px  page {page_idx + 1} @ {dpi}dpi")
    return out


# --------------------------------------------------------------- cli

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="show MinerU figure/table blocks")
    p_list.add_argument("--mineru", required=True)

    p_crop = sub.add_parser("crop", help="crop one figure at presentation dpi")
    p_crop.add_argument("--pdf", required=True)
    p_crop.add_argument("--mineru")
    p_crop.add_argument("--index", type=int, help="index from `list`")
    p_crop.add_argument("--page", type=int, help="1-based page, when no MinerU")
    p_crop.add_argument("--box", help="x0,y0,x1,y1 as page fractions")
    p_crop.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    p_crop.add_argument("-o", "--output", required=True)

    args = ap.parse_args()

    if args.cmd == "list":
        for i, b in enumerate(blocks(Path(args.mineru))):
            cap = (b["caption"] or "(no caption)")[:70]
            print(f"[{i:2}] {b['type']:5} p{b['page'] + 1:<3} {cap}")
        return

    if not shutil.which("pdftoppm"):
        raise SystemExit("pdftoppm not found (install poppler-utils)")

    pdf = Path(args.pdf)
    if args.mineru is not None and args.index is not None:
        b = blocks(Path(args.mineru))[args.index]
        if not b["bbox"]:
            raise SystemExit(f"block {args.index} carries no bbox")
        x0, y0, x1, y1 = (v / 1000.0 for v in b["bbox"])
        page = b["page"]
        if b["caption"]:
            print(f"  caption: {b['caption']}")
    elif args.page and args.box:
        page = args.page - 1
        x0, y0, x1, y1 = (float(v) for v in args.box.split(","))
    else:
        raise SystemExit("need either --mineru/--index or --page/--box")

    crop(pdf, page, (x0, y0, x1, y1), Path(args.output), args.dpi)


if __name__ == "__main__":
    main()
