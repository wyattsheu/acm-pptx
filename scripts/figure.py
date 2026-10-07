#!/usr/bin/env python3
"""Pull presentation-resolution figures out of a paper PDF.

Two entry points, both re-render the ORIGINAL PDF at high dpi and crop --
never reuse a bitmap somebody else already downsampled.

    # list what MinerU found
    python scripts/figure.py list --mineru out/paper/auto

    # crop MinerU block #3 at 300 dpi
    python scripts/figure.py crop --pdf paper.pdf --mineru out/paper/auto \
        --index 3 -o figs/fig2.png

    # no MinerU: see what is on the page first -- a grid of page fractions
    # (preview) and the candidate figure blocks it finds (detect)
    python3 scripts/figure.py preview --pdf paper.pdf --page 3 -o page3.png
    python3 scripts/figure.py detect  --pdf paper.pdf --page 3

    # then crop by fraction of the page (x0,y0,x1,y1 in 0..1)
    python3 scripts/figure.py crop --pdf paper.pdf --page 3 \
        --box 0.08,0.05,0.95,0.42 -o figs/teaser.png --trim

MinerU bbox is [x0,y0,x1,y1] normalised to 0-1000 and page_idx is 0-based,
so both forms reduce to the same fractional crop. `--box` is CORNERS, which
is also what an annotation's `xyxy` takes in the outline.

Guessing a box from whole-page fractions cost 11-19 crops per paper in the
decks this skill was first used for (feedback C1/E1/F12). `preview` and
`detect` exist so the first guess is usually the last: read the box off the
grid, or take the one `detect` printed. `--trim` removes the blank border
(and, with `--drop-caption`, a caption line left under the figure), and the
arXiv stamp down the left margin is kept out by default.
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


ARXIV_MARGIN = 0.055   # the vertical arXiv stamp lives inside this left fraction
INK = 235              # grey level below which a pixel counts as ink


def _ink_mask(im):
    """Boolean ink map (numpy) of a page or crop."""
    import numpy as np
    g = np.asarray(im.convert("L"))
    return g < INK


def _runs(profile, min_gap: int):
    """Spans of consecutive True in a 1-D profile, merging gaps under min_gap."""
    spans = []
    start = None
    for i, v in enumerate(list(profile) + [False]):
        if v and start is None:
            start = i
        elif not v and start is not None:
            spans.append([start, i])
            start = None
    merged = []
    for a, b in spans:
        if merged and a - merged[-1][1] < min_gap:
            merged[-1][1] = b
        else:
            merged.append([a, b])
    return merged


def strip_arxiv_stamp(mask, keep_left: float = ARXIV_MARGIN):
    """arXiv prints `arXiv:2503.17973v1 [cs.CV] 23 Mar 2025` rotated down the
    left margin; a crop that reaches the page edge brings it along (feedback
    C2). It is a thin tall strip of ink with nothing to its right for a long
    stretch, so blank that column band when it looks like that."""
    import numpy as np
    H, W = mask.shape
    band = mask[:, : int(W * keep_left)]
    rows_with_ink = band.any(axis=1)
    cols_with_ink = band.any(axis=0)
    # tall (ink on more than a quarter of the rows) and narrow (ink in under
    # half of the margin's columns): nothing but the stamp looks like that
    if rows_with_ink.mean() > 0.25 and cols_with_ink.mean() < 0.5:
        mask = mask.copy()
        mask[:, : int(W * keep_left)] = False
    return mask


def trim_box(im, *, drop_caption: bool = False):
    """Pixel box of the content inside `im`, blank border removed; with
    drop_caption, a short text band at the bottom separated by a gap is cut
    (feedback C3)."""
    import numpy as np
    mask = _ink_mask(im)
    H, W = mask.shape
    rows = mask.any(axis=1)
    cols = mask.any(axis=0)
    if not rows.any():
        return (0, 0, W, H)
    y0, y1 = int(np.argmax(rows)), H - int(np.argmax(rows[::-1]))
    x0, x1 = int(np.argmax(cols)), W - int(np.argmax(cols[::-1]))
    if drop_caption:
        bands = _runs(rows[y0:y1], min_gap=max(3, H // 150))
        if len(bands) >= 2:
            last = bands[-1]
            text_h = (last[1] - last[0])
            gap = last[0] - bands[-2][1]
            # a caption is a short band (under ~2.5 text lines) after a gap
            if text_h < 0.09 * H and gap > 0.012 * H:
                y1 = y0 + bands[-2][1]
    pad = max(2, int(0.006 * max(W, H)))
    return (max(0, x0 - pad), max(0, y0 - pad), min(W, x1 + pad), min(H, y1 + pad))


def detect_blocks(im, *, strip_stamp: bool = True) -> list[tuple[float, float, float, float]]:
    """Candidate figure blocks on a page image, as (x0, y0, x1, y1) fractions.

    Text is dense ink in regular lines; a figure is a band that is either
    tall with sparse ink (a plot) or carries wide solid runs (a photo, a
    diagram). Split the page into bands at blank rows, keep the bands that
    look like figures, then shrink each to its ink. One look at the
    `preview` confirms or adjusts; that is the eleven crops saved.
    """
    import numpy as np
    mask = _ink_mask(im)
    if strip_stamp:
        mask = strip_arxiv_stamp(mask)
    H, W = mask.shape
    row_ink = mask.mean(axis=1)
    # 1. cut the page into bands at blank rows; a text line is its own band
    bands = _runs(row_ink > 0.002, min_gap=max(2, int(H * 0.003)))

    def widest_run(a, b):
        widest = 0
        for r in range(a, b, max(1, (b - a) // 30)):
            runs = _runs(mask[r], min_gap=2)
            if runs:
                widest = max(widest, max(rb - ra for ra, rb in runs))
        return widest

    # 2. classify: a text line is short and made of word-sized runs; a figure
    # band is tall, or carries a wide solid run (axis, photo, box)
    kinds = []
    for a, b in bands:
        h = b - a
        solid = widest_run(a, b) > 0.12 * W
        if h < 0.025 * H and not solid:
            kinds.append("text")
            continue
        sub = mask[a:b]
        lines = _runs(sub.mean(axis=1) > 0.002, min_gap=1)
        line_h = np.median([l[1] - l[0] for l in lines]) if lines else h
        textlike = len(lines) >= 4 and line_h < 0.02 * H and sub.mean() > 0.03
        kinds.append("text" if textlike and not solid else "fig")

    # 3. merge figure bands separated by less than a caption gap (a sparse
    # plot has blank rows inside it), then drop what is still too short
    out = []
    cur = None
    for (a, b), kind in zip(bands, kinds):
        if kind != "fig":
            if cur is not None and b - a > 0.025 * H:   # a paragraph ends the block
                out.append(cur); cur = None
            continue
        if cur is not None and a - cur[1] < 0.02 * H:
            cur[1] = b
        else:
            if cur is not None:
                out.append(cur)
            cur = [a, b]
    if cur is not None:
        out.append(cur)

    boxes = []
    for a, b in out:
        if b - a < 0.04 * H:
            continue
        sub = mask[a:b]
        cols = sub.any(axis=0)
        if not cols.any():
            continue
        x0 = int(np.argmax(cols)); x1 = W - int(np.argmax(cols[::-1]))
        boxes.append((x0 / W, a / H, x1 / W, b / H))
    return boxes


def preview(pdf: Path, page_idx: int, out: Path, *, dpi: int = 80,
            boxes=None, image: Path | None = None) -> Path:
    """The page with a 0.1 grid labelled in page fractions, and any candidate
    boxes drawn and numbered, so a crop box can be read straight off it."""
    from PIL import Image, ImageDraw

    with tempfile.TemporaryDirectory() as td:
        png = image or render_page(pdf, page_idx, dpi, Path(td))
        im = Image.open(png).convert("RGB")
    W, H = im.size
    d = ImageDraw.Draw(im)
    for i in range(1, 10):
        x, y = int(W * i / 10), int(H * i / 10)
        d.line([(x, 0), (x, H)], fill=(120, 170, 255), width=1)
        d.line([(0, y), (W, y)], fill=(120, 170, 255), width=1)
        d.text((x + 3, 2), f"{i / 10:.1f}", fill=(0, 60, 200))
        d.text((2, y + 2), f"{i / 10:.1f}", fill=(0, 60, 200))
    for n, (x0, y0, x1, y1) in enumerate(boxes or [], start=1):
        d.rectangle([x0 * W, y0 * H, x1 * W, y1 * H], outline=(200, 0, 0), width=3)
        d.text((x0 * W + 4, y0 * H + 4), f"#{n}", fill=(200, 0, 0))
    out.parent.mkdir(parents=True, exist_ok=True)
    im.save(out)
    return out


def crop(pdf: Path, page_idx: int, box: tuple[float, float, float, float],
         out: Path, dpi: int = DEFAULT_DPI, pad: float = PAD, *,
         trim: bool = False, drop_caption: bool = False) -> Path:
    """box is (x0,y0,x1,y1) as fractions of the page."""
    from PIL import Image

    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        png = render_page(pdf, page_idx, dpi, Path(td))
        im = Image.open(png)
        W, H = im.size
        x0, y0, x1, y1 = box
        if x1 <= x0 or y1 <= y0:
            raise SystemExit(f"crop box {box} is not x0,y0,x1,y1 corners")
        rect = (
            max(0, int((x0 - pad) * W)), max(0, int((y0 - pad) * H)),
            min(W, int((x1 + pad) * W)), min(H, int((y1 + pad) * H)),
        )
        if rect[2] - rect[0] < 10 or rect[3] - rect[1] < 10:
            raise SystemExit(f"crop box degenerate: {box}")
        piece = im.crop(rect)
        if trim or drop_caption:
            piece = piece.crop(trim_box(piece, drop_caption=drop_caption))
        piece.save(out)

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
    p_crop.add_argument("--trim", action="store_true",
                        help="cut the blank border off the crop")
    p_crop.add_argument("--drop-caption", action="store_true",
                        help="also cut a caption line left under the figure")
    p_crop.add_argument("-o", "--output", required=True)

    p_prev = sub.add_parser("preview", help="page image with a 0.1 grid, to read a box off")
    p_prev.add_argument("--pdf")
    p_prev.add_argument("--page", type=int, help="1-based page")
    p_prev.add_argument("--image", help="an already-rendered page PNG instead of --pdf")
    p_prev.add_argument("--detect", action="store_true", help="also draw detected blocks")
    p_prev.add_argument("-o", "--output", required=True)

    p_det = sub.add_parser("detect", help="print candidate figure boxes on a page")
    p_det.add_argument("--pdf")
    p_det.add_argument("--page", type=int, help="1-based page")
    p_det.add_argument("--image", help="an already-rendered page PNG instead of --pdf")
    p_det.add_argument("--keep-stamp", action="store_true",
                       help="do not blank the arXiv stamp column")
    p_det.add_argument("-o", "--output", help="also write the preview with boxes drawn")

    args = ap.parse_args()

    if args.cmd in ("preview", "detect"):
        from PIL import Image
        if args.image:
            page_png = Path(args.image)
            page = (args.page or 1) - 1
        else:
            if not args.pdf or not args.page:
                raise SystemExit("need --pdf and --page, or --image")
            if not shutil.which("pdftoppm"):
                raise SystemExit("pdftoppm not found (install poppler-utils)")
            page = args.page - 1
            td = tempfile.mkdtemp()
            page_png = render_page(Path(args.pdf), page, 80, Path(td))
        boxes = []
        if args.cmd == "detect" or args.detect:
            boxes = detect_blocks(Image.open(page_png),
                                  strip_stamp=not getattr(args, "keep_stamp", False))
            if not boxes:
                print("no figure-like block found; use `preview` and read the box off the grid")
            for n, (x0, y0, x1, y1) in enumerate(boxes, start=1):
                print(f"#{n}  --box {x0:.3f},{y0:.3f},{x1:.3f},{y1:.3f}"
                      f"   ({(x1 - x0):.2f} x {(y1 - y0):.2f} of the page)")
        if args.output:
            out = preview(Path(args.pdf) if args.pdf else Path("."), page, Path(args.output),
                          boxes=boxes, image=page_png)
            print(f"{out}  (grid in page fractions{', boxes numbered' if boxes else ''})")
        return

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

    crop(pdf, page, (x0, y0, x1, y1), Path(args.output), args.dpi,
         trim=args.trim, drop_caption=args.drop_caption)


if __name__ == "__main__":
    main()
