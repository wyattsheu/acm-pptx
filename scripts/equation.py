#!/usr/bin/env python3
r"""Rebuild an equation as an image, one colour per term.

Use this only when the slide's job is to EXPLAIN the equation -- to split it
into named roles the way the reference decks do. When the slide's job is to
QUOTE what the paper wrote, crop it instead with figure.py: cropping keeps the
paper's exact notation, and boxes, arrows and labels work on a crop just as
well (see reference-decks.md).

    python scripts/equation.py -o figs/eq_delta.png \
        --part '\Delta' --part '=' \
        --part '\tilde{\Delta}:D2691E' --part '\hat{P}:1F6FEB'

Rendering is matplotlib's mathtext, so no TeX installation is required -- but
it only covers a subset of LaTeX. Multi-line aligned derivations do not lay
out well; crop those.
"""
from __future__ import annotations

import argparse
from pathlib import Path

DEFAULT_COLOR = "111111"
FONT_PT = 44
GAP = 0.028          # fraction of figure width between terms


def render(parts: list[tuple[str, str]], out: Path, *, dpi: int = 220,
           fontsize: int = FONT_PT) -> Path:
    """parts: [(latex_without_dollars, hex_colour), ...] laid out left to right."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    width = max(3.0, 0.9 * sum(len(t) for t, _ in parts) ** 0.6)
    fig = plt.figure(figsize=(width, 1.35))
    fig.patch.set_alpha(0)
    renderer = fig.canvas.get_renderer()

    x = 0.02
    for latex, colour in parts:
        obj = fig.text(x, 0.5, f"${latex}$", fontsize=fontsize,
                       color=f"#{colour}", va="center")
        bbox = obj.get_window_extent(renderer=renderer)
        x += bbox.width / fig.bbox.width + GAP

    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, dpi=dpi, transparent=True, bbox_inches="tight",
                pad_inches=0.04)
    plt.close(fig)
    return out


def parse_part(raw: str) -> tuple[str, str]:
    """'\\hat{P}:1F6FEB' -> ('\\hat{P}', '1F6FEB');  colour is optional."""
    if ":" in raw:
        latex, colour = raw.rsplit(":", 1)
        if len(colour) == 6 and all(c in "0123456789abcdefABCDEF" for c in colour):
            return latex, colour.upper()
    return raw, DEFAULT_COLOR


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--part", action="append", required=True,
                    help="LaTeX term, optionally ':RRGGBB'. Repeat, left to right.")
    ap.add_argument("--fontsize", type=int, default=FONT_PT)
    ap.add_argument("--dpi", type=int, default=220)
    ap.add_argument("-o", "--output", required=True)
    a = ap.parse_args()

    out = render([parse_part(p) for p in a.part], Path(a.output),
                 dpi=a.dpi, fontsize=a.fontsize)
    from PIL import Image
    print(f"{out}  {Image.open(out).size[0]}x{Image.open(out).size[1]}px")


if __name__ == "__main__":
    main()
