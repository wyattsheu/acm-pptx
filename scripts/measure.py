#!/usr/bin/env python3
"""How much room a line of slide text really takes.

Every height compose.py and exhibits.py compute -- a card that stops under
its last line, an equation placed right under the bullets, a group centred
in the space it leaves -- is only as good as the line count behind it. A flat
0.55em-per-glyph guess calls a 44-character Calibri bullet two lines when it
is one, and every such miss is a visible hole on the slide.

So measure with the face the deck is set in: Calibri where Office is
installed, Carlito (metric-identical, what LibreOffice renders with) where it
is not, Microsoft JhengHei for 中文. Without any of them, fall back to the
per-glyph estimate, which errs tall.

    lines("Driving signals injected through FiLM layers", 5.3, 20)  # -> 1
"""
from __future__ import annotations

import functools
import math
import os
import re
from pathlib import Path

CJK = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef\u3000-\u303f]")

_DIRS = [
    Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
    Path.home() / "AppData/Local/Microsoft/Windows/Fonts",
    Path("/Library/Fonts"), Path("/System/Library/Fonts"),
    Path.home() / "Library/Fonts",
    Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"),
    Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
    Path.home() / ".local/share/fonts", Path.home() / ".fonts",
]
_LATIN = {False: ("calibri.ttf", "Calibri.ttf", "Carlito-Regular.ttf"),
          True: ("calibrib.ttf", "Calibri Bold.ttf", "Carlito-Bold.ttf")}
_EA = ("msjh.ttc", "msjh.ttf", "Microsoft JhengHei.ttf", "NotoSansCJK-Regular.ttc",
       "NotoSansCJKtc-Regular.otf", "PingFang.ttc")


def _find(names) -> Path | None:
    for d in _DIRS:
        if not d.is_dir():
            continue
        for n in names:
            if (d / n).is_file():
                return d / n
        for n in names:                       # one level of vendor subfolders
            hits = list(d.glob(f"*/{n}")) or list(d.glob(f"*/*/{n}"))
            if hits:
                return hits[0]
    return None


@functools.lru_cache(maxsize=4)
def _font(kind: str):
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    path = _find(_EA if kind == "ea" else _LATIN[kind == "bold"])
    if path is None:
        return None
    try:
        return ImageFont.truetype(str(path), 100)
    except OSError:
        return None


def em_width(text: str, bold: bool = False) -> float:
    """Width of `text` in ems (1em = the point size)."""
    if not text:
        return 0.0
    latin, ea = _font("bold" if bold else "regular"), _font("ea")
    w = 0.0
    for chunk, is_cjk in _chunks(text):
        f = ea if is_cjk else latin
        if f is not None:
            w += f.getlength(chunk) / 100
        else:
            w += len(chunk) * (1.0 if is_cjk else (0.55 if bold else 0.52))
    return w


def _chunks(text: str):
    buf, kind = "", None
    for ch in text:
        k = bool(CJK.match(ch))
        if kind is not None and k != kind:
            yield buf, kind
            buf = ""
        buf += ch
        kind = k
    if buf:
        yield buf, kind


def width_in(text: str, pt: float, bold: bool = False) -> float:
    return em_width(text, bold) * pt / 72


def lines(text: str, line_in: float, pt: float, bold: bool = False) -> int:
    """Lines `text` wraps to in a column `line_in` inches wide at `pt`.

    Greedy word wrap, as PowerPoint does it; 中文 breaks between any two
    characters. 3% slack, because renderers kern and hint differently.
    """
    if not text:
        return 0
    limit = max(0.2, line_in) * 72 / pt / 1.03          # in ems
    n, cur = 1, 0.0
    for tok in re.findall(r"\s+|[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef\u3000-\u303f]|[^\s\u2e80-\u9fff]+", text):
        w = em_width(tok, bold)
        if tok.isspace():
            cur += w
            continue
        if cur > 0 and cur + w > limit:
            n += 1
            cur = 0.0
        if w > limit:                                   # a word longer than the line
            n += math.ceil(w / limit) - 1
            w = w % limit
        cur += w
    return n


def measured() -> bool:
    """True when a real Latin face was found (estimates otherwise)."""
    return _font("regular") is not None
