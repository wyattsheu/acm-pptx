#!/usr/bin/env python3
"""Make a video PowerPoint will actually play, and pull a poster frame for it.

Embedding a video is two decisions, and both go wrong silently:

1. **The container lies.** PowerPoint plays H.264 video with AAC audio in an
   MP4/MOV. It does not play HEVC, VP9, AV1, MKV, WebM or most AVI/WMV, and
   screen recorders hand you exactly those. The deck opens fine and the slide
   shows a black rectangle in the meeting.
2. **The part is typed `video/unknown`.** python-pptx defaults `mime_type` to
   `video/unknown`, which lands in `[Content_Types].xml` as an Override rather
   than the `mp4 -> video/mp4` Default that PowerPoint looks for. `compose.py`
   always passes the real type; this module is where that mapping lives.

    python scripts/video.py probe demo.mov          # will PowerPoint play it?
    python scripts/video.py prep demo.mov -o figs/demo.mp4 --clip 0:03-0:18
    python scripts/video.py poster figs/demo.mp4 -o figs/demo.png

`prep` is the one to reach for: it re-encodes to H.264 High / yuv420p / AAC,
caps the long edge at the slide's own pixel width, moves the index to the front
so PowerPoint can start without reading the whole file, and trims. A 90-second
4K screen capture goes from 400 MB to about 6 MB, and 6 MB is what you want
inside a deck you email to your advisor.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# extension -> the content type PowerPoint expects to find in [Content_Types].xml.
# Keys are what python-pptx's own default_content_types table recognises, so the
# part is registered as a Default and not as an Override.
MIME = {
    ".mp4": "video/mp4",
    ".m4v": "video/mp4",
    ".mov": "video/quicktime",
    ".mpg": "video/mpeg",
    ".mpeg": "video/mpeg",
    ".wmv": "video/x-ms-wmv",
    ".avi": "video/avi",
    ".asf": "video/x-ms-asf",
}
# Containers PowerPoint opens at all. Anything else has to be transcoded.
PLAYABLE_EXT = {".mp4", ".m4v", ".mov", ".wmv", ".avi", ".mpg", ".mpeg", ".asf"}
# Codecs PowerPoint decodes on both Windows and macOS without a codec pack.
PLAYABLE_VCODEC = {"h264", "mpeg4", "wmv3", "wmv2", "msmpeg4v3", "mpeg2video"}
PLAYABLE_ACODEC = {"aac", "mp3", "wmav2", "wmapro", "pcm_s16le", "ac3"}

SLIDE_PX = 1920           # the deck is 13.333in; beyond this the projector cannot show it
POSTER_PX = 1600          # the still only has to survive a projector, not a print
SIZE_WARN_MB = 40         # a deck past ~50MB starts bouncing off mail servers
DURATION_WARN_S = 45      # a lab-meeting clip longer than this is not a clip


def _have(tool: str) -> bool:
    return shutil.which(tool) is not None


def mime_for(path: Path) -> str:
    """The content type to hand python-pptx. Never `video/unknown`."""
    return MIME.get(path.suffix.lower(), "video/mp4")


def probe(src: Path) -> dict:
    """Container, codecs, size and duration. Empty dict if ffprobe is missing."""
    if not _have("ffprobe"):
        return {}
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-print_format", "json",
         "-show_format", "-show_streams", str(src)],
        capture_output=True, text=True)
    if out.returncode != 0:
        raise SystemExit(f"ffprobe could not read {src}: {out.stderr.strip()}")
    data = json.loads(out.stdout or "{}")
    v = next((s for s in data.get("streams", []) if s.get("codec_type") == "video"), {})
    a = next((s for s in data.get("streams", []) if s.get("codec_type") == "audio"), {})
    fmt = data.get("format", {})
    return {
        "vcodec": v.get("codec_name"),
        "acodec": a.get("codec_name"),
        "pix_fmt": v.get("pix_fmt"),
        "width": int(v.get("width") or 0),
        "height": int(v.get("height") or 0),
        "duration": float(fmt.get("duration") or 0.0),
        "size_mb": src.stat().st_size / 1e6,
        "container": Path(src).suffix.lower(),
    }


def problems(src: Path, info: dict | None = None) -> list[tuple[str, str]]:
    """Why PowerPoint will refuse, or complain about, this file.

    Returns `(level, message)`; `error` means the slide shows a black box in the
    meeting, `warn` means it plays but you will regret the file size or length.
    """
    info = probe(src) if info is None else info
    out: list[tuple[str, str]] = []
    ext = src.suffix.lower()
    if ext not in PLAYABLE_EXT:
        out.append(("error", f"container {ext or '(none)'} is not one PowerPoint "
                             f"opens"))
    if not info:
        return out                       # no ffprobe: the extension is all we know
    if info["vcodec"] and info["vcodec"] not in PLAYABLE_VCODEC:
        out.append(("error", f"video codec {info['vcodec']} is not decoded by "
                             f"PowerPoint, which wants h264"))
    if info["pix_fmt"] and info["pix_fmt"] not in ("yuv420p", "yuvj420p"):
        out.append(("error", f"pixel format {info['pix_fmt']} plays black in "
                             f"PowerPoint"))
    if info["acodec"] and info["acodec"] not in PLAYABLE_ACODEC:
        out.append(("error", f"audio codec {info['acodec']} is not decoded by "
                             f"PowerPoint, which wants aac"))
    if info["size_mb"] > SIZE_WARN_MB:
        out.append(("warn", f"{info['size_mb']:.0f} MB embedded in the deck - "
                            f"prep --max-width {SLIDE_PX} shrinks it"))
    if info["duration"] > DURATION_WARN_S:
        out.append(("warn", f"{info['duration']:.0f}s long - trim it with "
                            f"`prep --clip start-end`; nobody watches a minute "
                            f"of demo in a lab meeting"))
    return out


def fix_command(src: Path, out_dir: str = "figs") -> str:
    """The one command that turns this file into something a deck can carry."""
    return (f"python scripts/video.py prep {src} -o {out_dir}/{src.stem}.mp4 "
            f"--clip 0:00-0:15")


def _tc(spec: str) -> str:
    """`0:18`, `18`, `1:02.5` -> seconds ffmpeg accepts."""
    parts = str(spec).split(":")
    sec = 0.0
    for p in parts:
        sec = sec * 60 + float(p)
    return f"{sec:.3f}"


def prep(src: Path, out: Path, *, clip: str | None = None,
         max_width: int = SLIDE_PX, crf: int = 23, mute: bool = False) -> Path:
    """Re-encode to what PowerPoint plays: H.264 High / yuv420p / AAC / faststart."""
    if not _have("ffmpeg"):
        raise SystemExit("ffmpeg not found - install it, or hand compose.py an "
                         "mp4 that is already H.264")
    out.parent.mkdir(parents=True, exist_ok=True)
    cmd = ["ffmpeg", "-v", "error", "-y"]
    if clip:
        start, _, end = clip.partition("-")
        if start.strip():
            cmd += ["-ss", _tc(start.strip())]
        if end.strip():
            cmd += ["-to", _tc(end.strip())]
    cmd += ["-i", str(src)]
    # -2 keeps the height even, which yuv420p requires; min() never upscales
    cmd += ["-vf", f"scale='min({max_width},iw)':-2",
            "-c:v", "libx264", "-profile:v", "high", "-pix_fmt", "yuv420p",
            "-crf", str(crf), "-preset", "medium"]
    info = probe(src)
    if mute or not info.get("acodec"):
        cmd += ["-an"]
    else:
        cmd += ["-c:a", "aac", "-b:a", "128k"]
    cmd += ["-movflags", "+faststart", str(out)]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(f"ffmpeg failed on {src}:\n{r.stderr.strip()[-800:]}")
    return out


def poster(src: Path, out: Path, at: str | float | None = None) -> Path:
    """A still for the slide to show before anyone clicks play.

    PowerPoint falls back to a grey loudspeaker icon when a movie has no poster,
    which is what a "broken" video slide usually is. Default is 10% in, past the
    fade-from-black most screen recordings open with.
    """
    out.parent.mkdir(parents=True, exist_ok=True)
    if _have("ffmpeg"):
        if at is None:
            dur = probe(src).get("duration") or 0.0
            at = max(0.0, min(1.0, dur * 0.1)) if dur else 0.0
        ts = _tc(at) if isinstance(at, str) else f"{float(at):.3f}"
        # A poster wider than the slide itself is dead weight inside the .pptx.
        r = subprocess.run(
            ["ffmpeg", "-v", "error", "-y", "-ss", ts, "-i", str(src),
             "-frames:v", "1", "-vf", f"scale='min({POSTER_PX},iw)':-2",
             "-q:v", "3", str(out)],
            capture_output=True, text=True)
        if r.returncode == 0 and out.exists() and out.stat().st_size > 0:
            return out
    return _placeholder_poster(src, out)


def _placeholder_poster(src: Path, out: Path) -> Path:
    """No ffmpeg: a plain 16:9 card, so the slide never shows the speaker icon."""
    from PIL import Image, ImageDraw
    info = probe(src)
    w, h = info.get("width") or 1280, info.get("height") or 720
    img = Image.new("RGB", (min(w, 1280), max(1, int(min(w, 1280) * h / max(w, 1)))),
                    (0x1A, 0x1A, 0x1A))
    d = ImageDraw.Draw(img)
    cx, cy = img.width / 2, img.height / 2
    r = min(img.width, img.height) * 0.12
    d.polygon([(cx - r * 0.4, cy - r * 0.7), (cx - r * 0.4, cy + r * 0.7),
               (cx + r * 0.8, cy)], fill=(0xC0, 0x00, 0x00))
    img.save(out)
    return out


def auto_poster(src: Path, at: str | float | None = None) -> Path:
    """Poster into a temp dir, not the user's directory. Cached per source file."""
    cache = Path(tempfile.gettempdir()) / "acm-poster"
    stamp = "" if at is None else f"-{str(at).replace(':', '_')}"
    # jpg, not png: a frame out of an already-lossy video gains nothing from
    # lossless encoding and costs several MB inside the deck.
    out = cache / f"{src.stem}{stamp}-{int(src.stat().st_mtime)}.jpg"
    if out.exists():
        return out
    return poster(src, out, at)


# ---------------------------------------------------------------------- cli

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("probe", help="will PowerPoint play this?")
    p.add_argument("src")

    p = sub.add_parser("prep", help="re-encode to a deck-safe mp4")
    p.add_argument("src")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--clip", help="trim, e.g. 0:03-0:18 (either end may be empty)")
    p.add_argument("--max-width", type=int, default=SLIDE_PX)
    p.add_argument("--crf", type=int, default=23, help="18 sharper, 28 smaller")
    p.add_argument("--mute", action="store_true", help="drop the audio track")

    p = sub.add_parser("poster", help="pull one frame out as the still")
    p.add_argument("src")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--at", help="timestamp, e.g. 0:04 (default: 10%% in)")

    a = ap.parse_args()
    src = Path(a.src)
    if not src.exists():
        raise SystemExit(f"not found: {src}")

    if a.cmd == "probe":
        info = probe(src)
        if not info:
            print("ffprobe not installed - only the container could be checked")
        else:
            print(f"{info['container']}  {info['width']}x{info['height']}  "
                  f"{info['duration']:.1f}s  {info['size_mb']:.1f} MB")
            print(f"video {info['vcodec']} / {info['pix_fmt']}   "
                  f"audio {info['acodec'] or '(none)'}")
        issues = problems(src, info or None)
        for level, msg in issues:
            print(f"  {'!' if level == 'error' else '?'} {msg}")
        if not issues:
            print("PowerPoint will play this.")
        else:
            print(f"\nFix with: {fix_command(src)}")
        sys.exit(1 if any(lv == "error" for lv, _ in issues) else 0)

    if a.cmd == "prep":
        out = prep(src, Path(a.output), clip=a.clip, max_width=a.max_width,
                   crf=a.crf, mute=a.mute)
        info = probe(out)
        print(f"wrote {out} ({info.get('size_mb', 0):.1f} MB, "
              f"{info.get('width', 0)}x{info.get('height', 0)}, "
              f"{info.get('duration', 0):.1f}s)")
        return

    out = poster(src, Path(a.output), a.at)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
