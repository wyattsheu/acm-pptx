#!/usr/bin/env bash
# Install acm-pptx as a Codex personal skill (available in every project).
#
#   ./install.sh              copy into ~/.codex/skills/acm-pptx
#   ./install.sh --link       symlink instead, so edits here take effect at once
#   ./install.sh --project    install into ./.codex/skills/acm-pptx
#   ./install.sh --claude     install into ~/.claude/skills/acm-pptx
#   ./install.sh --check      skip installing, just run the dependency + smoke test
set -euo pipefail

SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST="${CODEX_HOME:-$HOME/.codex}/skills/acm-pptx"
MODE=copy
DO_INSTALL=1

for arg in "$@"; do
  case "$arg" in
    --link)    MODE=link ;;
    --project) DEST="$PWD/.codex/skills/acm-pptx" ;;
    --claude)  DEST="$HOME/.claude/skills/acm-pptx" ;;
    --check)   DO_INSTALL=0 ;;
    -h|--help) sed -n '2,8p' "$0"; exit 0 ;;
    *) echo "unknown option: $arg" >&2; exit 2 ;;
  esac
done

PY="${PYTHON:-python3}"

if [ "$DO_INSTALL" = 1 ]; then
  mkdir -p "$(dirname "$DEST")"
  if [ -e "$DEST" ] || [ -L "$DEST" ]; then
    BAK="$DEST.bak.$(date +%Y%m%d%H%M%S)"
    echo "moving existing install to $BAK"
    mv "$DEST" "$BAK"
  fi
  if [ "$MODE" = link ]; then
    ln -s "$SRC" "$DEST"
    echo "linked  $DEST -> $SRC"
  else
    mkdir -p "$DEST"
    (cd "$SRC" && tar --exclude='./.git' -cf - .) | (cd "$DEST" && tar -xf -)
    echo "copied  $SRC -> $DEST"
  fi
else
  DEST="$SRC"
fi

echo
echo "--- dependencies ---"
MISSING=0
for mod in pptx PIL matplotlib defusedxml lxml; do
  if "$PY" -c "import $mod" 2>/dev/null; then
    echo "  ok      python: $mod"
  else
    echo "  MISSING python: $mod"
    MISSING=1
  fi
done
for bin in pdftoppm soffice; do
  if command -v "$bin" >/dev/null 2>&1; then
    echo "  ok      binary: $bin"
  else
    echo "  MISSING binary: $bin"
    MISSING=1
  fi
done
for bin in ffmpeg ffprobe; do
  if command -v "$bin" >/dev/null 2>&1; then
    echo "  ok      binary: $bin"
  else
    echo "  absent  binary: $bin  (needed to embed video: poster frames and"
    echo "                        the H.264 transcode PowerPoint requires)"
  fi
done
if ! "$PY" -c "import markitdown" 2>/dev/null; then
  echo "  absent  python: markitdown  (only needed to read existing decks)"
fi

if [ "$MISSING" = 1 ]; then
  cat <<'EOF'

  pip install python-pptx Pillow matplotlib defusedxml lxml "markitdown[pptx]"
  macOS:  brew install poppler ffmpeg && brew install --cask libreoffice
  Debian: sudo apt install poppler-utils ffmpeg libreoffice
EOF
fi

echo
echo "--- smoke test ---"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
cp -R "$DEST/assets/examples/figs" "$TMP/figs"
# Both examples, because they exercise different halves: the paper talk covers
# figures, annotations, the matrix and both equation forms; the weekly report
# covers the embedded video, which is the piece that silently breaks.
for pair in "outline.paper.example.json talk.pptx" "outline.example.json report.pptx"; do
  set -- $pair
  cp "$DEST/assets/examples/$1" "$TMP/outline.json"
  (
    cd "$TMP"
    echo "  $1"
    "$PY" "$DEST/scripts/build_from_outline.py" outline.json -o "$2"
    "$PY" "$DEST/scripts/compose.py" outline.json "$2"
    "$PY" "$DEST/scripts/qa_check.py" outline.json "$2"
    "$PY" "$DEST/scripts/office/validate.py" "$2" \
          --original "$DEST/assets/acm_template.pptx" | tail -1
  )
done
if command -v ffprobe >/dev/null 2>&1; then
  "$PY" "$DEST/scripts/video.py" probe "$DEST/assets/examples/figs/demo.mp4" >/dev/null \
    && echo "  video: the bundled clip is PowerPoint-playable"
fi
echo
echo "installed at: $DEST"
echo 'in Codex, tell it: "用 $acm-pptx 做這篇論文的 paper study 投影片"'
