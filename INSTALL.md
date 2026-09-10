# acm-pptx v2 — install

This is the **complete** skill, not a patch. Your existing `acm-pptx` files are
all in here, with the paper-study track, the layout engine and the QA gate
added on top. Nothing needs to be merged by hand.

## Claude Code — available in every project

```bash
unzip acm-pptx-v2.zip
cd acm-pptx
./install.sh
```

Copies to `~/.claude/skills/acm-pptx`, checks dependencies, then builds the
example deck and validates it. Expect `0 error(s)` and `All validations PASSED!`.

| Flag | Effect |
|---|---|
| `./install.sh --link` | symlink instead of copy, so edits in this folder take effect immediately (Claude Code follows symlinks in the skills directory) |
| `./install.sh --project` | install into `./.claude/skills/` for one repo only |
| `./install.sh --check` | no install, just dependency check + smoke test |

An existing install is moved to `acm-pptx.bak.<timestamp>`, never overwritten.

## claude.ai / the Claude app

Upload `acm-pptx-v2.zip` under Settings → Capabilities → Skills. Code execution
and file creation must be enabled — `compose.py` and `figure.py` both run code.
The two installs are independent; installing in one does not affect the other.

## Dependencies

```
pip install python-pptx Pillow defusedxml lxml "markitdown[pptx]"
macOS:   brew install poppler && brew install --cask libreoffice
Debian:  sudo apt install poppler-utils libreoffice
```

`markitdown` is only needed to read existing decks. MinerU is optional and
never required.

## What changed from v1.2.0

**New**

- `references/blueprint-paper.md` — the whole paper-talk procedure: input
  contract (PDF / MinerU / Zotero), 5-5-7-6-1 slide budget, where the claim
  lives, the four questions each method component must answer, the required
  ending.
- `references/slide-patterns.md` — the visual fields (`subtitle`, `figure`,
  `annotations`, `matrix`, `equation`, `callout`) and the five layouts.
- `assets/examples/reference-decks.md` — measured teardown of the two decks
  that set the bar. Read this before writing an outline.
- `scripts/figure.py` — crop figures out of a paper PDF at 300 dpi, driven by
  MinerU bboxes or by page fraction.
- `scripts/compose.py` — second pass that lays in figures, captions, on-figure
  annotations, comparison tables, equation bands and callouts. Regions are
  computed from the named layout, never written by hand.
- `scripts/qa_check.py` — gate that exits non-zero.
- `assets/examples/outline.paper.example.json` + `figs/` — working example.
- `scripts/equation.py` -- rebuild an equation with one colour per term
  (matplotlib mathtext, no TeX install). Used when the slide explains a
  decomposition; quoted equations are cropped with `figure.py` instead.
- `install.sh`.

**Added after the first cut**

- `equation` now takes either `src` (quote: crop + `where` list + annotations)
  or `parts` (explain: rebuilt, one colour per term + `captions`). The old
  plain-text form still loads.
- `stage_figure` at the top of the outline plus `"stage": "<name>"` per slide,
  for the overview figure both reference decks repeat on every method slide.
- `"divider": true` draws the thin vertical rule between a text column and a
  figure column. Off by default, and worth using only when the two columns are
  genuinely in contrast (prior work vs ours); on an ordinary text-plus-figure
  slide it just eats the gutter.
- `"layout": "assertion-evidence"` — sentence headline at reading size, the
  exhibit filling the rest, bullets rejected. The one slide structure with a
  controlled experiment behind it (Alley / Garner, Penn State).
- Per-role content schema in `qa_check.py`: a method or results slide with no
  exhibit is now an error; related-work slides warn when they name no
  limitation. Documented as a table in `slide-patterns.md`.
- `qa_check.py --review` prints the claim sequence, a per-slide inventory, and
  a content / design / coherence rubric for you to judge after looking at the
  render.

**Changed**

- `SKILL.md` — paper track wired in; every command now resolves `$SKILL_DIR`
  first, because in Claude Code the working directory is the user's project,
  not the skill folder.
- `references/lab-rules.md` — **one rule was rewritten.** The old row said a
  generic title such as "Results" is a failure mode. Both reference decks use
  plain section labels (`Task Definition`, `Ablation Study`) because the
  breadcrumb above the title already names the section; the old rule pushed
  Claude to turn those into sentences and the output stopped looking like a lab
  deck. The claim now officially lives in the red `subtitle` and the bottom
  `callout`, and the ghost deck test runs on subtitles instead of titles. If
  you disagree, that row and the §Ghost deck test paragraph are the only two
  places to revert.

**Untouched**: `build_from_outline.py`, `add_slide.py`, `clean.py`,
`thumbnail.py`, everything under `scripts/office/`, `template-map.md`,
`outline-schema.md`, `assets/acm_template.pptx`.

## Verified, and not

Built and rendered here: the six-slide example, through
`build_from_outline.py` → `compose.py` → `qa_check.py` (0 errors) →
`validate.py --original` (All validations PASSED) → LibreOffice → 150 dpi
render, inspected. Layouts `text-only`, `figure-right`, `figure-bottom`,
`matrix` with `highlight_row`, `callout`, `subtitle`, and box + label
annotations were all looked at. `install.sh` was run end to end into a clean
home directory.

Not exercised, no MinerU output was available: `figure.py --mineru`. It reads
`*_content_list*.json`, treats `bbox` as 0–1000 and `page_idx` as 0-based, and
prefers `_v2` when both exist — check it on your first real MinerU run. The
`--page/--box` path is what produced both sample crops.

Implemented but not in the test deck: `figure-left`, `figure-full`, and
`arrow` annotations.

`equation` was exercised both ways in the example: a crop from a PDF with a
`where` list, and a rebuilt three-term expression with per-term colour and
captions. Caption text is centred under the whole equation, not under the term
it names -- aligning each caption to its own term is not implemented.

Not built yet: the progress-report blueprint (Claude Code scanning a repo,
diffing last week's todolist). `SKILL.md` currently routes 週報 to
`lab-rules.md` §Weekly meeting content, which is what v1 did.
