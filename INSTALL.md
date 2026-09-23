# acm-pptx — install

This is the **complete** skill, not a patch. Your existing `acm-pptx` files are
all in here, with the paper-study track, the layout engine and the QA gate
added on top. Nothing needs to be merged by hand.

## Codex — available in every project

```bash
unzip acm-pptx-v2.zip
cd acm-pptx
./install.sh
```

Copies to `${CODEX_HOME:-~/.codex}/skills/acm-pptx`, checks dependencies, then builds the
example deck and validates it. Expect `0 error(s)` and `All validations PASSED!`.

| Flag | Effect |
|---|---|
| `./install.sh --link` | symlink instead of copy, so edits in this folder take effect immediately |
| `./install.sh --project` | install into `./.codex/skills/` for one repo only |
| `./install.sh --claude` | install into `~/.claude/skills/` for Claude Code |
| `./install.sh --check` | no install, just dependency check + smoke test |

An existing install is moved to `acm-pptx.bak.<timestamp>`, never overwritten.

## Claude Code / claude.ai

Run `./install.sh --claude` for Claude Code. For claude.ai, upload the skill archive under Settings → Capabilities → Skills. Code execution
and file creation must be enabled — `compose.py` and `figure.py` both run code.
Codex and Claude installs are independent.

## Dependencies

```
pip install python-pptx Pillow matplotlib defusedxml lxml "markitdown[pptx]"
macOS:   brew install poppler ffmpeg && brew install --cask libreoffice
Debian:  sudo apt install poppler-utils ffmpeg libreoffice
```

`markitdown` is only needed to read existing decks; `ffmpeg`/`ffprobe` only to
embed video. MinerU is optional and never required.

## What changed in v2.3.0

**Video works.** `outline.json` had no `video` field, so the only route in was
a bare `python-pptx` `add_movie()` call — which types the media part
`video/unknown`, lands it in `[Content_Types].xml` as an Override instead of
the `mp4 -> video/mp4` Default, and leaves PowerPoint with a file it will not
decode. New `scripts/video.py` (`probe` / `prep` / `poster`) and a `video`
field handled by `compose.py`: correct content type, a real poster frame
instead of the grey speaker icon, autoplay and loop written into `p:timing`,
and a hard refusal — with the fixing command — for anything PowerPoint cannot
play. Verified against the ISO-29500 XSDs, `clean.py` round-trip included.

**Most QA no longer needs a render.** `qa_check.py` now finds leftover
template placeholders (`XXX`, `20XX`, `Ur Name`, `Conf.Name`), a title that
wraps down into the subtitle, and text PowerPoint will auto-shrink — all of
which previously cost ~1600 vision tokens per slide to spot. Font sizes are
resolved through the template's own `buSzPts`, so the overflow estimate is no
longer a guess against a 20pt default. Findings the template itself trips are
baselined out, so a clean deck reaches `0 error(s)` instead of permanently
reporting the `Todolist & Suggestion from Prof.` overflow.

**`render_qa.py --contact N` and `--changed`.** A contact sheet tiles N slides
into one image; an image costs the same whatever it holds, so a 25-slide
review runs ~13k tokens instead of ~40k. `--changed` renders only slides whose
XML moved since the last run.

**Four bugs.** Running `compose.py` twice silently stacked a second copy of
every figure, caption and callout (now refused). A `subtitle` on a conclusion
slide cleared the template's `Todolist & Suggestion from Prof.` label and
stretched it across the content area — and the QA gate was pushing you to add
exactly that subtitle. The title band's bottom edge sat 0.07in below the
subtitle's top. Rebuilt equation images were written into the user's `figs/`
instead of a temp dir.

**The weekly example now passes its own gate.** It previously reported six
errors, which taught the model that failing the gate is normal. It also now
demonstrates `matrix`, `callout` and `video`, and ships a 10 KB clip so
`install.sh` can smoke-test video embedding end to end.

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
  first, because in the coding agent the working directory is the user's project,
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

Not built yet: the progress-report blueprint (the coding agent scanning a repo,
diffing last week's todolist). `SKILL.md` currently routes 週報 to
`lab-rules.md` §Weekly meeting content, which is what v1 did.
