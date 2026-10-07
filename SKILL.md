---
name: acm-pptx
description: "Build ACM Lab (NYCU) presentations on the lab's official template — weekly progress reports and paper-study talks. Use this skill whenever the user asks for a 進度報告, 週報, progress report, 組會投影片, lab presentation, paper presentation, 論文報告, paper study, or any .pptx/.potx in ACM Lab format, and whenever such a deck is created, edited or read. Also use it when the user hands you a paper PDF or a MinerU directory and asks for slides, when they hand you an outline, when they want a demo video or supplementary clip inside a deck, or when they mention Prof. Huang Ching-Chun's lab meeting. Never build ACM Lab slides from scratch with pptxgenjs — always clone the bundled template."
license: Lab-internal use
metadata:
  version: "2.6.0"
  template_version: "acm_template.pptx (22 slides, 13.333in x 7.5in)"
---

# ACM Lab slide builder

The lab template is the only source of visual truth. Every slide in
`assets/acm_template.pptx` carries its own layout with the section colour bar
and the breadcrumb strip already positioned. **Never rebuild a slide — clone
the template slide whose role matches, then compose the content inside it.**
The template fixes the lab identity, not the form of the explanation. Choose
the evidence first; a table is one option, not the default content layout.

## Which track

| The user wants | Read first | Then |
|---|---|---|
| 週報 / 進度報告 | `references/lab-rules.md` §Weekly meeting content | build |
| 論文報告 / paper study, or hands you a paper PDF | `references/blueprint-paper.md` | build |
| Either | `references/lab-rules.md`, `references/slide-patterns.md` | — |

`references/lab-rules.md` is the content standard distilled from the two ACM
handbooks. `references/template-map.md` lists the roles;
`python scripts/build_from_outline.py --roles` prints the authoritative list.
`assets/examples/reference-decks.md` is a measured teardown of two decks that
set the bar — read it before writing an outline, it is worth more than any
rule in this file.

## The build path

**Every command below runs from the user's working directory, not from this
skill folder.** Resolve the skill root once and prefix every script with it:

```bash
SKILL_DIR="${CODEX_HOME:-$HOME/.codex}/skills/acm-pptx"
[ -f "$SKILL_DIR/SKILL.md" ] || SKILL_DIR="$PWD/.codex/skills/acm-pptx"
[ -f "$SKILL_DIR/SKILL.md" ] || SKILL_DIR="$HOME/.claude/skills/acm-pptx"
[ -f "$SKILL_DIR/SKILL.md" ] || SKILL_DIR="$PWD/acm-pptx"
```

That covers a Codex personal skill, a project skill, a Claude Code skill and
a cloned repo. If none exists, ask the user where they installed it. Commands
below say `python3`; on a machine where only `python` exists, alias it once.

```bash
python3 "$SKILL_DIR/scripts/build_from_outline.py" --roles        # role table
# write outline.json  (schema: references/outline-schema.md
#                      + visual fields: references/slide-patterns.md)
DECK=2026-09-18-Report.pptx   # date + what it is; see naming below
python3 "$SKILL_DIR/scripts/build_from_outline.py" outline.json -o "$DECK"   # text
python3 "$SKILL_DIR/scripts/compose.py" outline.json "$DECK"                 # exhibits
python3 "$SKILL_DIR/scripts/render_qa.py" outline.json "$DECK"               # gate + render what needs eyes
python3 "$SKILL_DIR/scripts/qa_check.py" outline.json "$DECK" --review       # before handing it over
python3 "$SKILL_DIR/scripts/office/validate.py" "$DECK" --original "$SKILL_DIR/assets/acm_template.pptx"
```

Two passes, in that order. `build_from_outline.py` clones role slides and
replaces text run-by-run so fonts, bullet tiers and table styling survive.
`compose.py` then resizes the body placeholder and lays in the figure or video,
annotations, comparison table, native editable diagram, equation band and
callout — it computes every region from the named layout, which is why nothing
overlaps. Running
`compose.py` twice on one deck is refused: rebuild first.

`build_from_outline.py` finds its own template, so `--template` is only needed
when the lab ships a revised one. `--original` on validate is required: the lab
template has pre-existing XSD quirks, and baselining against it keeps a real
regression from hiding.

## What lands on disk

**One `.pptx` in the working directory, and nothing else.** Name it
`YYYY-MM-DD-<what>.pptx` — `2026-09-18-Report.pptx`, `2026-09-18-TADSR.pptx`.
`Report` is the safe default; never `talk.pptx` or `output.pptx`. If the user
gave a name, use theirs.

`outline.json` is a working file: keep it, it is how edits are made, but it is
not a deliverable. Figures and video the user gave you live in a local `figs/`.
Everything the scripts generate — PDFs, rendered PNGs, poster frames, rebuilt
equations — goes to a temp dir, never to the user's directory. Do not convert
to PDF unless asked, and do not create a `build/` tree for a deck with two
figures.

**Rebuilding the deck is free; looking at it is not.** These scripts are local
compute, so never hand-patch a `.pptx` to dodge a rebuild. A rendered slide is
~1600 tokens of vision input, and that is where a session's budget goes.

## Before you write the outline

Inspect the available paper figures, plots, logs, screenshots and demo clips
before assigning layouts. Read `references/evidence-selection.md` for diagram,
experiment and paper-figure workflows. For each main claim, identify the actual
source and the visual relationship the audience needs to see. A text-filled
`matrix` does not become evidence merely because QA accepts it.


1. **Ask before inventing content.** If the user has not said what they did
   this week, ask. Never fabricate experimental numbers, metrics, or paper
   results.
2. **Get the outline confirmed before building anything over 10 slides.**
   Titles, subtitles, and which exhibit sits on each slide. Apply this to the
   actual deck length; paper talks can also be short.
3. **Ghost deck test on the titles.** Read the `title` lines in sequence, with
   a `subtitle` standing in wherever the title is a bare section label. They
   must carry the whole argument alone. If they read as a list of topics —
   `Method`, `Results`, `Conclusion` — the titles are not yet doing their job.
   `qa_check.py --review` prints this sequence for you.

## Choosing roles

Include only the tracks that have real content; for an empty track use its
`*_summary` slide with every term set to `NONE`.

| Situation | Roles |
|---|---|
| Weekly report, project only | `cover`, `summary`, `project_summary` … `project_conclusion` |
| Weekly report, research only | `cover`, `summary`, `research_summary` … `research_conclusion` |
| Both tracks | project block then research block; one `summary` at the front |
| Paper study | `cover`, then repeat `paper_intro`, `paper_related`, `paper_method`, `paper_results` as needed; `paper_conclusion` |
| Anything else | `others` |

Roles repeat by design — the breadcrumb stays on the right section. Section
counts follow the material and speaking time, not a fixed quota.

## Writing the content

- **The claim goes in the title.** Name the method, the mechanism or the
  finding: `Time-Aware Encoder (TAE)`, `Continuous Control by Scaling LoRA` —
  not `Method`, not `Approach`. Reuse one title across consecutive slides when
  the point continues; both reference decks do. A generic section label is fine
  on a results slide, where the figure is the argument, and otherwise needs a
  `subtitle` under it to carry the point. Keep it to one line — past about 50
  characters the template shrinks it below the 24pt floor, and `qa_check.py`
  says so.
- **Do not put a red box on every slide.** Across the two 2026 reference decks
  — 34 slides — there are zero callout boxes. Reserve `callout` for the one or
  two slides where the talk turns, such as the research question between
  motivation and contribution; for emphasis everywhere else use an inline red
  run inside a sentence, the way TADSR marks `distribution-to-distribution
  matching`. A box on every slide drains the two that deserve one.
- **Telegraphic wording.** `Cost down 23% (p < 0.01)`, never `Our results
  demonstrate that costs were significantly reduced`. But never compress away
  the baseline a number beats, its units and dataset, or an acronym's first
  expansion. `references/lab-rules.md` §Writing discipline is the standard.
- **One idea per slide**, ~40 words of body text as a hard ceiling. The cap
  covers slide body only — never captions, annotations, or notes.
- **One exhibit per slide, annotated.** Mark the decisive point on the figure
  itself. Review text-only stretches for missing evidence; do not add a decorative
  table just to meet an exhibit quota. A coordinated before/after pair or
  multi-panel comparison can be one exhibit.
- **Keep authored structure editable.** Build comparison tables as native
  PowerPoint tables (`matrix`) and new workflows, pipelines and architecture
  diagrams as native PowerPoint nodes and connectors (`diagram`). Never render
  Mermaid, Graphviz, matplotlib patches or HTML into a PNG for these. Images
  are for source paper figures, photos, screenshots and experimental output
  whose pixels are evidence. Read `references/evidence-selection.md` before
  choosing between `diagram` and `figure`.
- **Do not use the generic AI-card look.** Shape means function: rectangles for
  processes, diamonds for decisions, terminators only at starts/ends, cylinders
  only for storage. Use one neutral fill, thin grey connectors, square alignment
  and whitespace. Reserve the lab red for the node or row currently being
  argued; do not make every node a rounded pastel card.
- **Find the crop box before cropping.** `figure.py preview --pdf p.pdf
  --page 3 -o page3.png` draws a 0.1 grid in page fractions; `figure.py
  detect` prints candidate `--box` values for the figure blocks it finds.
  Read the box off one image instead of iterating on crops. `--trim` cuts the
  blank border, `--drop-caption` the caption line under it, and the arXiv
  stamp in the left margin is excluded by default.
- **Preserve useful source figures.** Crop and enlarge paper architecture
  figures, qualitative panels and plots with `scripts/figure.py` at 300 dpi;
  retain legends and cite the source. Re-plot when verified data is available
  and it improves the explanation, with axis labels 16pt or larger. Label
  simplified diagrams as adaptations; do not recreate experimental imagery.
- **Pick the exhibit by the shape of the content, not by habit.** A linear pipeline
  is a `flow`, a branching one a `diagram`; camps, contributions, pros-and-cons are `cards`; the number
  the talk turns on is a `bignum`; the field is a `quadrant`; a `matrix` is
  for the one comparison where the cells are the argument. The first decks
  this skill built were sixteen slides with seven tables and two bullet
  lists, and that is what the audience remembers. `qa_check.py` warns past a
  third of content slides as tables, and on three of any one shape in a row.
  When none fits, `"draw": "design.py:fn"` hands the slide's exhibit region
  to your own function. All of it: `references/slide-patterns.md`
  §Graphic exhibits.
- **A wide figure gets the full width.** A paper's architecture diagram
  squeezed into a half column lands at ~5in and its labels vanish. Use
  `figure-bottom` or `figure-full`, or crop to the one component the slide
  is about; `qa_check.py` computes where a figure lands and warns.
- **Say what is ours.** A paper talk carries at least one slide marked
  `"ours": true` — what you tried, what transfers to our work, where you
  disagree. It draws a small `OUR TAKE` tag, and `qa_check.py` warns when a
  paper deck has none. A deck that only relays the paper has skipped the part
  the meeting is for.
- **Boxes: `at` is `[x, y, w, h]`, `xyxy` is corners.** `figure.py crop`
  takes corners, annotations take corner-plus-size, and mixing them puts the
  red box in the wrong place. If you have corners, write `"xyxy"` and do not
  convert. The full table is `references/slide-patterns.md` §Coordinates.
- **Video only for claims a still cannot make** — tracking jitter, temporal
  flicker, a robot finishing the task. See §Video below; everything else is a
  figure.
- **Bullet tiers**: `level: 0` is the bold header tier, `level: 1` the
  bulleted tier. Never fake indentation with spaces or dashes.
- **中文 is first-class.** Every run carrying CJK is written `lang="zh-TW"`
  with an East-Asian typeface (`Microsoft JhengHei`; set `meta.cjk_font` to
  change it), so LibreOffice previews stop overlapping glyphs and PowerPoint
  stops guessing. The 40-word body cap counts 1.8 中文字 as one word, about
  70 characters; a claim's length is judged the same way. Chinese captions
  are set upright at 12pt or more, because italic CJK is a fake slant.
- **Bullets at `level: 1` always indent.** A template slide with one tier
  (the conclusion pages) gets a bulleted, indented tier synthesised, and the
  build says so; a header-plus-children structure no longer flattens.
- **Draw figures to the slot.** `compose.py --sizes` prints the exhibit box
  per layout (`figure-right` 5.75×4.68in, `figure-full` 12.10×4.68in). Make
  the plot that size and 12pt labels stay 12pt; make it twice the size and
  they land at 6pt. On `figure-bottom` the text strip is only as tall as its
  lines, so a one-line lead-in leaves the figure ~4.3in high instead of ~2.8.
- **Speaker notes** carry what you would say — in 中文, technical terms in
  English, full sentences, three to six per content slide, no length limit.
  Terse slides are only safe when the notes are full.
- **Tables** (`conclusion`, `paper_list`, `summary`) take a list; unused rows
  are blanked. Leave the numbering column alone. A comparison table you are
  building yourself is a `matrix`, not one of these. Keep only decisive rows
  and columns; short numeric columns should not receive the same width as the
  label column.
- Emoji status markers in the summary task table: 🔄 in progress · ✅ done ·
  ❌ dropped · ⏳ waiting · 🆕 new · 🔜 upcoming · 📌 important · 💤 on hold ·
  🛠️ needs fixing · 💡 idea · 📅 scheduled.

## Video

A deck can embed a demo clip or a paper's supplementary video. It goes in the
`video` field, where a `figure` would go, and takes the same layouts — full
field list in `references/slide-patterns.md`.

**Prepare the file first.** PowerPoint decodes H.264 video with AAC audio in an
`.mp4` or `.mov`, and nothing else. Screen recorders and paper websites hand
you HEVC, VP9, `.mkv`, `.webm` — all of which build into a deck with no error
and show a black rectangle in the meeting. Convert up front:

```bash
python3 "$SKILL_DIR/scripts/video.py" probe raw.mov                    # playable?
python3 "$SKILL_DIR/scripts/video.py" prep raw.mov -o figs/demo.mp4 --clip 0:03-0:18
```

Then name `figs/demo.mp4` in the outline. `compose.py` embeds it with the
correct media content type, pulls a poster frame so the slide is not a grey
speaker icon, and marks it `▶ 0:15`. It refuses any file PowerPoint could not
play, and prints the `prep` command that fixes it.

Four rules, none negotiable:

- **Trim to about fifteen seconds** — every second is embedded in the file the
  user emails their advisor.
- **`"autoplay": true, "loop": true`** on the one clip the slide is about, so
  you keep talking instead of hunting for a play button; everything else stays
  click-to-play.
- **Present from PowerPoint.** A PDF export drops the video entirely. Say so
  when you hand the deck over.
- **Embed, never link.** A linked video dies the moment the file moves.

A render shows the poster frame, not playback. Confirming the clip plays is the
user's job, in PowerPoint, once.

## Editing an existing deck instead

```bash
markitdown deck.pptx                                    # read content
python3 "$SKILL_DIR/scripts/thumbnail.py" deck.pptx deck-thumbs   # see layouts
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python3 "$SKILL_DIR/scripts/add_slide.py" unpacked/ slide2.xml --after slide2.xml
python3 "$SKILL_DIR/scripts/clean.py" unpacked/
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python3 "$SKILL_DIR/scripts/office/validate.py" out.pptx --original deck.pptx
```

Do all structural work — add, delete, reorder — **before** editing any slide's
content: `add_slide.py` copies verbatim, and `clean.py` deletes any slide
missing from `<p:sldIdLst>`.

If you script an XML transform, parse with `defusedxml.minidom`; round-tripping
OOXML through `xml.etree.ElementTree` rewrites namespace prefixes and corrupts
the deck.

## QA

`qa_check.py` is the cheap pass, and it now covers most of what a render used
to be opened for. It exits non-zero on: a content slide with no claim line,
body text over the word cap, a borrowed figure or video with no source, a video
PowerPoint cannot decode, text that will not fit its box, a title that wraps
down into the line beneath it, a template placeholder (`XXX`, `20XX`,
`Ur Name`, `Conf.Name`) still in the deck, a method or results slide with no
exhibit, an assertion-evidence slide carrying bullets or no evidence, a
generic title with nothing under it, an annotation box past the picture edge
(corners written as `at`), and a `draw` target that does not exist. For every
`diagram`, it also verifies that all declared nodes, edges and groups exist in
the `.pptx` as named native PowerPoint objects. It warns on thin notes,
piled-up callouts, a high text-only ratio, text PowerPoint will auto-shrink, a
table-heavy deck or three tables or three text-only slides in a row, a wide
figure squeezed into a half column, and a paper talk with no `ours` slide.
Findings the lab
template already trips on its own are baselined out, so a clean deck really
does reach `0 error(s)`. `--review` adds the claim sequence, a slide inventory
and a rubric for content, design and coherence — run it once before handing the
deck over.

What is left for your eyes is genuinely visual: a crop that needed to be
tighter, an annotation covering what it points at, an unreadable axis label, a
slide whose balance is wrong. That is what `render_qa.py` is for.

Work the loop this way:

1. `render_qa.py` — fix every error it reports, rebuild, re-run. Do not open an
   image while the exit code is non-zero.
2. Look only at the files it printed.
3. Fix what you saw by **editing `outline.json` in place** — a targeted string
   replacement on the one slide. Never re-emit the whole outline; at 25 slides
   that is ~10k output tokens to change one callout.
4. Rebuild, then `render_qa.py --changed` — only the slides whose XML moved;
   or `render_qa.py --pages 18-22` for exactly the slides you want to see.
5. Once, at the end: `render_qa.py --all --contact 6` plus
   `qa_check.py --review`.

`--contact N` tiles N slides into one image. An image costs ~1600 tokens
whatever it holds, so a 25-slide deck reviewed as contact sheets is ~13k
instead of ~40k. Read a sheet for balance and layout, not for caption text —
open individual slides for figure labels, legends and citations that are too
small on the sheet. XML checks cannot read text embedded in a paper figure.

`--dpi` defaults to 110, legible down to figure captions. Going above it buys
no detail: a 13.333in slide is downscaled to 1568px on its long edge
regardless, so 150 dpi and 300 dpi are the identical image at the identical
price.

One quirk is inherited, not yours: the `Todolist & Suggestion from Prof.` label
on the conclusion slides overflows its box in the template itself. `qa_check.py`
knows and stays quiet about it. Leave it.

## Dependencies

`python-pptx`, `markitdown[pptx]`, `Pillow`, `matplotlib`, `defusedxml`, `lxml` (pip) ·
LibreOffice via `scripts/office/soffice.py` · `pdftoppm` (Poppler, also used by
`figure.py`) · `ffmpeg`/`ffprobe`, needed only for the `video` field. MinerU is
optional and never required.

If the environment cannot execute code, you are in the wrong package — use
`acm-slides-plain`, which produces the outline and stops there.
