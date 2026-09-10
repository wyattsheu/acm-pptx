---
name: acm-pptx
description: "Build ACM Lab (NYCU) presentations on the lab's official template — weekly progress reports, project/research updates, and paper-study talks. Use this skill whenever the user asks for a 進度報告, 週報, progress report, 組會投影片, lab presentation, paper presentation, 論文報告, paper study, or any .pptx/.potx that should follow ACM Lab format, and whenever a deck is being created, edited, or read for this lab. Also use it when the user hands you a paper PDF or a MinerU output directory and asks for slides, when they hand you an outline and ask for slides, or when they mention Prof. Huang Ching-Chun's lab meeting. Do not build ACM Lab slides from scratch with pptxgenjs — always clone the bundled template."
license: Lab-internal use
metadata:
  version: "2.1.0"
  template_version: "acm_template.pptx (22 slides, 13.333in x 7.5in)"
---

# ACM Lab slide builder

The lab template is the only source of visual truth. Every slide in
`assets/acm_template.pptx` carries its own layout with the section colour bar
and the breadcrumb strip already positioned. **Never rebuild a slide — clone
the template slide whose role matches, then replace the text.**

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
[ -f "$SKILL_DIR/SKILL.md" ] || SKILL_DIR="$PWD/acm-pptx"
```

The fallbacks cover a Codex personal skill (`$CODEX_HOME/skills/` or
`~/.codex/skills/`), a project skill (`.codex/skills/`), and a cloned repo in
the working directory. If none exists, ask the user where they installed it.

```bash
python "$SKILL_DIR/scripts/build_from_outline.py" --roles        # role table
# write outline.json  (schema: references/outline-schema.md
#                      + visual fields: references/slide-patterns.md)
python "$SKILL_DIR/scripts/build_from_outline.py" outline.json -o talk.pptx   # text
python "$SKILL_DIR/scripts/compose.py" outline.json talk.pptx                 # figures, callouts, tables
python "$SKILL_DIR/scripts/qa_check.py" outline.json talk.pptx                # gate — must exit 0
python "$SKILL_DIR/scripts/qa_check.py" outline.json talk.pptx --review       # before handing it over
python "$SKILL_DIR/scripts/office/validate.py" talk.pptx --original "$SKILL_DIR/assets/acm_template.pptx"
python "$SKILL_DIR/scripts/office/soffice.py" --headless --convert-to pdf talk.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 150 talk.pdf slide
```

`build_from_outline.py` finds its own template, so `--template` is only needed
when the lab ships a revised one. Figures named in `outline.json` are resolved
relative to the working directory — keep them in a local `figs/`.

Two passes, in that order. `build_from_outline.py` clones role slides and
replaces text run-by-run so fonts, bullet tiers and table styling survive.
`compose.py` then resizes the body placeholder and lays in the figure,
annotations, comparison table, equation band and callout — it computes every
region from the named layout, which is why nothing overlaps.

`--original` on validate is required: the lab template has pre-existing XSD
quirks, and baselining against it keeps a real regression from hiding.

Render QA at **150 dpi or higher**. At 100 dpi, JPEG artifacts look like stray
shapes and you will chase bugs that do not exist.

## Before you write the outline

1. **Ask before inventing content.** If the user has not said what they did
   this week, ask. Never fabricate experimental numbers, metrics, or paper
   results.
2. **Get the outline confirmed before building anything over 10 slides.**
   Titles, subtitles, and which exhibit sits on each slide. A paper talk is
   always over 10 slides, so it is always confirmed first.
3. **Ghost deck test on the subtitles.** The template's titles are section
   labels and both reference decks leave them that way. Read the `subtitle`
   lines in sequence — they must carry the whole argument alone. If they read
   as a list of topics, fix the outline before building.

## Choosing roles

Include only the tracks that have real content; for an empty track use its
`*_summary` slide with every term set to `NONE`.

| Situation | Roles |
|---|---|
| Weekly report, project only | `cover`, `summary`, `project_summary` … `project_conclusion` |
| Weekly report, research only | `cover`, `summary`, `research_summary` … `research_conclusion` |
| Both tracks | project block then research block; one `summary` at the front |
| Paper study | `cover`, `paper_intro` ×5, `paper_related` ×5, `paper_method` ×7, `paper_results` ×6, `paper_conclusion` |
| Anything else | `others` |

Roles repeat by design — the breadcrumb stays on the right section.

## Writing the content

- **The claim goes in `subtitle` and `callout`, not the title.** One red line
  under the section label, one boxed sentence at the bottom. Every content
  slide gets at least one of the two; `qa_check.py` enforces it.
- **Telegraphic wording.** `Cost down 23% (p < 0.01)`, never `Our results
  demonstrate that costs were significantly reduced`. But never compress away
  the baseline a number beats, its units and dataset, or an acronym's first
  expansion. `references/lab-rules.md` §Writing discipline is the standard.
- **One idea per slide**, ~40 words of body text as a hard ceiling. The cap
  covers slide body only — never captions, annotations, or notes.
- **One exhibit per slide, annotated.** Mark the decisive point on the figure
  itself. A deck where more than ~15% of content slides carry no exhibit has
  drifted back into being an outline.
- **Rebuild figures, don't screenshot them.** Paper figures carry print-sized
  fonts. Crop at 300 dpi with `scripts/figure.py`; re-plot from raw numbers
  when you have them, with axis labels 16pt or larger.
- **Bullet tiers**: `level: 0` is the bold header tier, `level: 1` the
  bulleted tier. Never fake indentation with spaces or dashes.
- **Speaker notes** carry what you would say — in 中文, technical terms in
  English, full sentences, three to six per content slide, no length limit.
  Terse slides are only safe when the notes are full.
- **Tables** (`conclusion`, `paper_list`, `summary`) take a list; unused rows
  are blanked. Leave the numbering column alone. A comparison table you are
  building yourself is a `matrix`, not one of these.
- Emoji status markers in the summary task table: 🔄 in progress · ✅ done ·
  ❌ dropped · ⏳ waiting · 🆕 new · 🔜 upcoming · 📌 important · 💤 on hold ·
  🛠️ needs fixing · 💡 idea · 📅 scheduled.

## Editing an existing deck instead

```bash
markitdown deck.pptx                                    # read content
python "$SKILL_DIR/scripts/thumbnail.py" deck.pptx deck-thumbs   # see layouts
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python "$SKILL_DIR/scripts/add_slide.py" unpacked/ slide2.xml --after slide2.xml
python "$SKILL_DIR/scripts/clean.py" unpacked/
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python "$SKILL_DIR/scripts/office/validate.py" out.pptx --original deck.pptx
```

Do all structural work — add, delete, reorder — **before** editing any slide's
content: `add_slide.py` copies verbatim, and `clean.py` deletes any slide
missing from `<p:sldIdLst>`.

If you script an XML transform, parse with `defusedxml.minidom`; round-tripping
OOXML through `xml.etree.ElementTree` rewrites namespace prefixes and corrupts
the deck.

## QA

`qa_check.py` exits non-zero on: a content slide with no claim line, body text
over the word cap, a borrowed figure with no source, text that will not fit its
box, a method or results slide with no exhibit, and an assertion-evidence slide
that carries bullets or no evidence. `--review` adds the claim sequence, a
slide inventory and a rubric for content, design and coherence — the three
things no mechanical check can see. Run it once before handing the deck over. Warnings — thin notes, a high text-only ratio, a claim that reads as a
label — are judgement calls; the reference decks themselves trip some of them.

Then look at every rendered page: text overflowing its box or the slide edge, a
title that wrapped and now collides with the subtitle, leftover `XXX` /
`Ur Name` / `20XX`, table rows that should have been blanked, a breadcrumb whose
bold segment does not match the slide's role.

One quirk is inherited, not yours: the `Todolist & Suggestion from Prof.` label
on the conclusion slides overflows its box in the template itself. Leave it.

## Dependencies

`python-pptx`, `markitdown[pptx]`, `Pillow`, `matplotlib`, `defusedxml`, `lxml` (pip) ·
LibreOffice via `scripts/office/soffice.py` · `pdftoppm` (Poppler, also used by
`figure.py`). MinerU is optional and never required.

If the environment cannot execute code, you are in the wrong package — use
`acm-slides-plain`, which produces the outline and stops there.
