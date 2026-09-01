---
name: acm-pptx
description: "Build ACM Lab (NYCU) presentations on the lab's official template — weekly progress reports, project/research updates, and paper-study talks. Use this skill whenever the user asks for a 進度報告, 週報, progress report, 組會投影片, lab presentation, paper presentation, or any .pptx/.potx that should follow ACM Lab format, and whenever a deck is being created, edited, or read for this lab. Also use it when the user hands you an outline and asks for slides, or mentions Prof. Huang Ching-Chun's lab meeting. Do not build ACM Lab slides from scratch with pptxgenjs — always clone the bundled template."
version: 1.1.0
template_version: template_final.pptx (22 slides, 13.333in x 7.5in)
license: Lab-internal use
---

# ACM Lab slide builder

The lab template is the only source of visual truth. Every slide in
`assets/acm_template.pptx` carries its own layout with the section colour bar
and the breadcrumb strip already positioned. **Never rebuild a slide — clone
the template slide whose role matches, then replace the text.**

## The one-command path

```bash
python scripts/build_from_outline.py outline.json -o report.pptx
```

`outline.json` names each slide by **role**; the builder clones that template
slide, replaces the text run-by-run (fonts, bullet tiers, table styling
survive), and writes the deck.

Full workflow:

```bash
python scripts/build_from_outline.py --roles          # see the role table
# write outline.json  (schema: references/outline-schema.md)
python scripts/build_from_outline.py outline.json -o report.pptx
python scripts/office/validate.py report.pptx --original assets/acm_template.pptx
markitdown report.pptx | grep -inE "\bx{3,}\b|lorem|ipsum|\[insert|Ur Name|20XX"
```

`--original` is required: the lab template has pre-existing XSD quirks, and
baselining against it keeps a real regression from hiding among them.

## Before you write the outline

1. **Read `references/lab-rules.md`.** It is the slide-design and structure
   standard distilled from the two ACM handbooks. It governs titles, density,
   what each section must contain, and timing.
2. **Read `references/template-map.md`** to pick roles. Do not guess role
   names — `--roles` prints the authoritative list.
3. **Ask before inventing content.** If the user has not said what they did
   this week, ask; do not fill the deck with plausible-sounding progress.
   Never fabricate experimental numbers, metrics, or paper results.

## Choosing roles

A weekly progress report is assembled from tracks. Include only the tracks
that have real content; for an empty track use its `*_summary` slide with
every term set to `NONE` — that is what the template intends.

| Situation | Roles |
|---|---|
| Weekly report, project only | `cover`, `summary`, `project_summary` … `project_conclusion` |
| Weekly report, research only | `cover`, `summary`, `research_summary` … `research_conclusion` |
| Both tracks | project block then research block; one `summary` at the front |
| Paper study session | `cover`, `paper_list`, `paper_intro`, `paper_related`, `paper_method`, `paper_results`, `paper_conclusion` |
| Anything else | `others` |

A role may repeat — ask for `project_results` twice and you get two of that
slide, breadcrumbs identical, which is correct for a multi-experiment week.

## Writing the content

- **Claim-based titles.** `Latency drops 3x once the verifier is cached`, not
  `Results`. The handbook names generic titles as a failure mode.
- **Ghost deck test before building.** Read the outline's titles in sequence,
  bodies ignored. They must carry the whole argument alone. If they read as a
  list of topics, fix the outline first — do not build and patch later.
- **Telegraphic wording.** No hedges or framing clauses: `Cost down 23%
  (p < 0.01)`, never `Our results demonstrate that costs were significantly
  reduced`. Phrases like *we propose*, *it is worth noting*, *this shows that*
  are the tell of machine-written slides. `references/lab-rules.md` §Writing
  discipline is the full standard — read it, it governs every line you write.
- **One idea per slide**, short phrases, no paragraph blocks, ~40 words of body
  text as a hard ceiling.
- **Bullet tiers**: `level: 0` is the bold header tier, `level: 1` the
  bulleted tier. The builder maps levels onto the template's own indent
  tiers, so do not try to fake indentation with spaces or `-`.
- **Speaker notes** (`notes`) carry what you would say — the argument, the
  numbers you will quote, the questions you expect. Never put them on the
  slide.
- **Tables** (`conclusion`, `paper_list`, `summary`) take a list; unused rows
  are blanked automatically. Leave the numbering column alone.
- Emoji status markers in the summary task table are part of the template's
  convention: 🔄 in progress · ✅ done · ❌ dropped · ⏳ waiting · 🆕 new ·
  🔜 upcoming · 📌 important · 💤 on hold · 🛠️ needs fixing · 💡 idea ·
  📅 scheduled.

## Editing an existing deck instead

When the user hands you a deck rather than an outline:

```bash
markitdown deck.pptx                                    # read content
python scripts/thumbnail.py deck.pptx deck-thumbs       # see layouts
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python scripts/add_slide.py unpacked/ slide2.xml --after slide2.xml
python scripts/clean.py unpacked/
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python scripts/office/validate.py out.pptx --original deck.pptx
```

Do all structural work — add, delete, reorder — **before** editing any slide's
content: `add_slide.py` copies verbatim, and `clean.py` deletes any slide
missing from `<p:sldIdLst>`.

If you script an XML transform, parse with `defusedxml.minidom`; round-tripping
OOXML through `xml.etree.ElementTree` rewrites namespace prefixes and corrupts
the deck.

## QA — required, every time

```bash
python scripts/office/validate.py report.pptx --original assets/acm_template.pptx
python scripts/office/soffice.py --headless --convert-to pdf report.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 110 report.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

View every rendered page. Look for, in this order: text overflowing its box or
the slide edge; a title that wrapped to two lines and now collides with the
body; leftover `XXX` / `Ur Name` / `20XX`; table rows that should have been
blanked; a breadcrumb whose bold segment does not match the slide's actual
role.

One quirk is inherited, not yours: the `Todolist & Suggestion from Prof.` label
on the conclusion slides overflows its box in the template itself. Leave it.

## Dependencies

`python-pptx`, `markitdown[pptx]`, `Pillow`, `defusedxml`, `lxml` (pip) ·
LibreOffice via `scripts/office/soffice.py` · `pdftoppm` (Poppler).

If the environment cannot execute code, you are in the wrong package — use
`acm-slides-plain`, which produces the outline and stops there.
