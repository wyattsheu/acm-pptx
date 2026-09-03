# Paper study blueprint

For a 20–25 min paper presentation at group meeting. Read this together with
`lab-rules.md` (content standard) and `slide-patterns.md` (layout).

## Input contract

| Source | What it is for | Required |
|---|---|---|
| The original PDF | Understanding the argument. Claude reads a PDF as pages and can *see* the figures; a markdown conversion cannot show you what a teaser figure is doing. | **Yes** |
| MinerU output dir | Assets. `*_content_list.json` gives every image/table block a `page_idx` (0-based) and a `bbox` normalised to 0–1000, with caption and footnote already paired, and converts formulas to LaTeX. | No |
| Zotero export / annotations | Selection signal and metadata. The passages you highlighted are your own ranking of what matters; venue and year come from here. | No |

Never build from the MinerU markdown alone. Never crop a figure out of the
images MinerU wrote — those are downsampled bitmaps. Take MinerU's **bbox**
and re-render the original PDF at 300 dpi:

```bash
python scripts/figure.py list --mineru out/paper/auto
python scripts/figure.py crop --pdf paper.pdf --mineru out/paper/auto \
    --index 3 -o figs/fig3.png
```

Without MinerU, crop by page fraction and expect to iterate once or twice:

```bash
python scripts/figure.py crop --pdf paper.pdf --page 3 \
    --box 0.06,0.36,0.97,0.86 -o figs/fig1.png
```

MinerU earns its keep on one specific thing: composite figures. A CVPR teaser
is a dozen sub-images laid out as one figure; MinerU's layout detection keeps
it as a single block, while pulling embedded images straight out of the PDF
shatters it into fragments.

## Shape of the deck

Rebuild the paper as an argument. Do not walk its sections in order.

| Section | Slides | Time | Role |
|---|---|---|---|
| Title + teaser figure | 2 | 1 min | `cover`, `paper_intro` |
| Task definition + motivation + gap | 4–5 | 4 min | `paper_intro` ×4 |
| Contributions | 1 | 1 min | `paper_intro` |
| Related work: the camps and what each lacks | 4–5 | 3 min | `paper_related` ×5 |
| Method: overview then one slide per component | 6–7 | 7 min | `paper_method` ×7 |
| Setup, quantitative, qualitative, ablation | 5–6 | 6 min | `paper_results` ×6 |
| Takeaways, critique, relevance | 1–2 | 2 min | `paper_conclusion` |

Roles repeat — that is intended, the breadcrumb stays on the right section.
Measured against the two decks in `reference-decks.md`: 26 slides split
5 / 5 / 7 / 6 / 1 across intro, related, method, results, conclusion.

## Where the argument lives

The template's titles are section labels (`Proposed Method`,
`Experimental Results`) and both reference decks leave them that way. The
claim goes in two other places, on every content slide:

- **`subtitle`** — the red line under the title. One statement.
  `Nobody has both good initialization and good supervision`, not
  `Comparison of prior work`.
- **`callout`** — the boxed line at the bottom. The "so what" the audience
  should leave with. `Key idea:`, `Limitation:`, `Evidence:` as the label.

Run the ghost deck test on the **subtitles**, not the titles. Read them in
sequence; they must carry the whole argument alone.

## Per-section requirements

**Method slides.** For each component answer four questions: what does it
receive, what does it do, what does it produce, why is it needed. Naming
modules without design logic is the failure mode named in the handbook. One
component per slide, its figure on the right.

**Evidence slides.** A quantitative slide answers a claim; an ablation slide
answers why the method works. Name the metric, the strongest relevant
baseline, the decisive values, and the exact conclusion supported. Put the
numbers in a `matrix` with `highlight_row` on the paper's own method — never
paste a screenshot of the paper's table, its font is print-sized.

**Three statement types stay visibly separate**: author claim, evidence,
presenter interpretation. In the callout, label which one it is. In the notes,
say "作者宣稱…" versus "我的看法是…".

**The ending.** Three takeaways, strengths, weaknesses, limitations split into
author-acknowledged and presenter-identified, relevance to the lab, next
directions. Keep this slide up during Q&A. Ending on a bare "Questions?" is a
failure mode.

## Speaker notes

中文, full sentences, three to six per content slide, no length limit. Write
them as speech. Each note covers, as applicable: how this slide follows from
the previous one; what the numbers mean and against what baseline; why this
design choice and not the obvious alternative; what is still uncertain.

Add, on method and results slides, **the question the advisor is likely to ask
and your answer**. That is what the meeting is actually for.

## Workflow

```bash
# 1. read the PDF, propose the outline — titles, subtitles, exhibit per slide
#    STOP HERE and get it confirmed. Any deck over 10 slides gets confirmed
#    before a single slide is built.

# 2. figures
python scripts/figure.py list --mineru out/paper/auto
python scripts/figure.py crop --pdf paper.pdf --mineru out/paper/auto --index 3 -o figs/fig3.png

# 3. build
python scripts/build_from_outline.py outline.json -o talk.pptx
python scripts/compose.py outline.json talk.pptx

# 4. gate
python scripts/qa_check.py outline.json talk.pptx
python scripts/office/validate.py talk.pptx --original assets/acm_template.pptx

# 5. look at it
python scripts/office/soffice.py --headless --convert-to pdf talk.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 150 talk.pdf slide
```

Render at 150 dpi or higher. At 100 dpi JPEG artifacts look like stray shapes
and you will chase bugs that are not there.

## Never

- Invent a number, a metric, or an ablation result. If the paper does not
  state it, the slide does not claim it.
- Present a limitation as the author's when it is yours, or the reverse.
- Ship a slide whose figure you have not looked at in the render.
