# Paper study blueprint

For a 20–25 min paper presentation at group meeting. Read this together with
`lab-rules.md` (content standard) and `slide-patterns.md` (layout).

## Input contract

| Source | What it is for | Required |
|---|---|---|
| The original PDF | Understanding the argument. Read the PDF as pages and inspect the figures; a markdown conversion cannot show you what a teaser figure is doing. | **Yes** |
| MinerU output dir | Assets. `*_content_list.json` gives every image/table block a `page_idx` (0-based) and a `bbox` normalised to 0–1000, with caption and footnote already paired, and converts formulas to LaTeX. | No |
| Zotero export / annotations | Selection signal and metadata. The passages you highlighted are your own ranking of what matters; venue and year come from here. | No |

Never build from the MinerU markdown alone. Never crop a figure out of the
images MinerU wrote — those are downsampled bitmaps. Take MinerU's **bbox**
and re-render the original PDF at 300 dpi:

```bash
python3 scripts/figure.py list --mineru out/paper/auto
python3 scripts/figure.py crop --pdf paper.pdf --mineru out/paper/auto \
    --index 3 -o figs/fig3.png
```

Without MinerU, look at the page once, then crop:

```bash
python3 scripts/figure.py detect  --pdf paper.pdf --page 3 -o page3.png   # candidate boxes + grid
python3 scripts/figure.py crop --pdf paper.pdf --page 3 \
    --box 0.06,0.36,0.97,0.86 -o figs/fig1.png --trim --drop-caption
```

`detect` prints a `--box` per figure-like block and writes the page with a
0.1 grid and the boxes numbered; `preview` writes the grid alone. The first
decks built with this skill spent 11-19 crops per paper guessing boxes from
whole-page fractions. Read the box off the image instead. `--box` is corners,
`x0,y0,x1,y1`, the same form an annotation's `xyxy` takes.

MinerU earns its keep on one specific thing: composite figures. A CVPR teaser
is a dozen sub-images laid out as one figure; MinerU's layout detection keeps
it as a single block, while pulling embedded images straight out of the PDF
shatters it into fragments.

## Shape of the deck

Rebuild the paper as an argument. Do not walk its sections in order.
The allocation below is an example, not a required slide count. Spend time on
the mechanisms and evidence this paper needs; do not pad sections to match it.

| Section | Slides | Time | Role |
|---|---|---|---|
| Title + teaser figure | 2 | 1 min | `cover`, `paper_intro` |
| Task definition + motivation + gap | 4–5 | 4 min | `paper_intro` ×4 |
| Contributions | 1 | 1 min | `paper_intro` |
| Related work: the camps and what each lacks | 4–5 | 3 min | `paper_related` ×5 |
| Method: overview then one slide per component | 6–7 | 7 min | `paper_method` ×7 |
| Setup, quantitative, qualitative, ablation | 5–6 | 6 min | `paper_results` ×6 |
| Takeaways, critique, relevance | 1–2 | 2 min | `paper_conclusion` |
| **Our angle**: what we tried, what transfers, what we would do differently | 1 | 1 min | `paper_conclusion` with `"ours": true` |

Roles repeat — that is intended, the breadcrumb stays on the right section.
Measured against the two decks in `reference-decks.md`: 26 slides split
5 / 5 / 7 / 6 / 1 across intro, related, method, results, conclusion.

The last row is not optional. A deck that only relays the paper is the
failure the first decks built with this skill had: sixteen slides, all of
them the paper's, none of them ours. The meeting wants to know why *you* read
it — what you tried, what maps onto our project, where you disagree. Put that
on its own slide and mark it `"ours": true`; `qa_check.py` warns when a paper
talk has none.

## Which exhibit, per section

Tables are the default a text generator falls into, and seven of them in a
row is what the handbook means by "dense". The shapes in
`slide-patterns.md` §Graphic exhibits exist so each section has a better one:

| Section | Exhibit |
|---|---|
| Teaser | the paper's figure, `figure-full` or `assertion-evidence` |
| Motivation / gap | `cards` for the camps and what each lacks, or a `quadrant` of the field with the empty corner named |
| Contributions | `cards`, one per contribution, three at most |
| Related work | `cards` per camp with a `callout` naming the limitation; a `matrix` only for the one feature-by-method comparison |
| Method overview | `flow`, the stage of the next slide lit; then one slide per stage with the cropped figure on the right |
| Method component | cropped figure (`figure-right`, or `figure-bottom` if it is wide) with a `box` on the component, plus the four questions as bullets |
| Headline result | `bignum`, the winning number and its baseline, before the table |
| Quantitative | `matrix`, `highlight_row` on the paper, `highlight_col` on the decisive metric |
| Qualitative | `figure-full`, cropped to the three methods that matter |
| Ablation | `bignum` or a short `matrix`; the question it answers in the subtitle |
| Takeaways / our angle | `cards`: transfers / does not / next week, with `"ours": true` |

A whole talk in cards is as monotonous as a whole talk in tables; the QA
gate warns on three of the same shape in a row either way.

## Where the argument lives

Put the claim in the title, or in a subtitle when the title is a section label.
Use callouts sparingly at a turning point; they are not a required field.
Read titles and their claim-bearing subtitles in sequence to test the argument.

## Per-section requirements

**Method slides.** For each component answer four questions: what does it
receive, what does it do, what does it produce, why is it needed. Naming
modules without design logic is the failure mode named in the handbook. One
component per slide. Use the original architecture crop with a highlighted
path, or an explicitly adapted diagram showing inputs, transformations and
outputs. Use full width when the mechanism would be cramped in a half-slide.

**Evidence slides.** A quantitative slide answers a claim; an ablation slide
answers why the method works. Name the metric, the strongest relevant
baseline, the decisive values, and the exact conclusion supported. Choose the
visual from the question: a curve for convergence, a scatter plot for a
quality/cost tradeoff, paired image crops for visual quality, or a compact
`matrix` when exact cross-metric lookup matters. A readable crop of a paper
plot or table is valid; trim irrelevant rows or panels without losing labels,
conditions or caveats. Re-plot only from verified values. See
`evidence-selection.md` for extraction and provenance. Lead with the one number that decides it as a `bignum`, baseline named,
before the full table.

**Three statement types stay visibly separate**: author claim, evidence,
presenter interpretation. Label interpretations in the title, caption or body
where needed; this does not require a box. In the notes,
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

1. Read the original PDF and inspect its architecture, result and ablation
   figures before drafting the outline. Map selected panels to claims and
   retain page, figure number, caption and source.
2. Crop the selected evidence with `figure.py`; inspect the crops before
   choosing a layout. Use `evidence-selection.md` to decide what to preserve,
   annotate or redraw. Missing assets are not a reason to substitute prose tables.
3. Propose titles, claim-bearing subtitles where needed, and the actual exhibit
   for each slide. Follow `SKILL.md` for outline confirmation.
4. Build, compose and run the QA loop in `SKILL.md`. Inspect individual renders
   of dense figures; contact sheets alone cannot verify embedded axis labels.

## Never

- Invent a number, a metric, or an ablation result. If the paper does not
  state it, the slide does not claim it.
- Present a limitation as the author's when it is yours, or the reverse.
- Ship a slide whose figure you have not looked at in the render.
