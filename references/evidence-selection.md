# Choose evidence before layout

Start with what the audience needs to understand and the artifacts available.
Keep a small working map of claim → source file/page/run → relevant panel or
measurement → visual form. Put provenance in captions and notes; this is not
an additional deliverable or outline schema. Inspect sources before writing
the slide around them. If evidence is missing, narrow the claim, identify the
missing result, or ask for the needed artifact; do not invent it.

## Explain a mechanism

Show inputs, transformations and outputs with connected nodes, spatial groups
or a sequence. Label arrows with what moves between stages. For a change to a
system, highlight the changed module and show the consequence downstream.
A component/purpose/status table loses these relationships.

For a paper, start from the original overview figure. Crop and enlarge it,
highlight the relevant path, then reuse it with a different highlight when
explaining another stage. If simplifying helps, redraw only the relevant nodes
and connections with the native `diagram` field; caption it “Adapted from Fig.
N” and preserve the original semantics in the notes. Do not flatten a newly
authored workflow into PNG, SVG or a Mermaid screenshot.

Use semantic shapes sparingly: rectangle = process, diamond = decision,
terminator = start/end, cylinder = stored data. Prefer a single neutral fill,
thin connectors, aligned baselines and whitespace. Accent only the current
stage. Rounded pastel cards, gradients, shadows, ornamental icons and a unique
colour for every node make generated slides look generic and weaken topology.

Use a `figure` instead when the original pixels carry evidence: a paper's full
architecture drawing, a photograph, microscopy, a screenshot, a qualitative
result or a dense composite whose faithful native reconstruction would invent
detail. Native labels and callouts may still sit outside that image.

## Explain an execution result

Inspect the actual logs, CSV files, output images or videos. Separate what
ran from what was observed and what you infer. Preserve run/configuration,
dataset or scene, units, metric direction and comparable baseline in the
caption or notes. Report unavailable results as unavailable.

- Convergence or change over time: plot the measured series with labelled
  axes. Mark the event or interval relevant to the claim.
- A quality/speed tradeoff: plot the measured methods on common axes; label
  points and preserve measurement conditions.
- Visual improvement or failure: align input, baseline and new output using
  the same sample and crop. Mark the region that demonstrates the difference;
  keep a failure case when it affects the conclusion.
- Motion or execution order: use a short embedded clip, or labelled frames
  when stills explain the claim. A “demo completed” table is not a demo.
- Exact values across several metrics: a small `matrix` is useful. Keep the
  decisive rows/columns, units, conditions and source; avoid prose paragraphs
  inside cells. Let label columns be wider than short metric columns, avoid
  decorative zebra striping, and highlight only the row or value under
  discussion. Do not turn a single observed change into a table by default.

Create plots with matplotlib from the verified values, then use `figure.src`.
For aligned images, compose a labelled panel from the actual outputs without
altering their content. This is a scientific result figure, not generative
imagery. Save assets referenced by the outline where rebuilds can find them.

## Quote a paper figure

Read the original PDF page, caption and surrounding discussion. Locate the
architecture, qualitative comparison, quantitative plot and relevant ablation
before deciding how many slides they need. MinerU can locate panels; it does
not replace looking at the original figures.

Crop the original PDF using `figure.py` (commands in `blueprint-paper.md`).
Retain axes, legends, panel letters, scale bars and baseline labels needed for
the claim. Keep the citation in `figure.source`, the figure/panel number in
`figure.caption`, and the PDF page plus any crop/adaptation detail in notes.
For a shared `stage_figure`, include the citation in its `caption` (the stage
renderer does not display a separate `source` field).

Inspect the crop and its rendered slide. High DPI cannot fix tiny labels:
enlarge a panel, use another slide, or add a faithful explanatory annotation.
Original plots and compact table crops are valid when readable. Re-plot only
from verified values, preserving scales, units and uncertainty. Do not guess
data points from pixels or redraw qualitative outputs as if they were results.

## Review the choices

Template summary and to-do tables serve their administrative purpose. For
content slides, check whether the visual actually explains the claim. Several
tables in succession warrant review, not an automatic rejection: do their
rows and columns enable needed comparison, or merely repackage the narration?
Do not impose table/figure percentages or manufacture diagrams for variety.
Check that a method can be followed and a result can be seen, not just that
every slide has a populated exhibit field.
