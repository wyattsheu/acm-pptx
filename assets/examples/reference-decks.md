# Reference decks

Two decks presented at this lab that set the quality bar. Measured, not
remembered — the numbers below come from parsing the files.

## A. ELITE (CVPR 2026 head-avatar paper), 26 slides

| Section | Slides | Breadcrumb |
|---|---|---|
| Cover + paper info & teaser | 1–2 | — |
| Task definition, motivation ×3, contributions | 3–7 | Introduction |
| Taxonomy figure, three camps, the gap table | 8–12 | Related work |
| Overview, MGPM ×2, why adapt, stage 1, enhancer, stage 2 | 13–19 | Proposed Method |
| Setup, qualitative, comparison, quantitative, identity, ablation | 20–25 | Experimental Results |
| Conclusion + takeaways | 26 | Conclusion |

Only **2 of 26** slides carry no exhibit (contributions, conclusion).

What it does that a text generator does not:

- Every content slide has a **red claim line** under a plain section title.
  `Two families of prior — two opposite failure modes`;
  `Nobody has both good initialization and good supervision`;
  `Learned initialization is a good start — but not the answer`.
- Most slides end in a **red-outlined box** with the one sentence to remember:
  *That question — not the 3D representation — is what separates all the prior
  work in this area.*
- Slides 9, 10, 11 **reuse one taxonomy figure** and move a red dashed
  rectangle onto a different column each time.
- Hand-built **box-and-arrow diagrams** (slides 5, 6, 13) instead of screenshots.
- Comparison tables built in PowerPoint with the paper's own row highlighted
  green (slides 12, 20, 23).
- Figure captions in italic grey, numbered `Fig. N`, referenced from bullets.
- Equations on a grey band with `Eq. (1)` underneath.

## B. TADSR (one-step diffusion Real-ISR), 17 slides

Per-slide measurements:

| | pictures | own annotation boxes | figure area | body words | notes |
|---|---|---|---|---|---|
| s2 motivation | 0 | 2 | 0% | 40 | 394 ch |
| s4 motivation | 1 | 6 | 18% | 90 | 281 ch |
| s6 TAE | 3 | 2 | 35% | 54 | 343 ch |
| s7 TAVSD | 4 | 4 | 45% | 63 | — |
| s10–13 qualitative | 1 | 2 | 51% | 3 | — |
| s14–15 ablation | 3 | 2 | 68–72% | 3 | — |
| s16 conclusion | 0 | 13 | 0% | 122 | 372 ch |

12 of 17 slides carry figures. The text boxes beside the figures are the
presenter's own annotations — *only an increase in sharpness*, *recovers a more
realistic parrot image*, *ts larger → kernel size larger → more blur* — not
anything the paper wrote.

## What the reference decks tell us about the rules

- Section-label titles are **fine** here. Both decks use `Ablation Study`,
  `Quantitative Comparisons`, `Task Definition`. The claim lives in the red
  subtitle and the callout. Judge the ghost deck on those.
- The 40-word body cap is real but is measured on **body text only**. TADSR
  s4 and s16 exceed it and are the two most crowded slides in that deck.
- Speaker notes are full on conceptual slides and **empty on results slides**
  in TADSR. Thin notes are a warning worth looking at, not an error.
- Text-only slides run 8% (ELITE) to 12% (TADSR). A deck above ~15% is drifting
  back into an outline.
