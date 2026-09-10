# Slide patterns

`build_from_outline.py` fills text into a cloned template slide.
`compose.py` runs after it and adds everything else. These are the fields it
reads and the layouts it computes. **Never write coordinates into the
outline** — name a layout and the regions are derived, which is how overlap
is avoided.

## Fields added to a slide entry

```jsonc
{
  "role": "paper_method",
  "title": "Proposed Method",              // section label, from the template
  "subtitle": "MGPM maps UV maps plus driving signals to 2D Gaussians",
  "layout": "figure-right",                // see table below
  "bullets": [ … ],                        // as before, <= 40 words total

  "figure": {
    "src": "figs/fig3.png",
    "caption": "Fig. 3  MGPM takes mesh UV maps and driving signals …",
    "source": "Kim et al., CVPR 2026"      // required for borrowed figures
  },

  "annotations": [                          // coordinates are FRACTIONS OF THE
    {"type": "box",   "at": [0.02, 0.02, 0.30, 0.34],                    // PICTURE,
     "color": "C00000", "dash": true},                                    // so they
    {"type": "arrow", "at": [0.10, 0.50, 0.40, 0.50]},                    // survive
    {"type": "label", "at": [0.02, 0.38], "text": "334 IDs x 16 views"}   // resizing
  ],

  "matrix": {                               // a comparison table you build,
    "header": ["Method", "Duration", "PSNR", "CSIM"],   // not a screenshot
    "rows": [["CAP4D", "400 min", "19.478", "0.7064"],
             ["ELITE (Ours)", "20 min", "25.220", "0.7396"]],
    "highlight_row": 2                      // 1-based, the row being argued for
  },

  // QUOTE an equation -> crop it, keep the paper's notation
  "equation": {
    "src": "figs/eq1.png", "label": "Eq. (1)",
    "where": ["Mtex, Mgeo: canonical FLAME UV texture and geometry maps"],
    "annotations": [{"type":"label","at":[0.55,-0.2],"text":"one forward pass"}]
  },

  // or EXPLAIN one -> rebuild it, one colour per role
  "equation": {
    "parts": [["\\mathcal{F}_{\\varphi}"], ["\\rightarrow"],
              ["\\mathcal{F}_{\\varphi}^{*}", "D2691E"]],
    "captions": {"D2691E": "adapted on 3 real frames - identity-specific"}
  },

  "stage": "densify",     // needs a top-level `stage_figure`, see below
  "divider": true,        // thin rule between the text and figure columns

  "callout": {"label": "Key idea",
              "text": "Starting from a real image, not noise, is why one step suffices."},

  "notes": "中文, full sentences, 3–6 per content slide."
}
```

`callout` also accepts a bare string. `figure` also accepts a bare path, but
then it has no caption and no source, so `qa_check.py` will complain.

## Quote an equation, or rebuild it

Both reference decks do both, and the choice follows what the slide is for.

- **The slide quotes the paper** ("this is what they wrote") -> `src`. Crop it
  with `figure.py`, list the symbols in `where`, and annotate from outside.
  Boxes, arrows and labels all work on a crop; one deck underlines two terms of
  a cropped formula and labels them *prediction* and *actual observations*.
  Cropping also keeps the notation exactly as the paper has it, which is what
  you want when someone asks a question about it.
- **The slide explains a decomposition** ("this matrix does two jobs") ->
  `parts`. Here the colour carries the argument -- one factor is the learned
  action, the other the computed decision -- and you cannot recolour half of a
  screenshot. Rendering is matplotlib mathtext, so no TeX install is needed,
  but it covers only a subset of LaTeX: multi-line aligned derivations do not
  lay out well, so crop those.

A paper talk usually needs one or two rebuilt equations, in Proposed Method.
Everything in Related Work gets cropped. If you rebuild an equation, say so in
the notes ("記號已改寫") so a question about the original does not catch you out.

## The recurring pipeline strip

Name the overview figure once, at the top level of the outline, with a box for
each stage:

```jsonc
"stage_figure": {
  "src": "figs/overview.png",
  "caption": "Fig. 2  Pipeline; the boxed stage is the one on this slide.",
  "stages": {"match": [0.02, 0.05, 0.20, 0.90], "densify": [0.24, 0.05, 0.34, 0.90]}
}
```

Then each method slide says `"stage": "densify"` and nothing else -- the strip
is placed top-right, the box moves to that stage, and the body text takes the
left column. Both reference decks repeat their overview figure on every method
slide this way; it is the cheapest progressive disclosure available, and naming
it once keeps it out of every slide entry.

## Layouts

| `layout` | Body | Figure | Use for |
|---|---|---|---|
| `text-only` | full width | — | contributions, takeaways, a `matrix` slide |
| `figure-right` | left half, left-aligned | right half, vertically centred | one method component + its diagram |
| `figure-left` | right half | left 48% | when the figure reads left-to-right into the text |
| `figure-bottom` | top strip, ≤1.85in | full width below, top-aligned | task definition, teaser, wide pipeline figures |
| `figure-full` | cleared | whole content region | qualitative comparison grids |
| `assertion-evidence` | forbidden | whole content region | a slide that makes exactly one point |

Adding a `callout` automatically shortens the content region by 0.96in.
Adding a `subtitle` pushes it down 0.06in. You do not adjust for either.

Geometry, for reference only: title 0.92–2.42in wide band at y 0.33, subtitle
at y 1.18, content from y 1.72 to 6.80, callout occupies 6.02–6.80, figures
may bleed to x 0.62–12.72 while text stays inside 0.92–12.42.

## Assertion-evidence

`"layout": "assertion-evidence"` sets the claim in `subtitle` at reading size
where the small red line would go, fills the rest with the exhibit, and takes
no bullets at all -- you say the supporting sentences out loud instead, which
is what the notes are for.

This is the one slide structure with a controlled experiment behind it. Two
audiences heard identical words over different slides; the group that saw a
sentence headline plus visual evidence understood and remembered more than the
group that saw a topic headline plus bullets, significantly so, with fewer
misconceptions and lower reported cognitive load. A later study found the
presenter understands their own material better too, because the structure
forces you to decide the point of each slide before you build it.

Use it where a slide really does make one point: the teaser, a qualitative
comparison, the single result the talk turns on. Do not use it everywhere --
in a paper study the audience needs the four method questions written down,
and those are bullets. Two or three assertion-evidence slides in a 25-slide
talk is about right.

`qa_check.py` rejects this layout with bullets, without an exhibit, or without
a `subtitle` to serve as the assertion.

## What each slide type must carry

The machine-checkable half of this file. `qa_check.py` applies it by role:

| role | required | severity |
|---|---|---|
| `paper_method`, `paper_results` | an exhibit (figure, matrix, equation or stage) | error |
| `project_results`, `research_results` | an exhibit | error |
| `paper_related` | an exhibit, and a `callout` naming the limitation | warning |
| `paper_intro` | an exhibit | warning |
| `paper_conclusion` | `bullets` (the takeaways) | warning |
| every content slide | a claim in `subtitle` or `callout` | error |

## Reviewing the deck

Mechanical checks cannot see whether a claim follows from the one before it,
and that is the dimension where automatic decks are weakest. Before you hand a
deck over:

```bash
python "$SKILL_DIR/scripts/qa_check.py" outline.json talk.pptx --review
```

It prints the claim sequence on its own -- read it as one paragraph, the way
consultants read a deck by its titles alone -- plus a per-slide inventory and
a rubric covering content, design and coherence. Judge those three yourself
after looking at the render; the script only lays the evidence out.

## When to use which exhibit

- **Comparison across methods or camps** → `matrix`, `highlight_row` on the
  row you are arguing for. Not a figure of the paper's table.
- **A pipeline or architecture** → cropped figure, `figure-right`, with a
  `box` annotation on the component this slide is about.
- **The same figure across several slides** → reuse the same `src` and move
  the annotation. Both reference decks do this: one taxonomy figure appears on
  four consecutive slides with a red dashed box on a different column each
  time. It is the cheapest form of progressive disclosure available.
- **A loss or a definition** → `equation`, with the symbols expanded in
  bullets beside or below it. One equation per slide.
- **Qualitative results** → `figure-full`, and crop hard. A grid of eight
  methods is unreadable projected; crop to the three that matter.

## Annotation conventions

Red `C00000` for the thing under discussion, dashed for "this region of the
figure", solid for "this value". A label sits outside the picture edge where
possible. If an annotation needs more than four words it is a bullet, not an
annotation.
