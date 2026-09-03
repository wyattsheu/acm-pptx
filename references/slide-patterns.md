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

  "equation": {"text": "M = F( [Mtex, Mgeo], Θ )", "label": "Eq. (1)"},

  "callout": {"label": "Key idea",
              "text": "Starting from a real image, not noise, is why one step suffices."},

  "notes": "中文, full sentences, 3–6 per content slide."
}
```

`callout` also accepts a bare string. `figure` also accepts a bare path, but
then it has no caption and no source, so `qa_check.py` will complain.

## Layouts

| `layout` | Body | Figure | Use for |
|---|---|---|---|
| `text-only` | full width | — | contributions, takeaways, a `matrix` slide |
| `figure-right` | left half, left-aligned | right half, vertically centred | one method component + its diagram |
| `figure-left` | right half | left 48% | when the figure reads left-to-right into the text |
| `figure-bottom` | top strip, ≤1.85in | full width below, top-aligned | task definition, teaser, wide pipeline figures |
| `figure-full` | cleared | whole content region | qualitative comparison grids |

Adding a `callout` automatically shortens the content region by 0.96in.
Adding a `subtitle` pushes it down 0.06in. You do not adjust for either.

Geometry, for reference only: title 0.92–2.42in wide band at y 0.33, subtitle
at y 1.18, content from y 1.72 to 6.80, callout occupies 6.02–6.80, figures
may bleed to x 0.62–12.72 while text stays inside 0.92–12.42.

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
