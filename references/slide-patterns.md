# Slide patterns

`build_from_outline.py` fills text into a cloned template slide.
`compose.py` runs after it and adds everything else. These are the fields it
reads and the layouts it computes. **Never write coordinates into the
outline** — name a layout and the regions are derived, which is how overlap
is avoided.

New workflows, pipelines and architecture diagrams use native editable
PowerPoint shapes through `diagram`. Plots and source figures remain assets
passed through `figure`. The fields below show API options, not a checklist to
fill on every slide.

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

  // or a VIDEO, in the same slot -- a figure and a video cannot share a slide
  "video": {
    "src": "figs/demo.mp4",                 // H.264/AAC mp4; see below
    "caption": "Demo  tracking at 30 fps, uncut",
    "source": "ours",                       // same rule as a figure
    "poster": "figs/demo-still.png",        // optional; a frame is pulled if absent
    "poster_at": "0:04",                    // which frame, if you pull one
    "autoplay": true,                       // default: play on click
    "loop": true,
    "mute": false, "volume": 60,            // volume is a percentage
    "badge": true                           // the "▶ 0:12" marker, on by default
  },

  "annotations": [                          // fractions of the PICTURE, so they
    {"type": "box",   "at": [0.02, 0.02, 0.30, 0.34],         // survive resizing;
     "color": "C00000", "dash": true},                         // at = [x, y, w, h]
    {"type": "box",   "xyxy": [0.02, 0.02, 0.32, 0.36]},       // same box as corners
    {"type": "arrow", "at": [0.10, 0.50, 0.40, 0.50]},         // arrow: x0, y0, x1, y1
    {"type": "label", "at": [0.02, 0.38], "text": "334 IDs x 16 views"}
  ],                                        // see §Coordinates before writing any

  // GRAPHIC exhibits -- one per slide, in the figure slot; see §Graphic exhibits
  "cards":    {"items": [{"title": "...", "bullets": ["..."], "highlight": true}]},
  "flow":     {"steps": ["Track", {"label": "Prior", "sub": "one pass"}], "highlight": 2},
  "bignum":   {"items": [{"value": "20 min", "label": "fitting", "sub": "vs 400"}]},
  "quadrant": {"x": ["slow", "fast"], "y": ["low", "high"],
               "points": [{"label": "Ours", "at": [0.8, 0.85], "highlight": true}]},
  "draw":     "design.py:timeline",         // your own function; see §Your own drawing

  "ours": true,                             // this slide is the presenter's, not the paper's

  "diagram": {                              // OR a native editable flowchart
    "direction": "LR",                     // LR, RL, TB or BT
    "nodes": [
      {"id": "input", "text": "Sensor frame", "kind": "terminator"},
      {"id": "filter", "text": "Filter", "kind": "process"},
      {"id": "valid", "text": "Valid?", "kind": "decision"},
      {"id": "plan", "text": "Plan", "kind": "process", "accent": true}
    ],
    "edges": [
      {"from": "input", "to": "filter"},
      {"from": "filter", "to": "valid"},
      {"from": "valid", "to": "plan", "label": "yes"}
    ],
    "groups": [
      {"id": "perception", "label": "Perception", "nodes": ["input", "filter"]}
    ],
    "caption": "Editable system path; red marks the component discussed here."
  },

  "matrix": {                               // a comparison table you build,
    "header": ["Method", "Duration", "PSNR", "CSIM"],   // not a screenshot
    "rows": [["CAP4D", "400 min", "19.478", "0.7064"],
             ["ELITE (Ours)", "20 min", "25.220", "0.7396"]],
    "highlight_row": 2,                     // 1-based, the row being argued for (or a list)
    "highlight_col": 4,                     // 1-based, the decisive metric
    "col_widths": [2, 1, 1, 1],             // relative; default gives col 1 extra room
    "caption": "Self re-enactment on INSTA; higher is better"
  },                                        // type scales to the rows: 18pt for four

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

`callout` also accepts a bare string. `figure` and `video` also accept a bare
path, but then they have no caption and no source, so `qa_check.py` will
complain.

## Native editable diagrams

Use `diagram` for a workflow, dependency path, state transition or simplified
architecture that you are authoring. Every node, connector, edge label and
group frame becomes a separate PowerPoint object, named in the Selection Pane
as `diagram:node:*`, `diagram:edge:*`, `diagram:label:*` or `diagram:group:*`.
The user can edit text, recolour shapes and reroute arrows without rebuilding.

Supported node kinds are `process`, `rect`, `rounded`, `decision`,
`terminator`, `circle`, and `database`. Use them semantically, not for variety.
`accent: true` gives the one focal node the lab red outline. Optional `fill`,
`line`, `text_color`, `size`, and `bold` override a node only when the content
requires it. Edges accept `label`, `dashed`, `arrow`, `color`, and `width`.
Their endpoints attach to the nodes, so moving a node in PowerPoint carries the
connector with it; set `"attached": false` only when reproducing an unusual
route that must stay fixed.

Automatic layout handles small diagrams in LR/RL/TB/BT directions. Keep a
single slide to roughly 4–8 nodes; 20 nodes and 32 edges are hard limits, not
targets. When one node needs adjustment, add normalized `"at": [x,y,w,h]`
coordinates (0–1 inside the diagram region) to that node; avoid manually
positioning every node unless reproducing a source exactly.

Style is intentionally restrained: square process boxes, one neutral fill,
thin grey arrows, no shadows or gradients, and one red focus at most. This
matches the lab decks' hand-built box-and-arrow figures and avoids the generic
generated look of pastel cards, ornamental icons and equal emphasis everywhere.
Use `figure` instead for paper originals, plots, photographs, screenshots and
experimental imagery. Never render a new Mermaid/Graphviz/matplotlib workflow
to a bitmap merely to place it on the slide.

## Video

A video occupies the figure slot and obeys the same layouts. With `bullets` it
defaults to `figure-right`; without them, to `figure-full`.

**Prepare the file before you name it.** PowerPoint decodes H.264 video with
AAC audio in an `.mp4` or `.mov` and nothing else. Screen recorders and paper
supplementary sites hand you HEVC, VP9, `.mkv` and `.webm`, and all of those
open as a black rectangle in the meeting — with no error when the deck is
built. So:

```bash
python "$SKILL_DIR/scripts/video.py" probe capture.mov          # is it playable?
python "$SKILL_DIR/scripts/video.py" prep capture.mov -o figs/demo.mp4 --clip 0:03-0:18
```

`prep` re-encodes to H.264 High / yuv420p / AAC, caps the width at 1920, moves
the index to the front, and trims. `compose.py` refuses to embed a file that
would not play and prints that exact command, so this cannot reach a meeting
by accident.

Three more things follow from the format, not from taste:

- **Trim to fifteen seconds.** Nobody watches a minute of demo in a lab
  meeting, and every second is embedded in the file you email.
- **Autoplay the one clip the slide is about**, and loop it — you keep talking
  instead of hunting for the play button. Leave everything else click-to-play.
- **The deck must be presented from PowerPoint.** Export it to PDF and the
  video is gone. Say so when you hand the file over.

A poster frame is pulled automatically (10% in, past the fade-up) into a temp
dir; name `poster` yourself when the automatic frame is uninformative. There is
no play button in a static render, which is why `compose.py` draws the small
`▶ 0:12` badge in the corner.

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

| `layout` | Body | Exhibit | Use for |
|---|---|---|---|
| `text-only` | full width | — | contributions, takeaways, a `matrix` slide |
| `figure-right` | left half, left-aligned | right half, vertically centred | one method component + its diagram |
| `figure-left` | right half | left 48% | when the figure reads left-to-right into the text |
| `figure-bottom` | top strip, ≤1.85in | full width below, top-aligned | task definition, teaser, wide pipeline figures |
| `figure-full` | cleared | whole content region | qualitative comparison grids |
| `assertion-evidence` | forbidden | whole content region | a slide that makes exactly one point |

`video`, `diagram`, `matrix` and every graphic exhibit are placed by the
same table. When `layout` is omitted: a figure goes `figure-bottom`; a
video, a diagram or a quadrant goes `figure-right` beside the bullets; a
matrix, cards, flow or bignum goes `figure-bottom` under them; any of them
takes `figure-full` when there are no bullets. `text-only` with bullets
*and* a graphic is turned into `figure-bottom` rather than letting them
overlap.

Adding a `callout` automatically shortens the content region by 0.96in.
Adding a `subtitle` pushes it down 0.06in. You do not adjust for either.

Geometry, for reference only: title 0.92–2.42in wide band at y 0.33, subtitle
at y 1.18, content from y 1.72 to 6.80, callout occupies 6.02–6.80, figures
may bleed to x 0.62–12.72 while text stays inside 0.92–12.42.

## Graphic exhibits

Before v2.4 a slide could carry bullets, a table or one picture, so a
pipeline became a table and a comparison of camps became a table, and a
sixteen-slide paper talk came out as seven tables and two bullet lists. These
four shapes cover what those tables were standing in for. Each one takes the
exhibit slot, is drawn in the section colour of the slide's role, and lights
one element in red: the stage this slide is about, the camp the paper joins,
the number that wins.

| exhibit | what it is | use it for |
|---|---|---|
| `cards` | 2-6 panels, each a heading plus a few lines | the camps in related work; contributions; the four questions about one component; strengths vs weaknesses; what transfers to us vs what does not |
| `flow` | boxes joined by arrows, one lit | the method overview before the per-component slides; any pipeline you would otherwise type into a table; `"style": "chevron"` for a terser strip, `"direction": "column"` for a vertical one; a linear sequence only, `diagram` (below) handles branches, feedback loops and groups |
| `bignum` | 1-4 numbers at display size with a label and the baseline | the single result the talk turns on, *before* the full table; each `sub` names what the number beats |
| `quadrant` | two labelled axes, prior work as dots, ours in red | where the paper sits in the field; place by argument (`at` is a fraction of the plot, `[0,0]` bottom-left) |

```jsonc
"cards": {
  "items": [
    {"title": "Per-identity overfitting", "tag": "FlashAvatar, SplattingAvatar",
     "bullets": ["10-400 min per subject", "drifts off identity"]},
    {"title": "2D diffusion prior", "tag": "image enhancers", "highlight": true,
     "text": "Photoreal supervision, no 3D consistency on its own",
     "foot": "the piece ELITE borrows"}
  ],
  "columns": 3,              // optional; up to 4 fit one row, 5-6 wrap to two
  "style": "header"          // "header" (coloured bar) | "panel" (tint) | "outline"
}

"flow": {
  "steps": [{"label": "FLAME tracking", "sub": "UV maps"}, "MGPM",
            {"label": "Enhancer", "sub": "diffusion", "highlight": true}],
  "highlight": 2,            // alternative to a per-step flag; 1-based
  "style": "box",            // or "chevron"
  "direction": "row",        // or "column"
  "caption": "Fig. 2, redrawn"
}

"bignum": {"items": [
  {"value": "20 min", "label": "per-identity fitting", "sub": "CAP4D: 400 min",
   "highlight": true},
  {"value": "0.740", "label": "CSIM", "sub": "CAP4D: 0.706"}
]}

"quadrant": {
  "x": ["slow", "fast"], "y": ["low identity", "high identity"],
  "quadrants": ["", "goal", "", ""],          // corner labels TL, TR, BL, BR
  "points": [{"label": "CAP4D", "at": [0.15, 0.80]},
             {"label": "ELITE (ours)", "at": [0.80, 0.85], "highlight": true}]
}
```

Limits are deliberate: six cards, seven flow steps, four numbers. Past them
the slide is a figure, and the paper probably already drew it.

What stays a `matrix`: the one comparison where the cells *are* the argument,
with `highlight_row` on the paper's method and `highlight_col` on the decisive
metric. A deck where more than a third of the content slides are tables, or
three tables in a row, gets a warning from `qa_check.py`; so does a run of
three text-only slides.

## Your own drawing

When none of the shapes above is the right one, name a function:

```jsonc
{"role": "paper_intro", "title": "...", "subtitle": "...",
 "draw": "design.py:timeline",
 "timeline": {"events": [["2020", "NeRF"], ["2023", "3DGS"], ["2026", "ELITE"]]}}
```

`compose.py` imports `design.py` (relative to the outline) and calls
`timeline(slide, box, spec, kit)`. `box` is the exhibit region the layout left
free, in inches, so anything drawn inside it clears the title, subtitle,
bullets and callout. `spec` is the slide's own entry, so the function reads
its own fields from it. `kit` carries the helpers this skill draws with
(`rect`, `write`, `arrow`, `tint`, the palette, `kit.accent` for the section
colour) so the function needs no imports and matches the template.
`assets/examples/design.py` is the reference implementation: copy it beside
the outline and edit. `qa_check.py` counts a `draw` slide as carrying an
exhibit, and errors if the file or function is missing.

## Coordinates

Three places take a box, and before v2.4 they did not agree, which cost two
wrong red boxes on one figure. Now:

| where | form | meaning |
|---|---|---|
| `annotations[].at` (box, circle) | `[x, y, w, h]` | corner plus size, fractions of the placed picture |
| `annotations[].xyxy` | `[x0, y0, x1, y1]` | the same box as two corners |
| `annotations[].at` (arrow) | `[x0, y0, x1, y1]` | tail, then head |
| `annotations[].at` (label) | `[x, y]` | top-left of the text |
| `stage_figure.stages.*` | `[x, y, w, h]` or `{"xyxy": [...]}` | same rule as a box |
| `quadrant.points[].at` | `[x, y]` | fraction of the plot, origin bottom-left |
| `figure.py crop --box` | `x0,y0,x1,y1` | fractions of the PDF page |

If you have corners (from `figure.py`, from MinerU's bbox, from eyeballing a
render), write `xyxy` and stop converting. A box whose `x + w` or `y + h`
passes the picture edge is refused by both `qa_check.py` and `compose.py`
with this table's name in the message, because that is what corners written
into `at` look like.

## Marking what is ours

`"ours": true` on a slide draws a small red `OUR TAKE` tag at the bottom left;
a string (`"ours": "OUR EXPERIMENT"`, `"ours": "我們的實驗"`) draws that
instead. It is the visible half of the lab rule that author claim, evidence
and presenter interpretation stay separate: the one slide where you say what
you tried, what transfers to our work, or where you disagree with the paper.
A paper talk with no `ours` slide is a warning from `qa_check.py`: a deck that
only relays the paper has left out the part the meeting is for.

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
| `paper_method`, `paper_results` | an exhibit (figure, video, diagram, matrix, equation or stage) | error |
| `project_results`, `research_results` | an exhibit | error |
| `paper_related` | an exhibit; explain the limitation in the claim, caption or notes | warning |
| `paper_intro` | an exhibit | warning |
| `paper_conclusion` | `bullets` (the takeaways) | warning |
| every content slide | a claim in title, subtitle or callout | error |

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

- **Comparison across methods or camps** → choose the relationship first:
  aligned examples for visible differences, diagrams for mechanisms, plots for
  trends or tradeoffs. Use `matrix` for compact lookup across shared criteria,
  with `highlight_row` when a row is the focus. See `evidence-selection.md`.
- **A pipeline or architecture you author** → native `diagram`, usually
  `figure-right` beside the explanation or `figure-full` when the topology is
  the whole argument. Use `accent` on the component this slide is about.
- **A pipeline quoted from a paper** → crop the original as `figure-right`,
  with a `box` annotation; redraw with `diagram` only when explicitly adapting
  and simplifying it, and say so in the caption and notes.
- **Numbers across methods** → `matrix`, `highlight_row` on the row you are
  arguing for and `highlight_col` on the metric. Not a figure of the paper's
  table. The decisive number gets its own `bignum` slide first.
- **Camps, contributions, pros vs cons, ours vs theirs** → `cards`, one per
  camp, the paper's own lit. Not a table with a Yes/No column.
- **The method overview** → `flow` when it is a line, `diagram` when it
  branches, loops or has groups; one step lit, then one slide per step with
  the paper's cropped figure on the right and a `box` annotation on the
  component. A wide architecture figure goes `figure-bottom` or
  `figure-full`, never into a half column: `qa_check.py` warns when a figure
  wider than ~1.8:1 lands under 7in.
- **Where this sits in the field** → `quadrant`, ours in the corner the paper
  claims.
- **The same figure across several slides** → reuse the same `src` and move
  the annotation. Both reference decks do this: one taxonomy figure appears on
  four consecutive slides with a red dashed box on a different column each
  time. It is the cheapest form of progressive disclosure available.
- **A loss or a definition** → `equation`, with the symbols expanded in
  bullets beside or below it. One equation per slide.
- **Qualitative results** → `figure-full`, and crop hard. A grid of eight
  methods is unreadable projected; crop to the three that matter.
- **Something that only reads as motion** — tracking jitter, temporal
  flicker, a robot actually completing the task → `video`, trimmed to the
  seconds that show it, autoplaying and looping. A still cannot carry a
  temporal claim, and a GIF pasted as an image will not animate in PowerPoint.
  Everything else stays a figure: a video of a static result is slower to read
  and costs megabytes.

## Annotation conventions

Red `C00000` for the thing under discussion, dashed for "this region of the
figure", solid for "this value". A label sits outside the picture edge where
possible. If an annotation needs more than four words it is a bullet, not an
annotation.
