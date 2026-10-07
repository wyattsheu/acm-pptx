"""Reference implementation of a custom draw function for acm-pptx.

Name it from a slide in outline.json:

    {"role": "paper_method", "title": "...", "subtitle": "...",
     "draw": "design.py:timeline",
     "timeline": {"events": [["2020", "NeRF"], ["2023", "3DGS"], ["2026", "ELITE"]]}}

compose.py calls `timeline(slide, box, spec, kit)`:

    slide  the python-pptx slide, already carrying title / subtitle / bullets
    box    (x, y, w, h) in inches -- the exhibit region the layout left free;
           draw inside it and nothing will overlap
    spec   the slide's outline entry, so your own fields travel with it
    kit    helpers: kit.rect, kit.write, kit.arrow, kit.rgb, kit.tint,
           kit.accent (this section's tab colour), kit.RED / INK / GREY /
           LINE / PANEL / WHITE, kit.Inches, kit.Pt, kit.PP_ALIGN,
           kit.MSO_ANCHOR, kit.MSO_SHAPE

Copy this file next to your outline and edit; the path in `draw` is resolved
relative to the outline.
"""


def timeline(slide, box, spec, kit):
    """Dated milestones on a horizontal line; the last one is lit."""
    events = spec["timeline"]["events"]
    x, y, w, h = box
    n = len(events)
    cy = y + h * 0.45
    kit.arrow(slide, x, cy, x + w, cy, color=kit.GREY, width=2.0)
    step = w / n
    for i, (when, what) in enumerate(events):
        cx = x + step * (i + 0.5)
        last = i == n - 1
        colour = kit.RED if last else kit.accent
        d = 0.3 if last else 0.22
        kit.rect(slide, (cx - d / 2, cy - d / 2, d, d), fill=colour,
                 line=kit.WHITE, width=1.0, shape=kit.MSO_SHAPE.OVAL)
        kit.write(slide, when, box=(cx - step / 2, cy - 0.62, step, 0.36),
                  size=13, bold=True, color=colour, align=kit.PP_ALIGN.CENTER)
        kit.write(slide, what, box=(cx - step / 2, cy + 0.22, step, h * 0.45),
                  size=14 if last else 13, bold=last, align=kit.PP_ALIGN.CENTER)


def two_column_compare(slide, box, spec, kit):
    """Left: what the prior work does. Right: what this paper does instead."""
    left, right = spec["compare"]["left"], spec["compare"]["right"]
    x, y, w, h = box
    cw = (w - 0.6) / 2
    for i, (panel, lit) in enumerate(((left, False), (right, True))):
        px = x + i * (cw + 0.6)
        colour = kit.RED if lit else kit.accent
        kit.rect(slide, (px, y, cw, h), fill=kit.tint(colour, 0.9), rounded=True)
        kit.write(slide, panel["title"], box=(px, y + 0.08, cw, 0.5), size=18,
                  bold=True, color=colour, align=kit.PP_ALIGN.CENTER)
        kit.write(slide, [{"text": "• " + t} for t in panel["points"]],
                  box=(px + 0.1, y + 0.7, cw - 0.2, h - 0.8), size=15)
    kit.write(slide, "vs", box=(x + cw, y, 0.6, h), size=20, bold=True,
              color=kit.GREY, align=kit.PP_ALIGN.CENTER, anchor=kit.MSO_ANCHOR.MIDDLE)
