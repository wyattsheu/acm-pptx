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
           kit.shade, kit.text_lines (measured wrap), kit.accent (this
           section's tab colour), kit.tones (its dark / base / pale / wash),
           kit.RED / RED_PALE / INK / GREY / LINE / PANEL / WHITE, kit.Inches,
           kit.Pt, kit.PP_ALIGN, kit.MSO_ANCHOR, kit.MSO_SHAPE

Draw only as tall as your content needs: compose.py moves what you drew up
under the bullets and centres the group, so a short drawing sits well and a
frame stretched to the full `box` height just shows up as empty panel.

Copy this file next to your outline and edit; the path in `draw` is resolved
relative to the outline.
"""


def timeline(slide, box, spec, kit):
    """Dated milestones on a horizontal line; the last one is lit."""
    events = spec["timeline"]["events"]
    x, y, w, h = box
    n = len(events)
    cy = y + 0.75                      # compose.py centres the drawing afterwards
    kit.arrow(slide, x, cy, x + w, cy, color=kit.tones["dark"], width=2.5)
    step = w / n
    for i, (when, what) in enumerate(events):
        cx = x + step * (i + 0.5)
        last = i == n - 1
        colour = kit.RED if last else kit.tones["dark"]
        d = 0.42 if last else 0.30
        kit.rect(slide, (cx - d / 2, cy - d / 2, d, d), fill=colour,
                 line=kit.WHITE, width=1.5, shape=kit.MSO_SHAPE.OVAL)
        kit.write(slide, when, box=(cx - step / 2, cy - 0.78, step, 0.46),
                  size=20, bold=True, color=colour, align=kit.PP_ALIGN.CENTER)
        lines = kit.text_lines(what, step - 0.3, 18)
        kit.write(slide, what, box=(cx - step / 2, cy + 0.32, step, 0.2 + lines * 0.32),
                  size=18, bold=last, color=kit.RED if last else kit.INK,
                  align=kit.PP_ALIGN.CENTER)


def two_column_compare(slide, box, spec, kit):
    """Left: what the prior work does. Right: what this paper does instead."""
    left, right = spec["compare"]["left"], spec["compare"]["right"]
    x, y, w, h = box
    cw = (w - 0.6) / 2
    size = 18
    # as tall as the longer column's text, never the whole region
    body = max(sum(kit.text_lines("• " + t, cw - 0.4, size) for t in p["points"])
               for p in (left, right)) * size * 1.3 / 72
    ph = min(h, 0.75 + body + 0.3)
    for i, (panel, lit) in enumerate(((left, False), (right, True))):
        px = x + i * (cw + 0.6)
        colour = kit.RED if lit else kit.tones["dark"]
        kit.rect(slide, (px, y, cw, ph), fill=kit.RED_PALE if lit else kit.tones["wash"])
        kit.rect(slide, (px, y, cw, 0.08), fill=colour)
        kit.write(slide, panel["title"], box=(px, y + 0.14, cw, 0.5), size=20,
                  bold=True, color=colour, align=kit.PP_ALIGN.CENTER)
        kit.write(slide, [{"text": "• " + t, "space_before": 6} for t in panel["points"]],
                  box=(px + 0.1, y + 0.72, cw - 0.2, ph - 0.8), size=size)
    kit.write(slide, "vs", box=(x + cw, y, 0.6, ph), size=20, bold=True,
              color=kit.GREY, align=kit.PP_ALIGN.CENTER, anchor=kit.MSO_ANCHOR.MIDDLE)
