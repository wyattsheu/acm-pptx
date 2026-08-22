# ACM Lab presentation rules

Distilled from *ACM Lab Research Guidelines — From Student to Researcher* and
*ACM Lab Handbook — Reading, Analyzing, and Presenting Research Papers*.
These govern content; `template-map.md` governs layout.

## Slide design standard

| Area | Rule |
|---|---|
| Title | State the message or claim. A generic label such as "Results" is listed in the handbook as a failure mode. |
| Text | Short phrases, one main idea per slide. No paragraph-sized blocks. |
| Font | Readable sans-serif, roughly 24pt or larger for core content, strong contrast. |
| Figures | Crop, enlarge, annotate, simplify. Cite borrowed visuals. |
| Tables | Only decisive rows and columns; highlight the values being discussed. |
| Colour | Never rely on colour alone — add labels, shapes, line styles. Caption videos; describe key visuals verbally. |
| Animation | Only to control the order of explanation. Embed video locally and test it. |
| Delivery | Face the audience, speak slightly slower, pause after key claims, signpost, do not read the slides. |

## Weekly meeting content

The research guidelines define what a meeting is for: not a status update, but
a way to improve research decisions. Before the meeting the presenter should be
able to answer — and the deck should reflect —

- What did I complete this week?
- Which results matched expectations, and which did not?
- What do I think the reasons are?
- What have I already tried?
- What do I suggest doing next?
- What needs discussion or support?

So `*_results` carries what happened, `*_insight` carries the analysis and the
open questions, and `*_conclusion` carries the commitments. A deck that only
lists completed tasks has missed the point of the meeting.

On problems: state the real problem, the possible causes, what was tried, what
worked and what did not, and your current proposed solution. Raise a slipping
commitment early rather than explaining it at the end.

## Paper presentation structure (20–25 min)

Rebuild the paper as an argument. Do not walk through it section by section.

| Section | Slides / time | The audience should leave understanding |
|---|---|---|
| Paper information | 1 / 0.5 min | What it is, where it appeared, why you picked it, relevance to the lab |
| Motivation + gap | 2–4 / 4 min | Why the problem matters and what is unresolved |
| Concept + hypothesis + contributions | 3–4 / 4 min | How to think about the problem, what is claimed, what is contributed |
| Related work | 1–2 / 2 min | How this differs from prior approaches |
| Method overview + details | 3–4 / 6 min | The full input-to-output pipeline and the design rationale |
| Setup + evidence | 3–5 / 6 min | How claims are tested; results, ablations, examples, efficiency |
| Critique + lab relevance + takeaways | 2–3 / 3 min | Strengths, weaknesses, limitations, future work, what to remember |

Keep three statement types visibly separate: **author claim**, **evidence**,
and **presenter interpretation**.

For each method component answer four questions: what does it receive, what
does it do, what does it produce, why is it needed. Listing module names
without design logic is a failure mode.

Evidence slides: a quantitative slide answers a claim; an ablation slide
answers why the method works. Name the metric, the strongest relevant
baseline, the decisive values, and the exact conclusion supported.

The required ending — three takeaways, strengths, weaknesses, limitations
(separating author-acknowledged from presenter-identified), lab relevance, and
next directions. Ending on a bare "Questions?" is a failure mode; keep the
takeaway slide up during Q&A.

## Failure modes to avoid

Following the paper section by section · too much background · reading the
slides · dense uncropped figures and full tables · generic titles · no critique
or overclaiming · no rehearsal · ending with only "Questions?".

## What this means when generating slides

- Do not invent results, metrics, or numbers. Ask, or leave the slot for the user.
- Do not pad a slide to fill it — whitespace is fine, a paragraph is not.
- Turn every placeholder title into a claim before shipping.
- Put the argument in the speaker notes, the claim on the slide.
