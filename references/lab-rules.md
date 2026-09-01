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

## Writing discipline

The rules above decide what goes on a slide. These decide how it is worded —
they are what separates a deck that reads like a researcher's argument from one
that reads like generated filler.

**Telegraphic language.** Drop articles, hedges, and framing clauses whenever
meaning survives without them. Write `Cost down 23% (p < 0.01)`, not
`Our study found that the intervention significantly reduced costs`. Packaging
phrases — *we propose*, *it is worth noting*, *this demonstrates that*,
*in order to* — are the clearest tell of machine-written slides. Cut them.

**What telegraphic does not mean.** The target is words that carry no
information, not information itself. A line is too short the moment the
audience cannot reconstruct what it refers to. Never compress away:

- the baseline or condition a number is measured against — `+3.2 dB` alone is
  unreadable; `+3.2 dB PSNR over SwinIR` is not
- units, dataset, and split (`mIoU on Cityscapes val`)
- a symbol or acronym on first appearance — expand it once, then abbreviate
- which part of a pipeline a claim applies to, when the deck has several
- the difference between what an author claims and what the evidence shows

If a slide reads as terse but empty, the fix is usually to add back one of the
above — not to relax the word ceiling and write prose again. Test: hand the
slide to a labmate who is not on this project. If they must ask "compared to
what?" or "on which set?", the compression went too far.

**~40 words of body text per slide, hard.** Past that the slide is doing two
jobs: split it, push the detail into the speaker notes, or move it to the
appendix. Three to five bullets is typical; more than five is a warning. The
ceiling covers slide body text only. It has never applied to speaker notes.

**Speaker notes carry the detail the slide sheds.** The two rules work as a
pair: the slide holds the claim, the notes hold the argument. Compression on
the slide is only legitimate because the notes are full — a deck with terse
slides *and* thin notes has lost the content, not compressed it.

Notes are written in the language the presenter will actually speak — 中文 by
default for lab meetings, with technical terms left in English (`我們把
attention map 拿掉之後 mIoU 掉了 4 個點`). They are full sentences, not
fragments, and have no length limit; three to six sentences per content slide
is normal, more on a method or results slide.

Each note should cover, as applicable: how this slide follows from the previous
one; what the numbers on screen mean and against what baseline; why this design
choice rather than the obvious alternative; what is still uncertain; the
question this slide is likely to draw and the answer. Write them as speech, not
as a written summary — if it cannot be said out loud comfortably, rewrite it.

**One argument per talk.** The instinct is to present everything done. Pick the
claim that can be carried convincingly in the time available; everything else
is appendix material. A deck that covers three things well convinces of none.

**Ghost deck test.** Extract every title, read them in sequence, ignore the
bodies. They must tell the whole argument on their own. If the sequence is
merely a list of topics, the titles are labels, not claims — rewrite before
building the deck.

**Flow test.** Each title should make the next feel like the natural next step.
A slide that could sit anywhere in the deck without loss is either misplaced or
unnecessary.

**One exhibit per slide, annotated.** One chart, table, diagram, or equation
block. Mark the decisive point on the figure itself — arrow, call-out,
highlighted region, contrasting colour for the focal series. Two tests: cover
the exhibit, and the title should still stand; cover the title, and the
takeaway should still be obvious. Failing the first means the exhibit is
unnecessary; failing the second means it needs annotation.

**Self-sufficient slides.** Decks circulate as PDFs after the meeting. A slide
whose point collapses without narration needs a stronger annotation or a
sharper title.

**Rebuild figures, don't screenshot them.** Paper figures carry print-sized
fonts and captions that will not survive projection. Rebuild at presentation
scale, axis labels 16pt or larger.

**Bold and italics carry meaning, not decoration.** Bold for a key term on
first use, inline labels (`Note:`, `Limitation:`), and the focal number.
Italics for notation and titles. Nothing else.

Adapted from *academic-pptx-skill* (Gabberflast, MIT), whose content rules draw
on Minto's *Pyramid Principle* and Naegle (2021), "Ten simple rules for
effective presentation slides," *PLOS Comput Biol*. Its visual standards
(white background, single font) are deliberately not adopted — the lab template
governs all layout and colour.

## What this means when generating slides

- Do not invent results, metrics, or numbers. Ask, or leave the slot for the user.
- Do not pad a slide to fill it — whitespace is fine, a paragraph is not.
- Turn every placeholder title into a claim before shipping.
- Put the argument in the speaker notes, the claim on the slide.
