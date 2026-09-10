# ACM Lab template — role map

`assets/acm_template.pptx`, 22 slides, 13.333in × 7.5in. Each slide has its
own layout; the layout carries the coloured section tab and the breadcrumb
strip, with the current step already bold. That is why roles must be cloned
rather than re-created — the breadcrumb is not text you can retype into an
arbitrary layout.

Section colours: **Project** green · **Research** blue · **Paper** purple ·
**Others** pink.

| Role | Slide | Section | What's on it |
|---|---|---|---|
| `cover` | 1 | — | `Progress Report` / date, ACM logo, Advisor + Student lines |
| `summary` | 2 | — | 4-row status table (Project / Research / Paper Study / Other Tasks) + 8-row numbered weekly task table |
| `project_summary` | 3 | Project | Agenda page: Goal & Motivation / Architecture Recap / Status & Results / Discussion Items / Conclusion, with per-section timings |
| `project_why` | 4 | Project | Why This Project Matters — breadcrumb on *Goal & Motivation* |
| `project_pipeline` | 5 | Project | Current Pipeline Overview — breadcrumb on *Current Architecture Recap* |
| `project_results` | 6 | Project | Progress & Results — breadcrumb on *Current Status & Results* |
| `project_insight` | 7 | Project | Insight / things to raise with the advisor — breadcrumb on *Discussion Items* |
| `project_conclusion` | 8 | Project | Conclusion, 8-row numbered todolist table |
| `research_summary` | 9 | Research | Agenda page, same structure as slide 3 |
| `research_why` | 10 | Research | Why This Research Matters |
| `research_prev` | 11 | Research | Previous Method — prior work, gaps, how yours differs |
| `research_pipeline` | 12 | Research | Current Pipeline Overview |
| `research_results` | 13 | Research | Progress & Results |
| `research_insight` | 14 | Research | Insight / discussion items |
| `research_conclusion` | 15 | Research | Conclusion, 8-row numbered todolist table |
| `paper_list` | 16 | Paper | Paper Study — 8×4 table: done / no. / venue+year / title |
| `paper_intro` | 17 | Paper | Introduction — breadcrumb on *Introduction* |
| `paper_related` | 18 | Paper | Related work |
| `paper_method` | 19 | Paper | Proposed Method |
| `paper_results` | 20 | Paper | Experimental Results |
| `paper_conclusion` | 21 | Paper | Conclusion |
| `others` | 22 | Others | Free-form |

`project_prev` is an alias that reuses slide 11 — it carries the **blue**
Research tab, so only use it if the section colour is acceptable.

## Empty tracks

The template says it explicitly on slides 3 and 9: if there is no project (or
no research) this week, still show that summary slide and set every term to
`NONE`. Do not silently drop the track.

## If the template is revised

The role numbers above are positional. When the lab ships a new template,
re-run `python scripts/thumbnail.py assets/acm_template.pptx tpl-thumbs`,
compare against this table, and update both the table and `ROLES` in
`scripts/build_from_outline.py`. Bump `metadata.template_version` in SKILL.md so a
stale map is visible rather than silent.
