# outline.json schema

The contract between "what to say" and "what to build". Any front end can
produce it; `build_from_outline.py` consumes it. Keep it stable — it is what
lets the plain skill hand work off to the executable one.

```jsonc
{
  "meta": {                       // presence of meta adds the cover slide
    "title": "Progress Report",
    "date": "2026.08.23",
    "advisor": "Prof. Huang Ching-Chun",
    "student": "Your Name"
  },

  "summary": {                    // presence of summary adds slide 2
    "rows": [                     // 4 rows, [label, value] — use NONE for empty tracks
      ["Project Progress",  "[Sponsor] Project name"],
      ["Research Progress", "NONE"],
      ["Paper Study",       "[CVPR 2026] Paper title"],
      ["Other Tasks",       "..."]
    ],
    "tasks": [                    // up to 8; extra rows are blanked
      "✅ finished thing",
      "🔄 in-progress thing"
    ],
    "note": " summarize what did you do within a week"   // optional caption
  },

  "slides": [
    {
      "role": "project_results",  // must be a key from --roles
      "title": "Cache hit rate is what moved latency, not batch size",
      "subtitle": "optional grey line under the title (agenda slides only)",
      "bullets": [
        {"text": "Header tier — bold, no bullet", "level": 0},
        {"text": "Bulleted tier", "level": 1},
        "plain string is shorthand for level 0"
      ],
      "notes": "Speaker notes, in 中文, full sentences, 3-6 per content slide, no length limit. The argument, the numbers you'll quote and their baseline, the questions you expect."
    },
    {
      "role": "project_conclusion",
      "title": "Conclusion",
      "table": ["first todo", "second todo"]        // numbered table, rest blanked
    },
    {
      "role": "paper_list",
      "title": "Paper Study",
      "table": [                                    // 2-D form for multi-column tables
        ["✓", "1", "CVPR 2026", "Paper title"]
      ]
    }
  ]
}
```

## Rules

- `role` is required and must exist; the builder refuses unknown roles rather
  than guessing.
- `bullets` and `table` are mutually exclusive on a slide — a table role
  ignores bullets.
- `level` maps onto the template's own indent tiers, discovered per slide from
  the template's `marL` values. A level above the deepest available tier
  clamps to the deepest. Never fake indentation with spaces or dashes.
- Omit `meta` to build without a cover; omit `summary` to build without one.
  `slides` alone is a valid outline.
- Roles may repeat.
- Unused table rows are blanked, never left as `XXX`.

## Round-tripping

`acm-slides-plain` emits exactly this JSON and stops. Paste it into an
environment that can execute code and run:

```bash
python scripts/build_from_outline.py outline.json -o report.pptx
```

Nothing else needs to transfer — the template lives in the package, not in
the outline.
