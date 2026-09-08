---
name: document-synthesis
description: Plan, configure, build, and verify multi-source Word evidence packages in this repository. Use for material-package assembly, source strategy selection, project profiles, build diagnosis, or output QA; not for unrelated standalone document editing.
---

# Document synthesis

Work from the repository root. This skill covers the repository's unified
engine; do not create a project-specific build script when a profile or
manifest can express the variation.

## Start safely

1. Inspect the proposed source with `python3 synthesize.py --source path/to/source --plan`.
2. If a build is requested, first run `python3 synthesize.py --doctor --source path/to/source`.
3. Build only after the plan and environment status are acceptable. The engine
   stages work beneath its output directory and publishes only after its QA
   gates pass.

`--plan` is read-only and never calls Office. Prefer it for a new source,
configuration change, or troubleshooting request.

## Configure by source shape

- Use `directory_tree` for an ordinary ordered folder of materials.
- Use `explicit_tree` when the required hierarchy or order cannot be inferred.
  Use a `docx_outline` node when headings must be read from a DOCX without
  duplicating its body.
- Use `highlighted_docx` only for a single DOCX whose headings are deliberately
  marked with configured highlight colors.

Keep configuration in a project-local `manifest.json`, a reusable
`profiles/project-name.json`, or an explicit `--manifest`. Validate changes with
`--plan`; do not add a new entrypoint or hard-code a project name into the
engine.

## Build and diagnose

Use `synthesize.py` as the sole build entrypoint. Inputs are read-only. Output
names must stay within the output directory and use `.docx` names.

For Word Automation failures, distinguish file access from Apple-event
permission: Full Disk Access does not grant permission to control Word. Use the
`--doctor` result before suggesting a build retry. Do not claim exact printed
page numbers unless the Word export and page-level QA completed.

## Format learning and release evidence

The format-learning workflow also uses `synthesize.py` as its sole entrypoint:
`--analyze-format` creates `analysis.json`, `review.html`, and a draft
`decisions.example.json`; a compiler must receive explicit final decisions,
never a draft with `pending` values. Preserve the report ID and source SHA-256
when compiling a format package or a target RoleMap. Use `--render-preview` for
fictional approximate previews and `--migrate-manifest` for non-destructive
v1/v2 migration to a separate output file.

Before release, keep a status record tying each claimed capability to its test
ID, support level, and recent evidence. Run example commands in a fresh output
directory, including `--plan`, `--doctor`, build, and `smoke_test.py`; remove
`.work`, Office lock files, render caches, and Python caches afterward. Do not
generalize a real Word result beyond its fixture, environment, and declared
format boundary.

## Verify

Run the focused unit suite after source-strategy, safety, or configuration-code
changes:

```bash
python3 -m unittest discover -s tests -v
```

Run `git diff --check` before handing off code changes. Do not retain `.work`,
cache, rendering, or exploratory output directories after verification.

## Safe demo

`examples/minimal-demo/` is a fully fictional, location-free, person-free
demonstration project. Its README gives the read-only command and regeneration
command. Treat the demo as a smoke path, not as a template for sensitive source
materials.
