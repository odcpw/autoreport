# AutoBericht System Overview (Redesign)

Browser-first offline workflow with project-local state and direct Word/PPT export from the web app.

## High-Level Architecture

```text
┌──────────────────────────────────────────────────────────────┐
│                          AutoBericht                          │
└──────────────────────────────────────────────────────────────┘

┌──────────────────────┐      writes/reads      ┌────────────────────────┐
│  Minimal Editor UI   │ ─────────────────────▶ │  project_sidecar.json   │
│  (AutoBericht/mini)  │ ◀───────────────────── │  (working state)        │
└──────────────────────┘                        └────────────────────────┘
           │
           │ reads sidecar + templates
           ▼
┌──────────────────────┐                        ┌────────────────────────┐
│  Export Engine       │ ─────────────────────▶ │  Word/PPT templates     │
│  (in web app)        │                        │  (project/templates)    │
└──────────────────────┘                        └────────────────────────┘
           │
           ▼
┌──────────────────────────────────────────────────────────────┐
│ Outputs: Word report, PPT report deck, PPT training deck      │
└──────────────────────────────────────────────────────────────┘
```

## Project Folder Layout

```text
<Project>/
  project_sidecar.json      # canonical project state
  library_user_*.json       # reusable knowledge library
  inputs/                   # customer docs + self-assessment
  outputs/                  # generated reports and exports
  templates/                # DOCX/PPTX templates
  photos/
    raw/pm1|pm2|pm3
    resized/
    export/
```

## Responsibilities

- Browser editor
  - chapter editing
  - recommendation/finding text maintenance
  - filters, scoring, include/done workflow
  - 4.8 organization and checklist support
- PhotoSorter
  - raw import/rename/resize
  - report/observation/training tagging
  - tagged folder export
- Project page
  - locale bootstrap
  - library update/export
  - Word and PowerPoint exports

## Why This Split

- Simple offline flow for engineers.
- No Office add-ins/macros required for daily work.
- Clear separation: sidecar = project state, library = reusable knowledge.

## Related

- [Design Spec](design-spec.md)
- [Workflow](workflow.md)
- [Project Template](project-template.md)
