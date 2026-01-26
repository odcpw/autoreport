# AutoBericht System Overview (Redesign)

This overview documents the **current redesign**: browser-first editing with a
minimal Word/VBA export layer.

## High-Level Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                          AutoBericht                          │
└──────────────────────────────────────────────────────────────┘

┌──────────────────────┐      writes/reads      ┌────────────────────────┐
│  Minimal Editor UI   │ ─────────────────────▶ │  project_sidecar.json   │
│  (AutoBericht/mini)  │ ◀───────────────────── │  (working state)        │
└──────────────────────┘                        └────────────────────────┘
           │
           │ optional
           ▼
┌──────────────────────┐      reads             ┌────────────────────────┐
│  Word/VBA Exporter  │ ─────────────────────▶ │  Word/PPT templates     │
│  (thin layer)        │                        │  (in project root)      │
└──────────────────────┘                        └────────────────────────┘
           │
           ▼
┌──────────────────────────────────────────────────────────────┐
│ Outputs: Word report, PPT deck, PDF, optional photo folders   │
└──────────────────────────────────────────────────────────────┘
```

## Project Folder Layout

```
<Project>/
  project_sidecar.json      # editor state (canonical)
  project_db.xlsx           # optional readable archive
  self_assessment.xlsx      # customer input (static)
  photos/                   # raw photos (import uses photos/raw/pm1, pm2)
  out/                      # Word/PPT/PDF outputs
  cache/                    # optional materialized photo views
```

## Data Flow (Redesign)

```
Self-assessment + photos
          │
          ▼
Minimal Editor (edit findings, recommendations, levels, tags)
          │
          ▼
project_sidecar.json (canonical state)
          │
          ├── Optional: write project_db.xlsx (human-readable)
          │
          └── Word/VBA export → Word/PPT/PDF outputs
```

## Responsibilities

- **Browser editor**
  - Chapter-based editing.
  - Recommendation text editing + scoring level selection.
  - Photo tagging (virtual tags).
  - Sidecar load/save.

- **Word/VBA** (thin layer)
  - Read sidecar/export JSON for report generation.
  - Export Word/PPT/PDF using Word template macros.
  - Optional photo folder materialization.

## Why This Split

- Keeps daily work out of VBA (less friction).
- Preserves template fidelity via Office automation.
- Stays compliant with locked-down corporate IT.

## Related

- [Design Spec](design-spec.md)
- [Redesign Workflow](../guides/redesign-workflow.md)
