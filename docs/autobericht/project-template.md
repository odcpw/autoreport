# Project Folder Layout

Current flow: user selects an **empty** project folder, then AutoBericht creates/scaffolds required folders.

## Created Scaffold

```text
<Project>/
  inputs/
  outputs/
  backup/
  templates/
    Vorlage IST-Aufnahme-Bericht d.V01.docx
    Vorlage IST-Aufnahme-Bericht f.V01.docx
    Vorlage IST-Aufnahme-Bericht i.V01.docx
    Vorlage AutoBericht.pptx
  photos/
    raw/
      pm1/
      pm2/
      pm3/
    resized/
    export/
```

## Core Data Files

- `project_sidecar.json`: canonical project state (created on first save/bootstrap)
- `library_user_*.json`: reusable knowledge library (optional, project-local)

## Notes

- Keep all paths lowercase for consistency.
- Photo import reads from `photos/raw/pm1..pm3` and writes to `photos/resized`.
- Tagged photo export writes into `photos/export`.
- Template pickers should default to `<Project>/templates`.
