# AutoBericht Web UI Workflow Guide (Legacy)

> This guide documents the legacy three-tab UI. The current direction is the
> minimal editor workflow in `docs/guides/redesign-workflow.md`.

This guide covers the offline HTML/JavaScript interface in `AutoBericht/legacy/`.

The web UI adapts to the Excel/VBA `project.json` contract (VBA is canonical). Photo tagging is performed in Excel (PhotoSorterForm); the web UI consumes the resulting photo metadata and uses it for exports.

## Overview

**Three tabs**:

1. **Import**: Load `project.json` and/or `project.xlsm`, load photos, edit tag lists, run validation
2. **AutoBericht**: Edit findings/recommendations, choose levels, apply overrides, mark done
3. **Export**: Generate PDF/PPTX outputs and download `project.json`

## Opening the Application

### Recommended (Windows): repo launcher

Use the repository launcher to run a local, offline-only server so ES modules work without browser flags:

- `start-autobericht.cmd`
- `start-autobericht.ps1`

### Alternative: Python dev server

From the repository root:

```bash
python -m http.server 8000
```

Then open `http://127.0.0.1:8000/AutoBericht/legacy/`.

## Import & Setup (Import tab)

### 1) Load your project data

Preferred sources (use what you have):

- **Project workbook (`.xlsm`)**: imports structured sheets (Meta/Chapters/Rows/Photos/Lists/History)
- **Existing `project.json` (optional)**: restores an existing snapshot (must match the VBA-canonical shape)

### 2) Load photos (directory)

Upload the photo directory so exports can resolve filenames referenced by the project. Photo tags and notes come from Excel → `project.json` / workbook import.

### 3) Branding (optional)

Upload left/right logo images for PDF/PPTX exports.

### 4) Tag lists (Topics / Seminar)

Edit the tag button lists used by Excel PhotoSorter and seminar decks:

- Topics (`lists["photo.topic"]`)
- Seminar tags (`lists["photo.seminar"]`)

If you change these lists in the web UI, export `project.json` and re-import it in Excel so PhotoSorter stays in sync.

### 5) Validation

Validation must be green before exports enable. If validation fails:

- Confirm your `project.json` was exported by Excel/VBA
- Re-export `project.json` from Excel after imports/tagging

## Editing Content (AutoBericht tab)

1. Pick a finding row from the chapter tree.
2. Choose the appropriate recommendation level (1–4).
3. Use overrides only when needed:
   - Finding override: edits the master finding text
   - Level override: edits the recommendation text for a specific level
4. Toggle inclusion flags (finding/recommendation) and mark the row done when finished.

The web UI writes overrides into both `overrides.*` (VBA import/export) and `workstate.*` (UI state), so Excel can load the results reliably.

## Exports (Export tab)

- **Download `project.json`**: saves the current project snapshot for re-import to Excel.
- **PDF**: browser print output using the current PDF layout settings.
- **PPTX**: generates decks from the selected rows and photo tags.

## Round-trip back to Excel

1. In the web UI Export tab: download `project.json`.
2. In Excel: run the JSON import macro (see `docs/guides/vba-workflow.md`).
3. Verify:
   - Rows load with the expected selected levels and overrides
   - Photo tags and lists remain intact

## Keyboard shortcuts

- `Alt+1` / `Alt+2` / `Alt+3`: switch Import / AutoBericht / Export
- `Ctrl+/`: toggle shortcut overlay
- `Esc`: close the shortcut overlay

## Related documentation

- `docs/guides/getting-started.md`
- `docs/guides/vba-workflow.md`
- `docs/architecture/data-model.md`
- `docs/reference/roundtrip-smoke-checklist.md`
