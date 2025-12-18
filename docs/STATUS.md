# AutoBericht: Current State vs Target

This document is written for a corporate/offline environment:
- No external network calls at runtime (no CDN, no telemetry, no remote fonts/images).
- Prefer workflows that work on Windows without browser security flags.
- Canonical `project.json` contract: `docs/architecture/data-model.md` (root `chapters[].rows[]`, plus `photos`, `lists`, `history`).
- Canonical implementation: the Excel/VBA exporter/loader (`macros/modABProjectExport.bas`, `macros/modABProjectLoader.bas`) is the source of truth; the web UI adapts to VBA, not the other way around.

## Intended Functioning (Target)

### End-to-end workflow
1. **Excel (project.xlsm) imports sources**
   - Word “master findings” → structured sheets
   - Customer self-assessment Excel → structured sheets
   - PhotoSorter tags photos → structured sheets (`Photos`, `PhotoTags`, `Lists`)
2. **Excel exports a single canonical file**
   - `project.json` (unified project state: meta + chapters/rows + photos + lists + history)
3. **Web UI edits and exports**
   - Loads `project.json` (and/or reads the workbook directly)
   - Edits consultant workstate (levels, overrides, include flags, done/review)
   - Exports:
     - printable PDF via browser
     - PPTX decks
     - updated `project.json` (and optionally an updated workbook snapshot)
4. **Optional post-processing**
   - Excel/VBA can re-import updated JSON and/or apply file renaming/moves for photos.

### Runtime constraint (offline)
All web UI JS/CSS/libs are local under `AutoBericht/` and must not trigger any external requests.

## Where the Repo Is Now

### Good / already in place
- **Offline-first web UI** with local `libs/` and ES-module structure (`AutoBericht/index.html`).
- **Structured-sheet VBA architecture** to load/export `project.json` (`macros/modABProjectLoader.bas`, `macros/modABProjectExport.bas`).
- **PhotoSorter VBA modules** and sheet constants include `PhotoTags` (`macros/modABConstants.bas`).
- **No-flags launcher for Windows**: `start-autobericht.cmd` runs a tiny local server so modules work without `--disable-web-security`.

### Primary mismatch blocking “it just works”
There are currently *two different* “project.json” shapes in the repo:
- **Docs + VBA** describe/export a unified JSON with root `chapters` (and `rows` inside chapters): `docs/architecture/data-model.md`.
- **Web UI** currently expects a different scaffold shape with root `report` / `presentation` and uses `report.chapters` (see `AutoBericht/js/schema/definitions.js`).

This means:
- Exporting `project.json` from Excel and importing it into the web UI will fail (unless the web UI is adapted).
- The web UI exporters (PDF/PPTX) currently read the web-UI-specific `project.report.chapters`, not the unified `chapters[].rows[]`.

## Immediate Next Steps (Recommended)

### 1) Decide the single canonical data contract
Decision: treat the **Docs/VBA unified `project.json`** as canonical (it matches the Excel structured sheets and the stated architecture).

### 2) Make the web UI speak the canonical contract
Work items:
- Update `AutoBericht/js/schema/definitions.js` so `projectSchema` matches the unified JSON.
- Refactor `ProjectState` so `project.json` is the primary source of truth (not “master + selfEval + overrides” as separate files).
- Update the report UI + exporters to iterate `chapters[].rows[]`.
- No migration/adapter needed yet (alpha; no legacy snapshots).

### 3) PhotoSorter reliability checklist
Because PhotoSorter is Excel-driven, “working” means:
- Sheets exist and headers match constants (`Meta`, `Chapters`, `Rows`, `Photos`, `PhotoTags`, `Lists`).
- PhotoSorter reads button lists from `Lists` and writes tags into `PhotoTags`.
- Export includes `photos` + tag arrays and `lists` vocabularies.

## How to Run the Web UI (Corporate/Offline)
- Windows: double-click `start-autobericht.cmd`.
- Alternative: run any local static server from the `AutoBericht/` folder (no internet required).
