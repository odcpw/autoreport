# AutoBericht VBA Modules (Structured Sheets)

Import the `.bas` files under `macros/` into the macro-enabled workbook. All read/write operations then occur against the JSON-aligned structured sheets, ready for AutoBericht and VBA orchestration.

## Module Overview

| Module | Purpose |
| --- | --- |
| `modABConstants` | Central sheet names, header definitions, and shared constants (schema version, overwrite modes). |
| `modABWorkbookSetup` | Ensures required worksheets exist, writes header rows, and clears tables when needed. |
| `modABTableUtils` | Generic helpers to resolve column indexes, read/write tables, and upsert rows. |
| `modABRowsRepository` | Convenience functions for manipulating a single finding row, recording override history, and touching metadata. |
| `modABPhotosRepository` | Helpers for reading/writing photo metadata and PhotoSorter button lists. |
| `PhotoSorter` | Directory import helpers and folder utilities for the PhotoSorter experience. |
| `CPhotoTagButton` (class) | Wraps dynamic PhotoSorter buttons, tracks tag state, updates styling/counts. |
| `modABProjectLoader` | Reads `project.json` and populates all structured sheets (`Meta`, `Chapters`, `Rows`, `Photos`, `Lists`, `OverridesHistory`). |
| `modABProjectExport` | Builds the JSON snapshot (`project.json`) from the structured sheets. |

## Typical Workflow

1. **Initial setup**
   ```vb
   Sub SetupProject()
       EnsureAutoBerichtSheets clearExisting:=True
   End Sub
   ```
   Creates the structured sheets with headers.

2. **Load an existing project**
   ```vb
   Sub LoadProject()
       LoadProjectJson "C:\path\to\project.json"
   End Sub
   ```
   Populates every structured sheet from the JSON payload (meta, chapters, rows, photos, lists, override history). Use `PopulateProjectTables` if the JSON has already been parsed in VBA (e.g., when chaining imports).

3. **Edit data**
   - VBA forms (PhotoSorter, report editor) should read/write the structured sheets through `modABRowsRepository` and `modABPhotosRepository`.
   - Use `SetFindingOverride`, `SetRecommendationOverride`, `SetIncludeFlags`, and `TogglePhotoTag` to keep override history consistent.

4. **Export JSON**
   ```vb
   Sub ExportProject()
       ExportProjectJson "C:\\path\\to\\project.json"
   End Sub
   ```
   Produces a schema-compliant JSON payload:
   - `meta`, `chapters` (each with `rows` and override/customer/workstate blocks),
   - `photos` (with tagged chapter/category/training/topic arrays),
   - `lists` (button definitions for the PhotoSorter/filters),
   - `history` (override change log).

## Dependencies

- Add a **reference** to *Microsoft Scripting Runtime* (Dictionary support).
- Import **VBA-JSON** (`JsonConverter.bas`) and make it available in the project (`JsonConverter.ConvertToJson`).
- Legacy PhotoSorter/UserForms can continue to exist; update their code to consume the structured sheets instead of `PSHelperSheet` once the migration is complete.

## Structured Sheet Layout

The modules automatically create the following sheets:

- `Meta`: key/value pairs (projectId, company, locale, createdAt, author).
- `Chapters`: hierarchy with localisation columns (`defaultTitle_{de,fr,it,en}`), `parentId`, `orderIndex`, `pageSize`, `isActive`.
- `Rows`: one record per finding, including master texts, override fields (`overrideLevel1..4`, `useOverrideLevel1..4`), customer answers, inclusion toggles, and general workstate metadata.
- `Photos`: filename → display name, notes, chapter/category/training/topic tags, locale/capture metadata.
- `Lists`: vocabulary for PhotoSorter buttons / filters (supports localisation + chapter references).
- `ExportLog` (empty placeholder for future renumber tracking).
- `OverridesHistory`: append-only change log automatically filled by `modABRowsRepository.UpdateRowField`.

These sheets map one-to-one with the JSON described in `docs/unified_project_schema.md`. AutoBericht can now read/write the workbook directly using SheetJS without translating from ad-hoc helper tables.

### PhotoSorter Form
- Replace the existing `PhotoSorterForm` code-behind with `macros/PhotoSorterForm.frm` and import `CPhotoTagButton.cls`.
- The form now reads tag button definitions from the `Lists` sheet, displays photos from the `Photos` sheet, and updates tag assignments via `modABPhotosRepository`.
- Directory scanning (`ScanImagesIntoSheet`) populates new photo rows; toggles and counts operate solely on the structured tables.
