# AutoBericht Workbook Flow (Legacy-heavy)

> This describes the full VBA-driven pipeline. In the redesign, VBA is minimized
> to import/export and Office template automation.

This captures the end-to-end path from pulling macros out of GitHub, through the data import routines, into the Bericht editor, and finally to JSON export. Use it as the working reference while stabilising the Excel side before wiring up the web UI.

## 1. Macro refresh / distribution

1. **Sync sources**
   - Run `sync-autobericht.ps1` (see [PowerShell Sync](../reference/powershell-sync.md)).
   - Downloads the GitHub archive and drops `macros/` + `docs/` into the workstation folder (default `C:\Autobericht\macros`).
2. **Import into Excel**  
   - Open the workbook (`project.xlsm`).  
   - Run `modMacroSync.RefreshProjectMacros`.  
   - Flow:
     - Preserves `ThisWorkbook`, worksheets, and `modMacroSync`.  
     - Exports existing components to `_export\` for backup.  
     - Imports every `.bas/.cls/.frm` in the source folder, so new modules (`modReportData`, `modABIdUtils`, etc.) are automatically registered.

> Once this succeeds, the workbook codebase is in sync with the repository contents.

## 2. Master + customer data ingestion

### 2.1 Master findings (Word → Rows)
1. Run `ImportMyMaster` macro.
2. Word picker opens → choose the DOCX with strict header row `ReportItemID|Feststellung|Level1|...|Level4`.
3. Macro flow:
   - Parses Word tables, normalises IDs via `modABIdUtils.NormalizeReportItemId`.
   - Ensures each row exists in sheet `Rows` (uses `modABRowsRepository.EnsureRowRecord`).
   - Writes `masterFinding`, `masterLevel1..4`, and auto-fills `chapterId` with the parent ID.
   - Default row state set (include flags true, selected level = 2, overrides cleared).
4. Summary dialog shows inserted vs updated counts.

### 2.2 Customer self-assessment (Excel → Rows)
1. Run `ImportSelbstbeurteilungKunde`.
2. Choose the `Selbstbeurteilung Kunde` workbook.
3. Macro flow:
   - For each populated `ReportItemID`, normalises the ID.
   - Ensures a row exists in `Rows`.
   - Copies columns D/E/F into `customerAnswer`, `customerRemark`, `customerPriority`.
   - Tracks inserted vs updated rows and reports counts.

At this point `Rows` holds the union of master content and customer answers; overrides remain blank.

## 3. Bericht editor interaction

1. Launch `BerichtForm` (Report UI).
2. On initialise:
   - `modReportData.LoadAllCaches` snapshots `Rows` (master texts, overrides, inclusion flags, customer data).
   - Chapter buttons default to chapter 1.
3. For the selected chapter:
   - Filters `Rows` by `rowId` prefix (e.g., `1.`) and uses `IsLevel3Leaf` to present Level-3 items.
   - Renders one frame per finding using templates cloned on the fly.
4. Per frame (`CRowUI`):
   - Shows master finding and current level text.
   - Displays customer answer.
   - Ticks "Bericht" check box if `includeFinding` is true.
   - Pre-populates override checkboxes/text from `modReportData` cache.
   - When consultant toggles overrides or levels, `CRowUI` writes back through `modReportData.SaveOverride` / `SaveSBState`, which update `Rows` (and append to `OverridesHistory` via `modABRowsRepository`).

This keeps the UI and structured sheets in sync without touching legacy sheets.

## 4. JSON export

1. After edits, run `modABProjectExport.ExportProjectJson`.
2. Flow:
   - Runs `ValidateAutoBerichtWorkbook` to ensure required structured sheets + headers are present.
   - Reads each structured sheet (`Meta`, `Chapters`, `Rows`, `Photos`, `Lists`, `OverridesHistory`).
   - Builds the composite `project` dictionary (version, meta, chapters with rows, photos, lists, history).
   - Serialises via `JsonConverter.ConvertToJson` (2-space indent).
   - Writes to the chosen output path.

The resulting `project.json` is the canonical payload to feed the HTML/JS app.

## Supporting utilities

- `modABIdUtils`: shared ID normalisation + parent lookup. Ensures every module talks about the same canonical IDs.
- `modABRowsRepository`: shared row operations (ensure record, set overrides, include flags, history logging).
- `modABTableUtils`: header lookup, upsert helpers, Nz wrappers.
- `modABWorkbookSetup`: sheet creation + header enforcement.
- `modABPhotoConstants`: central list/tag identifiers for the PhotoSorter stack.
- `modABPhotosRepository`: reads/writes the structured `Photos` + `Lists` tables, seeds tags from folder names, and exposes helper functions (`GetButtonList`, `TogglePhotoTag`, etc.).
- `PhotoSorter` / `PhotoSorterForm` / `CPhotoTagButton`: UI and directory helpers that now consume the structured tables directly (chapter/category/training/topic buttons).

## Current open items / caveats

- **Photo workflow polish**: Lists/Photos integration is live; verify the `Lists` sheet reflects the desired button taxonomy (Bericht/Seminar/Topic) before running the PhotoSorter so folder seeding stays in sync.
- **Chapter orchestration**: ensure `Chapters` sheet is populated (either via `LoadProjectJson` or a staging import) before running the Bericht editor to keep button counts consistent.
- **Error handling**: import macros currently stop on unexpected headers — consider upgrading messaging / logging before rollout.
- **Testing**: manual smoke test required after every refresh (run both import macros + open BerichtForm + PhotoSorter) until automated regression coverage exists.

Everything else in the repository (e.g., [Data Model](../architecture/data-model.md)) already describes the JSON contract consumed by the web client.

---

**Related Documentation**:
- [Data Model](../architecture/data-model.md) - JSON schema and structure
- [System Overview](../architecture/system-overview.md) - How components integrate
- [Getting Started Guide](getting-started.md) - End-to-end workflow
- [VBA Modules Reference](../../macros/README.md) - Detailed module documentation
