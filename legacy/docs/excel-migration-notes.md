# Legacy Workbook Structure & JSON Alignment Plan

### Overview
Files extracted from `fromWork/drive-download-20251021T092517Z-1-001.zip` show the current Excel-based workflow:

- `2025-07-10 AutoBericht v0 Berichtform report selectors.xlsx` — workbook without macros, used here to inspect sheet layout.
- `2025-07-10 AutoBericht v0 Berichtform report selectors.xlsm` — macro-enabled version (PhotoSorter UI, import/export macros).
- `2025-10-02_ExportedVBA.txt` — text dump of the VBA modules (PhotoSorter form, folder operations, etc.).
- Additional Word/Excel files (`IST-Aufnahme…`, `Selbstbeurteilung…`) feed the master and self-evaluation content.

The XLSX exposes how data is currently organised before being consumed by VBA and the legacy UI.

### Sheet Summary

| Sheet | Purpose | Key Fields Observed |
| --- | --- | --- |
| `Start` | User instructions for the macro workflow. | Step-by-step description (import Selbstbeurteilung, import MyMaster, sort photos, edit report). |
| `Settings` | Reserved | Empty in the sample. |
| `SelbstbeurteilungKunde` | Raw customer self-evaluation. | Columns: `ReportItemID`, section headers, question text, `AntwortKunde` (numeric), `Bemerkungen des Kunden`, weighting columns, `Selected`, `SelectedLevel`. Hierarchical rows (`1`, `1.1`, `1.1.1`, …). |
| `MyMaster` | Master findings and per-level recommendations. | Columns: `ReportItemID`, `Feststellung`, `Level1`–`Level4`. Content matches the master JSON we reviewed earlier. |
| `ReportOverrides` | Consultant edits captured by VBA UI. | Columns: `ReportItemID`, `Component` (`Finding` / `Recommendation`), `LevelNum`, `OverrideText`, `Status`, `LastEdited`, `PromotedToMaster`. |
| `OutputChapterMapping` | Placeholder for renumbered output. | Only header row: `ReportItemID`, `ChapterItemID`, `HasFindings`, `HasPictures`. |
| `PSHelperSheet` | Working table for PhotoSorter tagging. | Column A = file name (e.g., `DSC00700.JPG`); subsequent columns correspond to chapter buttons. Cell value `1` marks the assignment. |
| `PSCategoryLabels` | Labels for PhotoSorter button groups. | Columns hold headings for Bericht buttons, Audit/VGSeminar buttons, and topic labels. Each row aligns to a chapter (same order as `BerichtKapitel`). |
| `BerichtKapitel` | Chapter list. | Column A = chapter ID (`1.1.`), Column B = German title. Used for menus. |

The VBA dump confirms PhotoSorter uses `PSHelperSheet`/`PSCategoryLabels`, handles folder creation, and writes override records into `ReportOverrides`.

### Mapping to Proposed JSON

| JSON Section | XLSX Source/Target | Comment |
| --- | --- | --- |
| `meta` | New `Meta` sheet (or reuse `Settings`). | Store `projectId`, `company`, `createdAt`, `locale`, `author`. |
| `chapters[].rows[]` | Combination of `BerichtKapitel`, `MyMaster`, `SelbstbeurteilungKunde`, `ReportOverrides`. | `master` fields from `MyMaster`, `customer` fields from `SelbstbeurteilungKunde`, `workstate` from `ReportOverrides`, `title`/structure from `BerichtKapitel`. |
| `photos` | `PSHelperSheet` + `PSCategoryLabels`. | Convert column flags into `tags.bericht`, `tags.topic`, `tags.seminar`. |
| `lists` | `PSCategoryLabels` (`berichtList`, `seminarList`, optional button definitions). | Provide PhotoSorter button metadata to the browser UI. |
| `exportHints` | Could align with `OutputChapterMapping` once renumber logic exists. |

### Recommendation

1. **Introduce structured sheets** that match JSON directly:
   - `Meta` — two-column key/value list.
   - `Rows` — flat table keyed by `rowId`, carrying all row fields (`title`, `masterFinding`, `level1..4`, `selectedLevel`, overrides, include flags, etc.).
   - `Photos` — filename + notes + tag columns (`bericht`, `seminar`, `topic`).
   - `Lists` — optional table for PhotoSorter button definitions.

   VBA can continue to read/write legacy sheets but should synchronise them with these structured tables so JSON export/import is straightforward.

2. **Adopt JSON as canonical** for AutoBericht, with the `.xlsm` acting as a “friendly editor”. Workflow:
   - VBA imports master/self-evaluation → fills structured sheets → exports `project.json`.
   - AutoBericht loads `project.json` or the `.xlsm` (via SheetJS) → UI editing.
   - On save/export, AutoBericht rewrites structured sheets and regenerates `project.json`. Users open Excel only when the web app is closed, preventing conflicts.

3. **SheetJS handling** (when the browser writes `.xlsm`):
   ```js
   const wb = XLSX.read(arrayBuffer, { bookVBA: true });
   // mutate structured sheets (Meta, Rows, Photos, Lists, OutputChapterMapping…)
   const bytes = XLSX.write(wb, { bookType: 'xlsm', type: 'array', bookSST: true });
   ```
   Ensure the macro bundle (`wb.vbaraw`) is preserved.

4. **Chapter 4.8**: generate additional rows in the `Rows` sheet with IDs like `4.8.custom-001`. Excel macros can copy them into `ReportOverrides`/legacy views if needed. Export renumbers them (e.g., `4.8.1`, `4.8.2`).

This setup keeps everything in a single `.xlsm`, remains approachable for non-technical staff, and aligns with the richer JSON model used by AutoBericht and automation tooling.

### Recommendation (Blank-Slate Workbook Layout)

Rather than synchronising legacy helper sheets, we can reshape the `.xlsm` so each worksheet directly mirrors the unified JSON schema. Existing macros/forms can be refactored to target these tables, keeping the workflow approachable for non-technical users while simplifying browser read/write.

1. **`Meta`** — key/value table.
   | key | value |
   | --- | --- |
   | projectId | 2025-ACME-001 |
   | company | ACME AG |
   | locale | de-CH |
   | createdAt | 2025-02-11T08:45:00Z |
   | author | Consultant Name |

2. **`Chapters`** — hierarchy with localisation.
   | chapterId | parentId | orderIndex | defaultTitle_de | defaultTitle_fr | defaultTitle_it | pageSize |
   | --- | --- | --- | --- | --- | --- | --- |
   | 1 |  | 1 | Leitbild, Sicherheitsziele … | Politique de sécurité … | … | 5 |
   | 1.1 | 1 | 1 | Unternehmensleitbild … | Politique d’entreprise … | … | 5 |
   | 1.1.1 | 1.1 | 1 | Verfügt das Unternehmen … | L’entreprise dispose-t-elle … | … | null |

   `parentId` and `orderIndex` define the tree; extra `title_{locale}` columns support future languages.

3. **`Rows`** — flattened row data.
   | rowId | chapterId | titleOverride | masterFinding | masterLevel1 | masterLevel2 | masterLevel3 | masterLevel4 | customerAnswer | customerRemark | customerPriority | selectedLevel | findingOverride | recommendationOverride | includeFinding | includeRecommendation | overwriteMode | done | notes | lastEditedBy | lastEditedAt |

   `chapterId` references `Chapters.chapterId`. Add new override columns at the end as needed.

4. **`Photos`** — PhotoSorter catalogue.
| fileName | displayName | notes | tagBericht | tagSeminar | tagTopic | preferredLocale |

   Comma-separated lists in `tag*` columns hold canonical IDs (e.g., `1.2,4.8.custom-001`). Both VBA and the web app split/join these values.

5. **`Lists`** — button/tag vocabularies.
   | listName | value | label_de | label_fr | label_it | label_en | group | sortOrder | chapterId |

   Rows with `listName = "photo.bericht"` feed the Bericht pane; `group` chooses the PhotoSorter column (Bericht, Seminar, Topic). Matching `value` to `Chapters.chapterId` lets us display localised titles automatically.

6. **Optional audit tables** — `OverridesHistory` (append-only log) and `ExportLog` (renumber map per export) if auditing is required.

### Workflow with the New Layout
1. Excel macros import Word/self-eval content directly into `Chapters`, `Rows`, `Photos`, `Lists`.
2. PhotoSorter form renders button labels from `Lists` and stores selections in `Photos`.
3. Report editing macros manipulate the `Rows` table for overrides, inclusion flags, levels, and done status.
4. AutoBericht loads the `.xlsm` (using SheetJS with `bookVBA: true`) and maps each sheet into the unified in-memory structure.
5. On save/export, AutoBericht rewrites the same sheets and emits `project.json`; Excel macros can then move photos or generate Word/PPT by reading `Rows` + `ExportLog`.

This blank-slate design keeps everything in one workbook, supports localisation, and aligns perfectly with the JSON schema that the web interface consumes.
