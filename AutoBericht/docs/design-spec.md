# AutoReport Redesign Spec (Interview Notes)

Date: 2026-01-04

## 1. Problem Statement

OHS engineers spend days juggling photos, Excel self-assessments, and Word/PPT templates.
The current tool chain is fragmented (folders, Excel macros, Word formatting quirks,
SharePoint sync issues), which creates friction and errors. The goal is a simple,
offline, policy-safe system that keeps all work in one place, minimizes file wrangling,
and produces consistent Word/PPT outputs without installing software.

## 2. Users and Context

- Users: ~15 senior safety engineers.
- Each project has a single lead author; peer review is done by email.
- Environment: locked-down corporate Windows/Edge, high security.
- No installs, no real web, no unsafe browser flags.
- Projects usually live on SharePoint (slow/sync issues); work is kept locally (C:)
  for speed, then archived to a network drive when done.

## 3. Goals

- One simple workspace for photos + self-assessment + report writing.
- No installs, no IT friction, no unsafe browser flags.
- Minimal file wrangling (no worse than today).
- Stable Word/PPT outputs from templates; one-click export.
- A per-engineer recommendation library that grows over time.
- Reduce report time from ~4-5 days to ~3 days.

## 4. Non-Goals

- Changing corporate IT policies or security posture.
- Real-time multi-user collaboration.
- Online cloud services.
- Replacing Word/PPT templates or report structure.

## 5. Inputs and Outputs

Inputs:
- Customer self-assessment Excel (fixed question IDs, ~200 questions).
- Customer documents (policies, evidence).
- Interview notes.
- Photos from site visit.

Outputs:
- Word report (formatted via templates).
- Report presentation (PPT).
- Draft PDF for customer review.
- Training slide pack for later training sessions.

## 6. Domain Notes and Rules

- Question IDs are stable (e.g., 1.1.1). Some questions are "tiroir" sub-items
  (e.g., 5.1.1a/b/c) used for self-assessment, but consolidated into one report line
  (5.1.1).
- Chapter 4.8 contains freeform field observations tied to photos and topics.
- Field observations appear as their own chapter entry in the left sidebar
  (e.g., 4.8 Beobachtungen) and are handled separately from standard findings.
- Report includes only improvement opportunities; positive findings are omitted.
  Remaining items are renumbered within each chapter.
- Self-assessment answers are binary; consultants adjust to percentage scores.
  These feed a spider chart in the report.
- Score rule (derived): selected level 1=0%, 2=25%, 3=50%, 4=75%.
  If a finding is excluded, treat it as 100%. Field observations have no score.
- Reports are single-language (DE/FR/IT), templates are fixed but updated over time.

## 7. Data Model (Draft)

Question
- id (e.g., 1.1.1, 5.1.1a)
- chapter (1-10)
- text_original
- type (standard | tiroir | field_observation)
- parent_id (for tiroir sub-items)

SelfAssessment
- question_id
- company_answer
- company_evidence

ReportItem (what ends up in the report)
- question_id (consolidated line, e.g., 5.1.1)
- type (standard | field_observation)
- finding_text (reworded negative; softened/qualified)
- consultant_score_pct (derived; see score rule above)
- recommendation_text (single editable block)
- include_in_report (bool)

Photo
- id
- filename
- tags (many-to-many: report chapter, observation topic, training)
- notes

RecommendationLibrary (per engineer)
- question_id
- recommendation_text (single block)
- last_used

## 8. UX Notes (Photo Sorter)

- Implemented layout: single-page, scroll-free tagging workflow sized for 1920×1080 @125%.
- Tag panes use uniform buttons with tooltips; photo preview and tagging controls stay visible.
- Two-layout toggle is NOT used; current design keeps one layout (decision: single layout sufficient).

## 9. Proposed Architecture (Policy-Safe)

Core principle: no unsafe flags, no implicit disk access.

One project folder:
- project_sidecar.json (editor state)
- project_db.xlsx (optional readable archive)
- self_assessment.xlsx
- photos/
- out/
- cache/ (optional materialized views)

Responsibilities:
- Browser editor: edit report content, tag photos, autosave to sidecar JSON.
- Word template macros: read sidecar/export JSON -> generate Word/PPT/PDF.
- Templates: updated centrally; exports use latest templates.

### 9a. System Diagram

```
┌──────────────────────────────────────────────────────────────┐
│                          AutoReport                           │
└──────────────────────────────────────────────────────────────┘

┌──────────────────────┐      reads/writes     ┌────────────────────────┐
│  Minimal Editor UI   │ ───────────────────▶  │  project_sidecar.json  │
│  (browser, offline)  │ ◀───────────────────  │  (working state)       │
└──────────────────────┘                       └────────────────────────┘
           │
           │ optional
           ▼
┌──────────────────────┐     reads sidecar     ┌────────────────────────┐
│  Word/VBA Exporter  │ ───────────────────▶  │  Word/PPT Templates     │
│  (thin layer)        │                       │  (corp branding)        │
└──────────────────────┘                       └────────────────────────┘
           │
           ▼
┌──────────────────────────────────────────────────────────────┐
│ Outputs: Word report, PPT deck, PDF, photo folder views       │
└──────────────────────────────────────────────────────────────┘
```

### 9b. Data Flow Diagram

```
Self-assessment Excel + Photos
          │
          ▼
Minimal Editor (chapter editing, recommendations, tags)
          │
          ▼
project_sidecar.json (canonical state)
          │
          └── Word/PPT VBA export → Word/PPT/PDF outputs
```

### 8c. Project Folder Layout

```
<Project>/
  AutoBericht/              # app bundle + seeds + docs
  project_sidecar.json      # editor state (canonical)
  project_db.xlsx           # optional readable archive (future)
  Inputs/                   # customer inputs (self-assessment + docs)
  photos/                   # raw photos (pm1/pm2), resized, export
  Outputs/                  # Word/PPT/PDF outputs
  Photos/export/            # materialized tag folders (current export)
```

Access:
- File System Access API with folder picker each session.
- If blocked, the app cannot operate (no fallback by design).

## 9. Photo Workflow

Default: virtual tags stored in sidecar JSON, with fast filtering UI.
Import/rename/resize pipeline is provided in the PhotoSorter UI; tags remain virtual.
Optional: materialize folders via export into `Photos/export/`.
Future: `Photos/_views/` for disk-based browsing (not implemented).

Tags are many-to-many. No primary tag requirement.
Observation tags can be added by the user and removed via Settings.
Removing a tag clears it from all photos (no orphan tags).

### Photo import/export (PhotoSorter)

Goal: a simple raw -> resized pipeline plus an explicit export that materializes
folders for sharing or import into Office/PPT workflows.

Import rules (PhotoSorter "Import Photos" modal):
- Source: `Photos/raw/<pmCode>/` where `<pmCode>` is a 3-letter PM code (lowercase).
- Accept: all image files.
- Timestamp: prefer EXIF `DateTimeOriginal`; fallback to file creation time if
  available; otherwise use the file timestamp exposed by the browser
  (`lastModified`).
- Filename format: `YYYY-MM-DD-HH-MM-<pmCode>-0001.jpg` (4-digit sequence).
  Sequence increments per (timestamp, pmCode) to avoid collisions.
- Resize: longest side = 1200px, JPEG quality ~0.85.
- Destination: `Photos/resized/` (lowercase) acts as the unsorted reservoir.
- Raw files are never deleted or moved.

Export rules (PhotoSorter "Export" action in the same modal):
- Source: the current photo root (typically `Photos/resized/`).
- Destination: `Photos/export/` with subfolders:
  - `unsorted/` for photos with no tags in any group.
  - `<group>/<tag>/` where `<group>` is `report`, `observations`, or `training`.
- Photos with multiple tags are copied into each group/tag folder (duplicates OK).
- Tag folder names are sanitized for filesystem safety (no slashes, reserved
  characters).
- Existing files are not overwritten; append a numeric suffix if needed.

UI:
- The top "Import Photos" button opens a modal overlay with only Import and
  Export actions plus status/progress text.

## 10. Editor UX (Draft)

- One chapter per page.
- Filters to show/hide:
  - company answered yes / 100%
  - included in report
  - done vs open
- For each item:
  - find finding_text (reworded negative)
  - score percentage
  - recommendation text (single editable block; save to report or library)
  - photo attachment via section tags only (no per-finding photo linkage)
- Minimal formatting controls (markdown-lite: bold, bullets, links).
- Word templates handle final formatting.
- Optional preview panel to approximate Word layout.

## 11. Export Rules

- Report contains only improvement opportunities.
- Items are renumbered within each chapter.
- Field observations are exported as their own chapter (e.g., 4.8 Beobachtungen).
- Spider chart uses consultant scores; company answers retained in data.

## 11a. Export Path (Word Macro, Recommended)

Goal: remove Excel as an orchestration hop while keeping Word in charge of layout.

Phase A (browser)
- The editor writes canonical state to `project_sidecar.json`.
- Export uses the sidecar directly (or an optional lightweight “export.json” view).
- Content is plain text + simple list markers (markdown-lite).

Phase B (Word macro)
- Word template contains one content control per chapter (e.g., `Chapter1`).
- Optional: embed a RibbonX tab (“AutoBericht”) in the template for macro buttons.
- Macro reads sidecar/export JSON and injects chapter content into the controls.
- Macro applies styles (Heading 1/2/3, body, tables), converts list markers to
  proper Word lists, inserts section breaks, and updates TOC/fields.
- Macro is responsible for filtering, renumbering, and computing score values
  (no separate export file required).

Benefits
- Word remains the single source of formatting truth.
- No unsafe browser flags and no Excel dependency for export.
- Templates can be updated independently by corporate design.
- Chapter tables: inserted at 100% width with percent-based columns (35/58/7) set before any merges; table wrapping enabled (WrapAroundText). Header merges happen after sizing, and we avoid AutoFit-to-contents to keep column 3 stable.

## 11b. PowerPoint Export

Goal: generate two slide decks from the sidecar JSON.

Inputs:
- `project_sidecar.json` (report + photo tags)
- PowerPoint templates in the project root (or `<Project>/<Templates>/` when
  `AB_PPT_TEMPLATE_FOLDER` is set in the VBA config)
- Active Word report (used to capture chapter screenshots)

Outputs:
- Report presentation (`YYYY-MM-DD_Bericht_Besprechung.pptx`)
- Training deck (`YYYY-MM-DD_Seminar_Slides_D/F.pptx`) in the project folder

Template (current working file)
- Report deck template: `<Project>/Vorlage AutoBericht.pptx`
  (`AB_PPT_REPORT_TEMPLATE` in `AutoBericht/vba/autobericht_config.bas`)
- Training decks: `<Project>/Training_D.pptx` and `<Project>/Training_F.pptx`
- Seed file: `test/AutoBericht_slides.pptx` (layout reference)

### Layout naming (in the template)
Report layouts:
- `report_title`
- `report_chapter_separator`
- `report_chapter_screenshot`
- `report_section_text`
- `report_section_photo_3`
- `report_section_photo_6`
- `report_section_48_separator`
- `report_section_48_text_photo_3`
- `report_section_48_text_photo_6`

Training layouts:
- `chapterorange` (tag divider)
- `picture` (for Iceberg / Pyramide / STOP / SOS)
- German `*_d`:
  - `seminar_d` (title slide)
  - `unterlassen_d`
  - `dulden_d`
  - `handeln_d`
  - `verhindern_d`
  - `audit_d`
  - `risikobeurteilung_d`
  - `aviva_d`
  - `vorbild_d`
- French `*_f`: same names with `_f`

### Placeholder expectations
The macro targets standard placeholders by type:
- `title` for slide title
- `body` (largest non-title text box) for text
- `pic` for image slots (3/6 grids and screenshot)

### Export logic (current implementation)
1) Report deck export (Word ribbon “Bericht Besprechung”):
   - Template: `AB_PPT_REPORT_TEMPLATE` (default `Vorlage AutoBericht.pptx`).
   - Output: `<project>/<yyyy-mm-dd>_Bericht_Besprechung.pptx`.
   - Optional title slide: uses `report_title` and `meta.projectName` /
     `meta.company` (fallback: "Report").
   - Chapter 0:
     - Separator slide titled “Management Summary”.
     - Recommendation slides: 3 per slide (blank line between), no screenshot.
   - Chapters 1..:
     - Separator slide titled `1. <chapter title>`.
     - Screenshot slide of the first page of the chapter from Word
       (bookmark `Chapter<id>_start`, dots replaced with `_`).
     - For each section (1.1 level):
       - Text slides (`report_section_text`) with 3 findings per slide.
         Each line: `1.2.1 Finding text` (renumbered IDs).
       - Photo slides (`report_section_photo_3` or `_6`) with photos tagged
         to the section id.
   - Renumbering: only included findings count; numbering is re-packed within
     each chapter and used in slide titles and finding lines.
   - Chapter 4.8 (field observations):
     - Display section number is `4.<next>` where `<next>` is the number of
       sections in chapter 4 after renumbering + 1 (e.g., if 4.3 is missing,
       4.8 becomes 4.7).
     - Separator slide uses `report_section_48_separator`.
     - Each observation row becomes a slide titled `4.x.y <tag label>` with
       text + photos (`report_section_48_text_photo_3` or `_6`).
     - Additional photos spill to photo-only slides
       (`report_section_photo_3/6`) with the same title; no text.

2) Training deck export (Word ribbon buttons “VG Seminar D/F”):
   - Templates: `<project>/Training_D.pptx` and `Training_F.pptx`.
   - For each training tag with photos (order: known tags first, then others):
     - Insert a `chapterorange` divider slide titled with the tag.
     - Choose layout by tag name:
       - `unterlassen`, `dulden`, `handeln`, `vorbild`, `audit`,
         `risikobeurteilung`, `aviva`, `verhindern` → `<tag>_d` or `<tag>_f`
       - `iceberg`, `pyramide`, `stop`, `sos`, `stgb art. 230`, unknown →
         `picture`
     - Count picture placeholders in the layout; batch photos into that many
       slots per slide (e.g., AVIVA with 6 placeholders packs up to 6 per slide).
     - Fill the tag name into the title placeholder if present; drop photos
       into picture placeholders in order.
   - Optional seminar title slide: uses `seminar_d` if present and enabled in
     VBA config.
   - Output: `<project>/<yyyy-mm-dd>_Seminar_Slides_D.pptx` (or `_F`).

### Training tag alignment
Training tag names in the knowledge base must match the layout mapping above (case-insensitive).

## 12. Maintainability

- project_db.xlsx is a future optional archive (not generated by the app today).
- project_sidecar.json is the working state (portable, inspectable).
- Recommendation library stored per engineer in a user folder; optional sharing.

Seed and library resolution (fresh projects):
- Prefer existing `project_sidecar.json` if present.
- If no sidecar, load a user **knowledge base** in project root
  (e.g., `library_user_XX_de-CH.json`) and use it **as-is**.
- If multiple user knowledge bases are present, prompt the user to choose one.
- If no user knowledge base, load bundled seed knowledge base from
  `AutoBericht/data/seed/knowledge_base_*.json` (single source of defaults: tags, structure, content).
- If neither exists, the app initializes an empty project and prompts the user.

### 12a. Current Sidecar Schema (as implemented; stored in project folder)

```
project_sidecar.json
├─ meta
├─ report
│  └─ project
│     ├─ meta { locale, author, createdAt, ... }
│     └─ chapters[]
│        ├─ id: "1", "4.8", "11", ...
│        ├─ title: { de: "Leitbild, Sicherheitsziele ..." }
│        └─ rows[]
│           ├─ kind: "section" (for 1.1 / 1.2 headers)
│           │  └─ id / title
│           └─ item rows
│              ├─ id: "1.1.1" (collapsed ID)
│              ├─ sectionId / sectionLabel
│              ├─ type: "standard" | "field_observation"
│              ├─ master { finding, recommendation }
│              ├─ customer { answer, remark, items[] }
│              └─ workstate { selectedLevel, include/done, findingText, recommendationText, libraryAction, libraryHash }
└─ photos
   ├─ photoRoot
   ├─ photoTagOptions { report, observations, training }
   └─ photos{ path → { tags, notes } }
```

### 12b. Knowledge Base Schema (seed + user library)

```
knowledge_base_*.json   (seed + user library share the same schema)
├─ schemaVersion
├─ meta { locale, updatedAt, author?, source? }
├─ structure
│  └─ items[] (self-assessment questions + grouping metadata)
│     ├─ id / groupId / collapsedId
│     ├─ chapter / chapterLabel
│     ├─ sectionLabel
│     └─ question
├─ library
│  └─ entries[]
│     ├─ id
│     ├─ finding
│     ├─ recommendation
│     └─ lastUsed?
└─ tags
   ├─ observations[]
   ├─ training[]
   └─ report[] (optional; can be derived from structure)
```

Notes:
- **IDs carry the numbers**; text fields should be human text only.
- **Display labels** are derived as `id + title` in the UI.
- **Chapter 4.8** is a special `field_observation` chapter, with rows generated
  from observation tags.
- **Observation row IDs** are numeric (e.g., `4.8.1`, `4.8.2`) even when derived
  from tags; the tag label is stored on the row (`tag` / `titleOverride`).
- **Chapter 2 renumbering:** during seed build, the 4‑level IDs in chapter 2 are
  shifted up so the report stays at 3 levels. The original self‑assessment ID is
  preserved as `originalId` for import mapping.

## 13. Risks and Mitigations

- File System Access API blocked: app cannot operate (no fallback by design).
- SharePoint sync issues: work locally, archive later.
- Template updates: strict mapping rules and warnings.
- Large photo sets: lazy indexing and thumbnail caching.

## 14. Security & Data Handling (IT Review)

Design goal: keep all customer data local, with explicit user consent for any file access.

### Data locality
- No cloud storage, no external API calls, no telemetry.
- All data stays in the project folder on disk.
- Editor state lives in `project_sidecar.json` (human-readable JSON).

### File access model
- Uses File System Access API with explicit folder picker each session.
- Browser cannot access disk by path; access is capability-based only.

### Network exposure
- App is served locally (e.g., `http://localhost`) or `file://` in a locked-down browser.
- No outbound connections are required for normal operation.
- Libraries (e.g., SheetJS) are bundled locally, not loaded from CDNs.

### Data scope and mutation rules
- AutoBericht only reads/writes `project_sidecar.json`.
- PhotoSorter only reads images and writes tags/notes into `project_sidecar.json`.
- No files are moved, deleted, or rewritten by the browser.
- Export to Word/PPT/PDF is performed locally via Office/VBA (no cloud services).

### Frontend module layout
- `AutoBericht/mini/app.js` is the thin orchestrator (wiring only).
- Shared modules live in `AutoBericht/mini/shared/`:
  - `state.js` (default project model + helpers)
  - `normalize.js` (workstate/meta normalization + observation helpers)
  - `seeds.js` (seed + library loading/building)
  - `io-sidecar.js` (sidecar save/load + library generation)
  - `import-self.js` (Excel self-assessment import)
  - `render.js` (UI rendering + overlays)
  - `elements.js` (DOM element lookup)
  - `bind-events.js` (event wiring)
- This keeps the surface area small and makes future localization/feature work safer.
- PhotoSorter mirrors this layout with modules in `AutoBericht/mini/photosorter/` and a thin `AutoBericht/mini/photosorter.js` orchestrator.

### Checklist data source
- Checklist lists are JSON files in `AutoBericht/data/checklists/`:
  - `checklists_de.json`, `checklists_fr.json`, `checklists_it.json`
- Loaded at runtime by the editor from `http://` or `https://` context.
- No JS fallback; checklist overlay requires HTTP hosting.

### Security assumptions
- Endpoint security and disk encryption are managed by IT policy.
- If the browser profile is compromised, local data could be exposed (same as any local file).
- Folder access can be revoked by clearing site permissions/data.

### Compliance notes
- The system can operate fully offline.
- No special browser flags or security bypasses are required.

## 15. Phased Plan (Draft)

1) Project folder + sidecar schema + minimal editor load/save.
2) Chapter editor flow + recommendation library integration.
3) Photo tagging + filtering + optional materialize.
4) Word macro export (content controls → styled Word/PPT/PDF).
5) Validation, logging, and error handling.

## 16. Open Decisions

- Confirm File System Access API availability in locked-down Edge.
- Decide final Word macro contract (content controls + JSON schema).
- Decide default library storage location (user folder vs project).
