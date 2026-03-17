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
- Score rule (current): selected level maps directly to percentages
  1=0%, 2=33%, 3=66%, 4=100%.
- Inclusion in report is controlled independently (`include` + `done`);
  excluded findings are not auto-converted to 100%. Field observations have no score.
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
- Tag pills are split actions: left side toggles filter, right side toggles tagging on the current photo.
- Filters are cumulative (AND across active filters).
- A clear-filters control is available in the photo viewer.
- UI locale policy:
  - UI controls/messages remain in English.
  - Instructional hint text is localized by project language (DE/FR/IT).

## 9. Proposed Architecture (Policy-Safe)

Core principle: no unsafe flags, no implicit disk access.

One project folder:
- project_sidecar.json (editor state)
- library_user_*.json (optional reusable knowledge base)
- inputs/
- outputs/
- templates/
- photos/

Responsibilities:
- Browser editor: edit report content, tag photos, autosave to sidecar JSON.
- Web export engine (in app): read sidecar + template and generate Word/PPT.
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
│  Export Engine      │ ───────────────────▶  │  Word/PPT Templates     │
│  (thin layer)        │                       │  (corp branding)        │
└──────────────────────┘                       └────────────────────────┘
           │
           ▼
┌──────────────────────────────────────────────────────────────┐
│ Outputs: Word report, PPT decks, photo folder views            │
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
          └── Word/PPT export → Word/PPT outputs
```

### 9c. Project Folder Layout

```text
<Project>/
  project_sidecar.json      # editor state (canonical)
  library_user_*.json       # optional reusable knowledge base
  inputs/                   # customer inputs (self-assessment + docs)
  outputs/                  # Word/PPT outputs
  templates/                # DOCX/PPTX templates
  photos/
    raw/
      pm1/
      pm2/
      pm3/
    resized/
    export/
```

Access:
- File System Access API with folder picker each session.
- If blocked, the app cannot operate (no fallback by design).

## 9. Photo Workflow

Default: virtual tags stored in sidecar JSON, with fast filtering UI.
Import/rename/resize pipeline is provided in the PhotoSorter UI; tags remain virtual.
Optional: materialize folders via export into `photos/export/`.
Future: `photos/_views/` for disk-based browsing (not implemented).

Tags are many-to-many. No primary tag requirement.
Observation tags can be added by the user and removed via Settings.
Removing a tag clears it from all photos (no orphan tags).

### Photo import/export (PhotoSorter)

Goal: a simple raw -> resized pipeline plus an explicit export that materializes
folders for sharing or import into Office/PPT workflows.

Import rules (PhotoSorter "Import Photos" modal):
- Source: `photos/raw/<owner>/` where `<owner>` is any 3-character folder name
  (default scaffold: `pm1`, `pm2`, `pm3`).
- Accept: all image files.
- Timestamp: prefer EXIF `DateTimeOriginal`; fallback to file creation time if
  available; otherwise use the file timestamp exposed by the browser
  (`lastModified`).
- Filename format: `YYYY-MM-DD-HH-MM_<owner>_0001.jpg` (4-digit sequence).
  Sequence increments per owner to avoid collisions.
- Resize: longest side = 1920px, JPEG quality ~0.85.
- Destination: `photos/resized/` (lowercase) acts as the unsorted reservoir.
- Raw files are never deleted or moved.

Export rules (PhotoSorter "Export" action in the same modal):
- Source: the current photo root (typically `photos/resized/`).
- Destination: `photos/export/` with subfolders:
  - `unsorted/` for photos with no tags in any group.
  - report tags at root level (`<report-tag>/`).
  - observations under localized 4.8 folder:
    - DE: `4.8 Beobachtungen/<tag>/`
    - FR: `4.8 Observations/<tag>/`
    - IT: `4.8 Osservazioni/<tag>/`
  - training under `Training/<tag>/`.
- Photos with multiple tags are copied into each matching folder (duplicates OK).
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

## 11a. Export Path (Web Export Engine, Recommended)

Goal: remove Excel as an orchestration hop while keeping Word in charge of layout.

Phase A (browser)
- The editor writes canonical state to `project_sidecar.json`.
- Export uses the sidecar directly (or an optional lightweight “export.json” view).
- Content is plain text + simple list markers (markdown-lite).

Phase B (web export engine)
- Word template contains one content control per chapter (e.g., `Chapter1`).
- Exporter reads sidecar/export JSON and injects chapter content into the controls.
- Exporter applies styles (Heading 1/2/3, body, tables), converts list markers to
  proper Word lists, inserts section breaks, and updates TOC/fields.
- Exporter is responsible for filtering, renumbering, and computing score values
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
- PowerPoint templates in `<Project>/templates/`

Outputs:
- Report presentation (`YYYY-MM-DD-company-Bericht-Besprechung.pptx`)
- Training deck (`YYYY-MM-DD-company-Seminar-Slides.pptx`) in `outputs/`

Template (current working file)
- Report deck template: `<Project>/templates/Vorlage AutoBericht.pptx`
- Training content is generated from the same template using training layouts.

### Layout naming (in the template)
Report layouts (required):
- `ab_title`
- `ab_chapterorange`
- `ab_titleandpicture`
- `ab_textandpicture`
- `ab_4pictures`
- `ab_6pictures`
- `ab_titleandtext`

Report layouts (optional):
- `ab_2pictures`
- `ab_3pictures`

Training layouts:
- `ab_title` (intro)
- `ab_chapterorange` (tag divider)
- `ab_picture` (default picture layout)
- localized tag layouts:
  - `ab_unterlassen_{d|f|i}`
  - `ab_dulden_{d|f|i}`
  - `ab_handeln_{d|f|i}`
  - `ab_vorbild_{d|f|i}`
  - `ab_verhindern_{d|f|i}`
  - `ab_audit_{d|f|i}`
  - `ab_risikobeurteilung_{d|f|i}`
  - `ab_aviva_{d|f|i}`

### Placeholder expectations
The exporter targets standard placeholders by type:
- `title` for slide title
- `body` (largest non-title text box) for text
- `pic` for image slots

Layout validation is strict: export fails with explicit errors if required layouts
or placeholders are missing.

### Export logic (current implementation)
1) Report deck export (Project page button “PowerPoint Export (Report)”):
   - Template: `templates/Vorlage AutoBericht.pptx`.
   - Output: `outputs/<yyyy-mm-dd>-<company>-Bericht-Besprechung.pptx`.
   - Cover slide uses `ab_title`:
     - title: localized report title
     - body: date + moderator
     - preserves image placeholder(s) for manual user image insertion
   - Chapter 0:
     - separator slide (Management Summary, localized)
     - recommendation text slides in chunks of 3 items (`ab_titleandtext`)
     - additional localized section for assessment summary:
       - separator slide
       - spider chart slide in `ab_titleandpicture`
   - Chapters 1..:
     - separator slide (`ab_chapterorange`)
     - chapter snapshot slide (`ab_titleandpicture`) with generated thermo image
     - For each section (1.1 level):
       - text slides (`ab_textandpicture`) with 3 findings per slide.
         Each line: `1.2.1 Finding text` (renumbered IDs).
       - photo slides with layout selection by chunk size:
         - 1 image -> `ab_titleandpicture`
         - 2 images -> `ab_2pictures` (if present)
         - 3 images -> `ab_3pictures` (if present)
         - up to 4 -> `ab_4pictures`
         - up to 6 -> `ab_6pictures`
   - Renumbering: only included findings count; numbering is re-packed within
     each chapter and used in slide titles and finding lines.
   - Chapter 4.8 (field observations):
     - Display section number is `4.<next>` where `<next>` is the number of
       sections in chapter 4 after renumbering + 1 (e.g., if 4.3 is missing,
       4.8 becomes 4.7).
     - separator slide uses `ab_chapterorange`.
     - each observation row starts with `ab_textandpicture` (finding text + first photo if available)
     - remaining photos spill to photo-only slides using count-based layout selection.

2) Training deck export (Project page button “PowerPoint Export (Training)”):
   - Template: `templates/Vorlage AutoBericht.pptx` (training layouts).
   - Output: `outputs/<yyyy-mm-dd>-<company>-Seminar-Slides.pptx`.
   - Intro slide: `ab_title`.
   - For each training tag with photos (order: known tags first, then others):
     - Insert an `ab_chapterorange` divider slide titled with the tag.
     - Choose layout by tag name:
       - `unterlassen`, `dulden`, `handeln`, `vorbild`, `audit`,
         `risikobeurteilung`, `aviva`, `verhindern` -> `ab_<tag>_{d|f|i}`
       - other tags -> `ab_picture`
     - Count picture placeholders in the layout; batch photos into that many
       slots per slide.
     - Fill the tag name into the title placeholder if present; drop photos
       into picture placeholders in order.
   - Locale suffix mapping: `de -> d`, `fr -> f`, `it -> i`.

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
- Export to Word/PPT is performed locally in the web app (no cloud services).

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
4) Web export engine (content controls/placeholders → styled Word/PPT).
5) Validation, logging, and error handling.

## 16. Open Decisions

- Template governance process:
  - placeholder map ownership,
  - style lock QA checklist before rollout.
- Release regression scope:
  - minimum export matrix per locale/template variant,
  - sidecar compatibility checks for future schema updates.
