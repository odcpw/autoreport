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
- Chapter 4.6 contains freeform field observations tied to photos and topics.
- Report includes only improvement opportunities; positive findings are omitted.
  Remaining items are renumbered within each chapter.
- Self-assessment answers are binary; consultants adjust to percentage scores.
  These feed a spider chart in the report.
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
- finding_text (reworded negative; softened/qualified)
- consultant_score_pct
- recommendations (list)
- photo_refs (list)
- include_in_report (bool)

FieldObservation (Chapter 4.6)
- topic
- finding_text
- recommendations
- photo_refs
- chapter = 4.6

Photo
- id
- filename
- tags (many-to-many: report chapter, observation topic, training)
- notes

RecommendationLibrary (per engineer)
- question_id
- level (1-4)
- bullets (variant phrasing)
- last_used

## 8. UX Requirements (Photo Sorter)

- Target screen: 1920×1080 @ 125% scaling (≈1536×864 effective).
- Entire workflow must fit on one screen: no page scrolling.
- No internal scrollbars in tag panes (Report / Observations / Training).
- Photo preview and tagging controls must be visible simultaneously.
- Tag buttons should be uniform width, dense grid, and single-click.
- Button labels may be abbreviated to fit; full text shown on hover (tooltip).
- Provide two layout variants:
  - **Stacked panels** (all three panes visible at once).
  - **Tabbed panels** (1/2/3 keys to switch panes).
- User can toggle between the two layouts inside the UI.

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
- Excel macros: import sidecar -> update Excel -> generate Word/PPT/PDF.
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
│  Excel/VBA Exporter  │ ───────────────────▶  │  Word/PPT Templates     │
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
          ├── Optional: write project_db.xlsx (human-readable)
          │
          └── Excel/VBA export → Word/PPT/PDF outputs
```

### 8c. Project Folder Layout

```
<Project>/
  project_sidecar.json      # editor state (canonical)
  project_db.xlsx           # optional readable archive
  self_assessment.xlsx      # customer input (static)
  photos/                   # raw photos
  out/                      # Word/PPT/PDF outputs
  cache/                    # optional materialized photo views
```

Access:
- File System Access API with folder picker each session.
- If blocked, fallback to manual import/export (zip or folder copy).

## 9. Photo Workflow

Default: virtual tags stored in sidecar JSON, with fast filtering UI.
Optional: materialize folders for users who want disk views.
- cache/photos_by_topic/
- cache/photos_by_chapter/
- cache/photos_by_training/

Tags are many-to-many. No primary tag requirement.

## 10. Editor UX (Draft)

- One chapter per page.
- Filters to show/hide:
  - company answered yes / 100%
  - included in report
  - done vs open
- For each item:
  - find finding_text (reworded negative)
  - score percentage
  - recommendations (4 levels; switch, edit, save to report or library)
  - photo attachment via tags or direct selection
- Minimal formatting controls (markdown-lite: bold, bullets, links).
- Word templates handle final formatting.
- Optional preview panel to approximate Word layout.

## 11. Export Rules

- Report contains only improvement opportunities.
- Items are renumbered within each chapter.
- Chapter 4.6 appended as field observations.
- Spider chart uses consultant scores; company answers retained in data.

## 12. Maintainability

- project_db.xlsx is the human-readable fallback.
- project_sidecar.json is the working state (portable, inspectable).
- Recommendation library stored per engineer in a user folder; optional sharing.

## 13. Risks and Mitigations

- File System Access API blocked: provide import/export fallback.
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
- Optional convenience: the last picked folder handle is stored in browser IndexedDB.
  - This is local to the browser profile, not transmitted anywhere.
  - User can clear site data to revoke it.

### Network exposure
- App is served locally (e.g., `http://localhost`) or `file://` in a locked-down browser.
- No outbound connections are required for normal operation.
- Libraries (e.g., SheetJS) are bundled locally, not loaded from CDNs.

### Data scope and mutation rules
- AutoBericht only reads/writes `project_sidecar.json`.
- PhotoSorter only reads images and writes tags/notes into `project_sidecar.json`.
- No files are moved, deleted, or rewritten by the browser.
- Export to Word/PPT/PDF is performed locally via Office/VBA (no cloud services).

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
4) Excel macro orchestration (import sidecar, export Word/PPT/PDF).
5) Validation, logging, and error handling.

## 16. Open Decisions

- Confirm File System Access API availability in locked-down Edge.
- Decide how much export logic stays in Excel vs browser.
- Decide default library storage location (user folder vs project).
