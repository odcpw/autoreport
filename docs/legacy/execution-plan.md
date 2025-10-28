# AutoBericht Offline MVP — Execution Plan

This plan translates the final offline brief into concrete work packages for delivering a first working HTML/JS prototype. Tasks are grouped into sequential phases with clear deliverables so a small team can work in parallel when possible.

## Phase 0 — Foundations (Week 0)
- Finalise project structure under `/AutoBericht/` with `index.html`, `css/`, `js/`, `libs/`, `assets/`, and sample data in `fixtures/`.
- Document coding standards (ES modules, lint rules, formatting) and decide on build tooling (minimal script to fetch and checksum third-party libs, no bundler initially).
- Collect and pin library versions for markdown-it, CodeMirror 6, AJV, Paged.js, PptxGenJS, pica, hotkeys.js, Papa Parse; add manual offline copies into `/libs/`.
- Draft developer handbook detailing how to run the app offline (open `index.html` directly) and how to refresh libraries safely.

## Phase 1 — Data & State Layer (Week 1)
- Implement schema definitions for `master.json`, `self_eval.json`, and working `project.json` in a dedicated `js/schema/` module.
- Build merge logic that ingests master + self-eval into an in-memory `ProjectState` object with change tracking and undo primitives.
- Create persistence helpers for importing/exporting JSON using browser file pickers (upload/download only) and mock hooks for future File-System-Access API.
- Wire AJV validation to run on every state mutation and queue human-readable errors for the Settings tab.

## Phase 2 — UI Shell & Navigation (Week 1–2)
- Lay out the single-page shell with three tabs (PhotoSorter, AutoBericht, Settings) using semantic HTML and CSS grid/flexbox.
- Implement global keyboard focus management to satisfy Tab, Shift+Tab, arrow-navigation, and space/enter behaviours.
- Integrate hotkeys.js for cross-tab shortcuts (e.g., quick switcher or focus commands).
- Establish shared UI components: breadcrumb header, status bar showing validation/export readiness, file upload widgets, and toast notifications.

## Phase 3 — Feature Workstreams (Week 2–4)

**3A. PhotoSorter**
- Build lazy-loaded thumbnail grid with virtual scrolling; load full-size preview on demand.
- Implement metadata editor: notes textarea plus three toggle rows (categories/chapters/training) driven by `ProjectState.lists`.
- Persist changes back into `project.photos` and reflect validation errors for unknown tags.
- Add keyboard support for moving between thumbnails and toggles; include basic image zoom and rotation controls if feasible.

**3B. AutoBericht Editor**
- Render tree view (H1/H2/H3) using master IDs; show client responses inline.
- For each finding, provide dual Markdown editors (CodeMirror 6) for finding and recommendation text with `useAdjusted` toggles and level/inclusion selectors.
- Implement side-by-side preview pane using markdown-it + DOMPurify.
- Track proposed master updates and ensure state synchronises with tree selection.

**3C. Settings & Exports**
- Build upload section for `master.json`, `self_eval.json`, and photo directory (directory read-only via input type="file" webkitdirectory).
- Add branding picker with logo previews and persistence to `project.branding`.
- Implement PDF configuration form (margins, headers/footers toggle, locale selector) and hook validation to disable export buttons until sane.
- Wire export buttons: JSON snapshot, PptxGenJS templates for report/training decks, and Paged.js pipeline for PDF (print-to-file instructions for the tester).

## Phase 4 — Output Pipelines (Week 3–5)
- PDF: build printable HTML template matching layout rules (TOC, headers/footers, first-page logos) with configurable margins; integrate Paged.js runtime.
- PPTX report: implement slide composer that maps findings to slides, converts Markdown to text, lays out images using defined grids, and downscales with pica.
- PPTX training: generate per-training decks grouped by chapter; reuse slide grid logic; ensure file-naming convention is applied consistently.
- Ensure exports respect locale selection (date formatting, UI strings via simple i18n dictionary).

## Phase 5 — Quality, Packaging, and Handover (Week 5)
- Execute acceptance tests 1–7 from the brief with fixture data (≈10 photos) and document results.
- Add automated regression checks: unit tests for merge/validation logic (via playwright-test or vanilla mocha run in headless browser), and visual smoke test scripts if time permits.
- Produce user documentation: quick-start PDF, keyboard cheat sheet, and troubleshooting guide for offline use.
- Package deliverable folder (`AutoBericht/`) with README, changelog, checksum manifest for `/libs/`, and instructions for loading in Edge/Chrome.

## Cross-Cutting Concerns
- **Performance:** benchmark initial load with 700 photo stubs; implement chunked metadata parsing and indexed lookups to keep state operations <50 ms.
- **Accessibility:** ensure focus outlines, ARIA roles for tree view, and high-contrast theme toggle.
- **Security:** sanitise Markdown rendering, avoid `eval`, and keep all network calls disabled/absent.
- **Future FS Access:** abstract file IO calls so File-System-Access API can be enabled via feature flag later.

## Roles & Responsibilities (suggested)
- **Tech Lead:** Owns state architecture, validation, and export pipelines; coordinates reviews.
- **Frontend Engineer:** Focus on UI shell, PhotoSorter, and keyboard interactions.
- **Content/QA:** Curates sample data, performs acceptance testing, maintains documentation.

## Milestones & Checkpoints
- **M1 (End Week 1):** Data ingestion working, basic shell renders with sample data.
- **M2 (End Week 3):** PhotoSorter and AutoBericht tabs feature-complete with validation.
- **M3 (End Week 4):** All export formats functional, blocking validations enforced.
- **M4 (End Week 5):** Acceptance tests passed, documentation complete, deliverable packaged.

Adhering to this plan will produce the offline AutoBericht MVP that mirrors the Excel/VBA workflow while remaining future-proof for deeper integration and automation.
