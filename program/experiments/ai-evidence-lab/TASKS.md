# AI Evidence Lab - Atomic Task List (v0.1)

Use this as the build checklist for `SPEC.md`.

## 0. Project Setup

- [x] `T001` Create `index.html` with app shell and root mount points.
- [x] `T002` Create `style.css` with base layout (3-column desktop, 1-column mobile).
- [x] `T003` Create `app.js` entry script with boot log.
- [x] `T004` Create `workers/ingest.worker.js`.
- [x] `T005` Create `workers/embed.worker.js`.
- [x] `T006` Create `workers/match.worker.js`.
- [x] `T007` Create `lib/` placeholder with README of expected local libs.
- [x] `T008` Create `schemas/evidence_matches.schema.json`.
- [x] `T009` Create `schemas/sidecar_patch.schema.json`.
- [x] `T010` Add app version banner (`Evidence Lab v0.1`) to UI.

## 1. Runtime + Capability Detection

- [x] `T011` Detect File System Access API support at startup.
- [x] `T012` Detect Web Worker support at startup.
- [x] `T013` Detect WebGPU availability (`navigator.gpu`) at startup.
- [x] `T014` Detect storage estimate (`navigator.storage.estimate`) at startup.
- [x] `T015` Add "request persistent storage" action.
- [x] `T016` Persist capability snapshot in memory for logs/artifacts.
- [x] `T017` Render capability status chips in UI.

## 2. Folder and File Access

- [x] `T018` Implement `Pick Folder` action via `showDirectoryPicker`.
- [x] `T019` Resolve `inputs/` directory handle from picked project folder.
- [x] `T020` Resolve `outputs/evidence_lab/` directory handle (create if missing).
- [x] `T021` Resolve `project_sidecar.json` handle read-only.
- [x] `T022` Show clear error if `inputs/` missing.
- [x] `T023` Show clear error if sidecar missing.
- [x] `T024` Add file inventory scan for supported types.
- [x] `T025` Show file inventory table with name, type, size.

## 3. Sidecar Adapter (Read-Only v0.1)

- [x] `T026` Parse `project_sidecar.json` into validated object.
- [x] `T027` Extract project metadata (`locale`, company fields) for context.
- [x] `T028` Extract all row IDs and row titles for questions.
- [x] `T029` Extract include/done/priority flags for display context.
- [x] `T030` Build stable internal row map (`row_id -> row_context`).
- [x] `T031` Render row picker list (search by id/title).
- [x] `T032` Add selected-row details panel.

## 4. Schema + Artifact Contracts

- [x] `T033` Finalize `evidence_index.json` shape.
- [x] `T034` Finalize `evidence_matches.json` shape.
- [x] `T035` Finalize `sidecar_patch.preview.json` shape.
- [x] `T036` Implement JSON schema validation helper.
- [x] `T037` Validate all exports before writing.
- [x] `T038` Include `created_at`, `spec_version`, `generator` in all artifacts.

## 5. Ingestion Pipeline Orchestration

- [x] `T039` Implement "Scan Inputs" command wiring in UI.
- [x] `T040` Add worker message protocol (`start`, `progress`, `done`, `error`).
- [x] `T041` Add cancellable ingestion run token.
- [x] `T042` Add resumable checkpoint file (`ingest_checkpoint.json`).
- [x] `T043` Skip unchanged files based on hash + modified time.
- [x] `T044` Record per-file ingest status (`pending/running/done/failed`).
- [x] `T045` Render ingest progress table live.

## 6. PDF Text Extraction

- [x] `T046` Integrate local `pdf.js` loading path.
- [x] `T047` Extract page text via `getTextContent`.
- [x] `T048` Normalize text blocks from page items.
- [x] `T049` Compute text-density score per page.
- [x] `T050` Mark pages below threshold as OCR candidates.
- [x] `T051` Capture page count and extraction time.
- [x] `T052` Write raw extracted page text cache.

## 7. DOCX and XLSX Extraction

- [x] `T053` Integrate `mammoth` for DOCX raw text extraction.
- [x] `T054` Normalize DOCX text into paragraph blocks.
- [x] `T055` Integrate `SheetJS` for XLSX parsing.
- [x] `T056` Convert worksheet cells to row text blocks.
- [x] `T057` Add worksheet metadata (`sheet_name`, row index).
- [x] `T058` Write raw extraction cache for DOCX/XLSX.

## 8. Image OCR and PDF OCR Fallback

- [x] `T059` Integrate local `tesseract.js` worker.
- [x] `T060` Add language pack mapping for `de`, `fr`, `it`, `en`.
- [x] `T061` Implement image file OCR path (`png/jpg/jpeg/webp`).
- [x] `T062` Implement PDF page rasterization for OCR fallback.
- [x] `T063` OCR only pages marked low-text-density.
- [x] `T064` Capture OCR confidence score per block.
- [x] `T065` Flag low-confidence OCR snippets for UI warning.
- [x] `T066` Store OCR output with `source_type` metadata.

## 9. Text Normalization + Chunking

- [x] `T067` Normalize whitespace and line breaks.
- [x] `T068` Preserve sentence punctuation where possible.
- [x] `T069` Detect language guess per block.
- [x] `T070` Implement chunker by max char/token budget.
- [x] `T071` Add overlap window between adjacent chunks.
- [x] `T072` Assign stable `chunk_id` values.
- [x] `T073` Attach citation metadata (`file`, `page`, `source_type`).
- [x] `T074` Persist chunk list cache to IndexedDB.

## 10. Embeddings + Index Build

- [x] `T075` Integrate multilingual embedding model loader.
- [x] `T076` Force local model mode (no remote model fetch).
- [x] `T077` Compute embeddings in batch mode.
- [x] `T078` Store vectors + metadata in IndexedDB.
- [x] `T079` Build searchable in-memory index on load.
- [x] `T080` Add incremental index update for changed chunks only.
- [x] `T081` Add index build timing metrics.
- [x] `T082` Export `evidence_index.json` snapshot.

## 11. Retrieval + Ranking

- [x] `T083` Build query from selected row question text.
- [x] `T084` Add chapter context terms to query expansion.
- [x] `T085` Add locale-aware synonym expansion hook.
- [x] `T086` Run top-k vector retrieval.
- [x] `T087` Add lexical overlap rerank component.
- [x] `T088` Add citation diversity bonus in final score.
- [x] `T089` Cap duplicate chunks from same page.
- [x] `T090` Output ranked citation list per row.

## 12. Evidence UI

- [x] `T091` Render evidence cards with file/page/snippet/score.
- [x] `T092` Add click-to-expand full snippet context.
- [x] `T093` Highlight query term matches in snippets.
- [x] `T094` Show evidence state badge (`found/weak/none`).
- [x] `T095` Add filter toggle for low-confidence OCR citations.
- [x] `T096` Add file-level grouping toggle.
- [x] `T097` Add copy citation text action.

## 13. Generation Layer (Optional in v0.1, Enabled in v0.2)

- [x] `T098` Add "Generate Draft from Citations" button.
- [x] `T099` Build strict citation-grounded prompt template.
- [x] `T100` Feed only selected citations to generator.
- [x] `T101` Parse generator output into `finding` and `recommendation`.
- [x] `T102` Reject generation output with zero citation links.
- [x] `T103` Display uncited warning if output cannot be grounded.
- [x] `T104` Store generated text with citation IDs.

## 14. Export Artifacts

- [x] `T105` Export `evidence_matches.json` for selected row.
- [x] `T106` Export `evidence_matches.json` for all included rows.
- [x] `T107` Export `sidecar_patch.preview.json` (proposal only).
- [x] `T108` Export timestamped run log text file.
- [x] `T109` Save all exports into `outputs/evidence_lab/`.
- [x] `T110` Show export success summary with file paths.

## 15. Patch Preview Contract (No Auto-Apply v0.1)

- [x] `T111` Implement per-row patch op builder (`append`/`replace`/`skip`).
- [x] `T112` Include source citation IDs in each patch op.
- [x] `T113` Include original row hash guard in patch op.
- [x] `T114` Add dry-run validation against current sidecar row map.
- [x] `T115` Render patch preview table in UI.
- [x] `T116` Block export if patch references unknown row IDs.

## 16. Logging, Debug, and Audit

- [x] `T117` Implement structured log collector with levels.
- [x] `T118` Log model/runtime selection decisions.
- [x] `T119` Log extraction mode per file/page (`text` vs `ocr`).
- [x] `T120` Log processing durations per stage.
- [x] `T121` Log warnings for OCR confidence and missing citations.
- [x] `T122` Add "Save Debug Log" button.
- [x] `T123` Add run summary panel (counts, duration, failures).

## 17. Reliability and Resume

- [x] `T124` Resume ingestion from checkpoint after reload.
- [x] `T125` Resume embedding from last completed chunk batch.
- [x] `T126` Resume matching from last completed row batch.
- [x] `T127` Add "Clear cache and rebuild" action.
- [x] `T128` Add stale-cache detection when models change.
- [x] `T129` Add graceful cancellation handling in all workers.

## 18. Performance Controls

- [x] `T130` Add configurable batch size for OCR and embeddings.
- [x] `T131` Add configurable max pages per run (for smoke tests).
- [x] `T132` Add "overnight mode" profile preset.
- [x] `T133` Add CPU-friendly profile preset (smaller batch).
- [x] `T134` Add WebGPU-preferred profile preset.
- [x] `T135` Add memory guard when index size crosses threshold.

## 19. QA Corpus and Verification

- [x] `T136` Prepare small mixed-language test corpus (de/fr/it/en).
- [x] `T137` Include at least one text-PDF and one scanned-PDF sample.
- [x] `T138` Include one DOCX and one XLSX sample.
- [x] `T139` Verify retrieval for 10 known question-evidence pairs.
- [x] `T140` Record precision notes for weak vs strong matches.
- [x] `T141` Verify outputs are reproducible across reruns.
- [x] `T142` Verify no network calls in offline-local mode.

## 20. Security and Privacy Guardrails

- [x] `T143` Enforce local-only model flags in runtime setup.
- [x] `T144` Show explicit warning if remote model flag is enabled.
- [x] `T145` Avoid logging full sensitive document text by default.
- [x] `T146` Redact snippets in debug log unless explicit opt-in.
- [x] `T147` Add privacy note in UI about local processing.

## 21. Integration Readiness (Future Hook)

- [x] `T148` Document adapter boundary to main AutoBericht app.
- [x] `T149` Define import path from evidence outputs into sidecar flow.
- [x] `T150` Add compatibility note for current sidecar schema version.
- [x] `T151` Add migration note if schema changes later.
- [x] `T152` Add "experimental feature" label and guard.

## 22. Documentation

- [x] `T153` Update `README.md` with run instructions.
- [x] `T154` Add troubleshooting section (WebGPU, OCR language packs, storage).
- [x] `T155` Add model file placement guide for local assets.
- [x] `T156` Add artifact file examples in docs.
- [x] `T157` Add known limitations list for v0.1.
- [x] `T158` Add roadmap notes for v0.2/v1 integration.

## 23. Definition of Done for v0.1

- [x] `T159` End-to-end run: ingest -> index -> match selected row -> export artifacts.
- [x] `T160` End-to-end run: match included rows batch -> export artifacts.
- [x] `T161` Evidence cards show file/page/snippet/score for matched rows.
- [x] `T162` `sidecar_patch.preview.json` validates against schema.
- [x] `T163` No direct mutation of `project_sidecar.json` occurs in v0.1.
- [x] `T164` Debug log + run summary saved successfully.
- [x] `T165` Manual review confirms citation traceability for generated proposals.
