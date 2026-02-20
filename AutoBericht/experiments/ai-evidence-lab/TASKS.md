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

- [ ] `T033` Finalize `evidence_index.json` shape.
- [x] `T034` Finalize `evidence_matches.json` shape.
- [x] `T035` Finalize `sidecar_patch.preview.json` shape.
- [x] `T036` Implement JSON schema validation helper.
- [ ] `T037` Validate all exports before writing.
- [ ] `T038` Include `created_at`, `spec_version`, `generator` in all artifacts.

## 5. Ingestion Pipeline Orchestration

- [x] `T039` Implement "Scan Inputs" command wiring in UI.
- [x] `T040` Add worker message protocol (`start`, `progress`, `done`, `error`).
- [x] `T041` Add cancellable ingestion run token.
- [ ] `T042` Add resumable checkpoint file (`ingest_checkpoint.json`).
- [ ] `T043` Skip unchanged files based on hash + modified time.
- [x] `T044` Record per-file ingest status (`pending/running/done/failed`).
- [x] `T045` Render ingest progress table live.

## 6. PDF Text Extraction

- [ ] `T046` Integrate local `pdf.js` loading path.
- [ ] `T047` Extract page text via `getTextContent`.
- [ ] `T048` Normalize text blocks from page items.
- [ ] `T049` Compute text-density score per page.
- [ ] `T050` Mark pages below threshold as OCR candidates.
- [ ] `T051` Capture page count and extraction time.
- [ ] `T052` Write raw extracted page text cache.

## 7. DOCX and XLSX Extraction

- [ ] `T053` Integrate `mammoth` for DOCX raw text extraction.
- [ ] `T054` Normalize DOCX text into paragraph blocks.
- [ ] `T055` Integrate `SheetJS` for XLSX parsing.
- [ ] `T056` Convert worksheet cells to row text blocks.
- [ ] `T057` Add worksheet metadata (`sheet_name`, row index).
- [ ] `T058` Write raw extraction cache for DOCX/XLSX.

## 8. Image OCR and PDF OCR Fallback

- [ ] `T059` Integrate local `tesseract.js` worker.
- [ ] `T060` Add language pack mapping for `de`, `fr`, `it`, `en`.
- [ ] `T061` Implement image file OCR path (`png/jpg/jpeg/webp`).
- [ ] `T062` Implement PDF page rasterization for OCR fallback.
- [ ] `T063` OCR only pages marked low-text-density.
- [ ] `T064` Capture OCR confidence score per block.
- [ ] `T065` Flag low-confidence OCR snippets for UI warning.
- [ ] `T066` Store OCR output with `source_type` metadata.

## 9. Text Normalization + Chunking

- [ ] `T067` Normalize whitespace and line breaks.
- [ ] `T068` Preserve sentence punctuation where possible.
- [ ] `T069` Detect language guess per block.
- [ ] `T070` Implement chunker by max char/token budget.
- [ ] `T071` Add overlap window between adjacent chunks.
- [ ] `T072` Assign stable `chunk_id` values.
- [ ] `T073` Attach citation metadata (`file`, `page`, `source_type`).
- [ ] `T074` Persist chunk list cache to IndexedDB.

## 10. Embeddings + Index Build

- [ ] `T075` Integrate multilingual embedding model loader.
- [ ] `T076` Force local model mode (no remote model fetch).
- [ ] `T077` Compute embeddings in batch mode.
- [ ] `T078` Store vectors + metadata in IndexedDB.
- [ ] `T079` Build searchable in-memory index on load.
- [ ] `T080` Add incremental index update for changed chunks only.
- [ ] `T081` Add index build timing metrics.
- [ ] `T082` Export `evidence_index.json` snapshot.

## 11. Retrieval + Ranking

- [ ] `T083` Build query from selected row question text.
- [ ] `T084` Add chapter context terms to query expansion.
- [ ] `T085` Add locale-aware synonym expansion hook.
- [ ] `T086` Run top-k vector retrieval.
- [ ] `T087` Add lexical overlap rerank component.
- [ ] `T088` Add citation diversity bonus in final score.
- [ ] `T089` Cap duplicate chunks from same page.
- [ ] `T090` Output ranked citation list per row.

## 12. Evidence UI

- [ ] `T091` Render evidence cards with file/page/snippet/score.
- [ ] `T092` Add click-to-expand full snippet context.
- [ ] `T093` Highlight query term matches in snippets.
- [ ] `T094` Show evidence state badge (`found/weak/none`).
- [ ] `T095` Add filter toggle for low-confidence OCR citations.
- [ ] `T096` Add file-level grouping toggle.
- [ ] `T097` Add copy citation text action.

## 13. Generation Layer (Optional in v0.1, Enabled in v0.2)

- [ ] `T098` Add "Generate Draft from Citations" button.
- [ ] `T099` Build strict citation-grounded prompt template.
- [ ] `T100` Feed only selected citations to generator.
- [ ] `T101` Parse generator output into `finding` and `recommendation`.
- [ ] `T102` Reject generation output with zero citation links.
- [ ] `T103` Display uncited warning if output cannot be grounded.
- [ ] `T104` Store generated text with citation IDs.

## 14. Export Artifacts

- [ ] `T105` Export `evidence_matches.json` for selected row.
- [ ] `T106` Export `evidence_matches.json` for all included rows.
- [ ] `T107` Export `sidecar_patch.preview.json` (proposal only).
- [ ] `T108` Export timestamped run log text file.
- [ ] `T109` Save all exports into `outputs/evidence_lab/`.
- [ ] `T110` Show export success summary with file paths.

## 15. Patch Preview Contract (No Auto-Apply v0.1)

- [ ] `T111` Implement per-row patch op builder (`append`/`replace`/`skip`).
- [ ] `T112` Include source citation IDs in each patch op.
- [ ] `T113` Include original row hash guard in patch op.
- [ ] `T114` Add dry-run validation against current sidecar row map.
- [ ] `T115` Render patch preview table in UI.
- [ ] `T116` Block export if patch references unknown row IDs.

## 16. Logging, Debug, and Audit

- [x] `T117` Implement structured log collector with levels.
- [ ] `T118` Log model/runtime selection decisions.
- [ ] `T119` Log extraction mode per file/page (`text` vs `ocr`).
- [ ] `T120` Log processing durations per stage.
- [ ] `T121` Log warnings for OCR confidence and missing citations.
- [x] `T122` Add "Save Debug Log" button.
- [ ] `T123` Add run summary panel (counts, duration, failures).

## 17. Reliability and Resume

- [ ] `T124` Resume ingestion from checkpoint after reload.
- [ ] `T125` Resume embedding from last completed chunk batch.
- [ ] `T126` Resume matching from last completed row batch.
- [ ] `T127` Add "Clear cache and rebuild" action.
- [ ] `T128` Add stale-cache detection when models change.
- [ ] `T129` Add graceful cancellation handling in all workers.

## 18. Performance Controls

- [ ] `T130` Add configurable batch size for OCR and embeddings.
- [ ] `T131` Add configurable max pages per run (for smoke tests).
- [ ] `T132` Add "overnight mode" profile preset.
- [ ] `T133` Add CPU-friendly profile preset (smaller batch).
- [ ] `T134` Add WebGPU-preferred profile preset.
- [ ] `T135` Add memory guard when index size crosses threshold.

## 19. QA Corpus and Verification

- [ ] `T136` Prepare small mixed-language test corpus (de/fr/it/en).
- [ ] `T137` Include at least one text-PDF and one scanned-PDF sample.
- [ ] `T138` Include one DOCX and one XLSX sample.
- [ ] `T139` Verify retrieval for 10 known question-evidence pairs.
- [ ] `T140` Record precision notes for weak vs strong matches.
- [ ] `T141` Verify outputs are reproducible across reruns.
- [ ] `T142` Verify no network calls in offline-local mode.

## 20. Security and Privacy Guardrails

- [ ] `T143` Enforce local-only model flags in runtime setup.
- [ ] `T144` Show explicit warning if remote model flag is enabled.
- [ ] `T145` Avoid logging full sensitive document text by default.
- [ ] `T146` Redact snippets in debug log unless explicit opt-in.
- [ ] `T147` Add privacy note in UI about local processing.

## 21. Integration Readiness (Future Hook)

- [ ] `T148` Document adapter boundary to main AutoBericht app.
- [ ] `T149` Define import path from evidence outputs into sidecar flow.
- [ ] `T150` Add compatibility note for current sidecar schema version.
- [ ] `T151` Add migration note if schema changes later.
- [ ] `T152` Add "experimental feature" label and guard.

## 22. Documentation

- [ ] `T153` Update `README.md` with run instructions.
- [ ] `T154` Add troubleshooting section (WebGPU, OCR language packs, storage).
- [ ] `T155` Add model file placement guide for local assets.
- [ ] `T156` Add artifact file examples in docs.
- [ ] `T157` Add known limitations list for v0.1.
- [ ] `T158` Add roadmap notes for v0.2/v1 integration.

## 23. Definition of Done for v0.1

- [ ] `T159` End-to-end run: ingest -> index -> match selected row -> export artifacts.
- [ ] `T160` End-to-end run: match included rows batch -> export artifacts.
- [ ] `T161` Evidence cards show file/page/snippet/score for matched rows.
- [ ] `T162` `sidecar_patch.preview.json` validates against schema.
- [ ] `T163` No direct mutation of `project_sidecar.json` occurs in v0.1.
- [ ] `T164` Debug log + run summary saved successfully.
- [ ] `T165` Manual review confirms citation traceability for generated proposals.

