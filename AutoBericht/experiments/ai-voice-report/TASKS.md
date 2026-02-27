# AI Voice Report - Atomic Task List (v0.1)

Use this as the build checklist for `SPEC.md`.

## 0. Project Setup

- [ ] `V001` Keep standalone shell at `index.html` + `style.css` + `app.js`.
- [ ] `V002` Add version banner (`Voice Report v0.1`) in UI.
- [ ] `V003` Add app-level config object (`model ids`, `batch size`, `locale defaults`).

## 1. Runtime + Capability Detection

- [ ] `V004` Detect File System Access support at startup.
- [ ] `V005` Detect microphone capture support (`getUserMedia`).
- [ ] `V006` Detect MediaRecorder mime support matrix.
- [ ] `V007` Detect WebGPU availability.
- [ ] `V008` Detect `shader-f16` capability for dtype policy.
- [ ] `V009` Render capability chips in UI.

## 2. Folder + Sidecar Access

- [ ] `V010` Implement `Pick Project Folder` via `showDirectoryPicker`.
- [ ] `V011` Implement sidecar load (`project_sidecar.json`).
- [ ] `V012` Validate sidecar shape and report clear errors.
- [ ] `V013` Build row map (`row_id -> row object`).
- [ ] `V014` Build sub-id map (`sub_id -> row_id`).
- [ ] `V015` Render sidecar load summary (locale, row count).

## 3. Recording Engine

- [ ] `V016` Implement global record start/stop flow.
- [ ] `V017` Implement card-level mic record start/stop flow.
- [ ] `V018` Enforce single active recording across modes.
- [ ] `V019` Add recording timer UI.
- [ ] `V020` Stop and release media tracks reliably.

## 4. ASR Pipeline

- [ ] `V021` Load transformers runtime from local bundle.
- [ ] `V022` Configure local-only model mode.
- [ ] `V023` Resolve ASR dtype (`fp16` preferred, `fp32` fallback).
- [ ] `V024` Validate required Whisper ONNX files before run.
- [ ] `V025` Decode audio to mono 16kHz float array.
- [ ] `V026` Run transcription with chunk + stride options.
- [ ] `V027` Write transcript to raw textarea.
- [ ] `V028` Log ASR stage timings.

## 5. Mention Parsing

- [ ] `V029` Parse dotted numeric IDs.
- [ ] `V030` Parse spaced numeric IDs.
- [ ] `V031` Parse German spoken-number IDs.
- [ ] `V032` Resolve overlaps deterministically.
- [ ] `V033` Map subquestion mentions to parent row IDs.
- [ ] `V034` Segment transcript by mention boundaries.
- [ ] `V035` Group notes by row and output sorted preview.
- [ ] `V036` Render parser warnings for empty or unresolved segments.

## 6. Extraction Pipeline

- [ ] `V037` Build strict extraction prompt template.
- [ ] `V038` Batch row notes into micro-batches (target 6).
- [ ] `V039` Run local Liquid extraction per batch.
- [ ] `V040` Parse response blocks by `ID/FINDING/RECOMMENDATION`.
- [ ] `V041` Record missing-ID parse warnings.
- [ ] `V042` Persist raw model output in take history.

## 7. Draft Cards UX

- [ ] `V043` Seed cards from sidecar rows.
- [ ] `V044` Show current vs proposed text per field.
- [ ] `V045` Add action radios (`skip`, `append`, `replace`) per field.
- [ ] `V046` Show per-card take list metadata.
- [ ] `V047` Keep apply button disabled when no actionable changes exist.
- [ ] `V048` Keep controls coherent during recording/transcribing states.

## 8. Draft Export

- [ ] `V049` Export `voice_draft_<timestamp>.json`.
- [ ] `V050` Include transcript, parsed segmentation, and card payload.
- [ ] `V051` Include runtime metadata (models, locale, timestamp).
- [ ] `V052` Confirm export write success in status/log.

## 9. Sidecar Apply

- [ ] `V053` Re-read sidecar immediately before apply.
- [ ] `V054` Backup current sidecar to `backup/` before mutation.
- [ ] `V055` Apply merge policy (`skip`, `append`, `replace`).
- [ ] `V056` Update `meta.updatedAt`.
- [ ] `V057` Save mutated sidecar and report applied field count.
- [ ] `V058` Handle missing rows gracefully during apply.

## 10. Logging + Audit

- [ ] `V059` Keep timestamped event log stream.
- [ ] `V060` Log backend/provider selection decisions.
- [ ] `V061` Log ASR and extraction duration metrics.
- [ ] `V062` Log parser warnings and missing IDs.
- [ ] `V063` Add optional download/save log action.

## 11. Reliability and Recovery

- [ ] `V064` Guard against double-click re-entry for long actions.
- [ ] `V065` Add cancellation path for long extraction batches.
- [ ] `V066` Preserve current draft state on recoverable errors.
- [ ] `V067` Add clear/reset action for transcript + parsed state.

## 12. Privacy + Security Guardrails

- [ ] `V068` Keep remote model loading disabled by default.
- [ ] `V069` Warn clearly if remote mode is ever enabled.
- [ ] `V070` Avoid writing raw audio blobs to disk by default.
- [ ] `V071` Avoid logging sensitive full sidecar payloads.

## 13. v0.2 Forward Hooks

- [ ] `V072` Move extraction to worker thread.
- [ ] `V073` Add partial/streaming transcript preview.
- [ ] `V074` Add row-level confidence scoring for mention resolution.
- [ ] `V075` Add patch-preview export mode (`voice_patch.preview.json`).
- [ ] `V076` Add locale packs for French and Italian spoken-number parsing.
