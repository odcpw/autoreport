# AI Voice Report (Standalone Experiment) - Spec v0.1

Date: 2026-02-20  
Status: Draft for implementation and parallel execution  
Scope: Standalone experiment in `AutoBericht/experiments/`, no dependency on main app runtime.

## 1) Purpose

Build an offline-first browser experiment that turns spoken consultant notes into row-scoped report drafts for `project_sidecar.json`.

Core user value:
- "I speak once, get structured draft text per row."
- Example spoken flow: "für 1.1.1 ... für 1.1.2 ..."
- Output: reviewable card drafts with per-field apply mode (`skip`, `append`, `replace`).

## 2) Product Principles

- Offline by default.
- Human-in-control edits only (no hidden writes).
- Row-accurate parsing first, polish second.
- Fast iteration loop: capture -> draft -> review -> apply.
- Standalone now, clean adapter path into main AutoBericht later.

## 3) Non-Goals (v0.1)

- No cloud APIs.
- No always-on listening / wake word.
- No autonomous sidecar mutation without explicit apply click.
- No direct Word/VBA integration.
- No multi-user collaboration features.

## 4) Deployment Mode

- Separate page under `AutoBericht/experiments/ai-voice-report/`.
- User selects project folder via File System Access.
- Reads from:
  - `project_sidecar.json`
- Writes to:
  - `voice_draft_<timestamp>.json`
  - `backup/project_sidecar_<timestamp>.json` (before apply)
  - updated `project_sidecar.json` only on explicit apply.

## 5) High-Level Workflow

1. Load local AI runtime.
2. Pick project folder.
3. Load and index `project_sidecar.json` rows.
4. Capture voice by:
   - global recording mode (multi-row narration), or
   - per-card mic mode (single row).
5. Transcribe audio locally (Whisper).
6. Parse ID mentions and segment notes by row.
7. Run local extraction model to produce:
   - `FINDING`
   - `RECOMMENDATION`
8. Review card drafts and choose action per field (`skip`, `append`, `replace`).
9. Save draft JSON and/or apply selected edits to sidecar (with backup first).

## 6) Architecture (Standalone)

### 6.1 Frontend Components

- `index.html`: app shell and controls
- `style.css`: layout and card UX
- `app.js`: runtime orchestration (ASR, parsing, extraction, apply)

### 6.2 Runtime State Domains

- AI state:
  - transformers runtime
  - tokenizer/session caches
  - ORT provider selection
- File-system state:
  - directory handle
  - loaded sidecar + row index
- Recording state:
  - global vs card mode
  - media recorder lifecycle
  - elapsed timer
- Run state:
  - parsed mentions/segments
  - extraction outputs
- UI card state:
  - current sidecar text
  - proposed text
  - apply action
  - take history

### 6.3 Data Flow

- Audio Blob -> decode -> Whisper transcript
- Transcript -> mention parser -> row segments
- Row segments -> Liquid extract prompt -> parsed draft blocks
- Draft blocks -> card review -> optional sidecar merge

## 7) Model Strategy

Use specialized local models instead of one general model.

### 7.1 ASR (Speech -> Text)

Primary:
- `Xenova/whisper-tiny`
- `Xenova/whisper-base`

Policy:
- Prefer WebGPU + fp16 where available.
- Fallback to fp32/WASM when `shader-f16` is unavailable.
- Use chunked transcription (`chunk_length_s`, `stride_length_s`) for long recordings.

### 7.2 Extraction (Notes -> Structured Fields)

Primary:
- `LiquidAI/LFM2.5-VL-1.6B-ONNX` local ONNX path

Prompt contract:
- strict block output per row:
  - `ID: <row_id>`
  - `FINDING: ...`
  - `RECOMMENDATION: ...`

Parser contract:
- ignore unknown IDs
- allow missing fields
- keep raw model output for audit in card take history.

## 8) Mention Parsing and Segmentation Design

### 8.1 Supported Mention Patterns

- Dotted numeric IDs (e.g., `1.1.3`, `2.1.2b`)
- Spaced numeric IDs (e.g., `1 1 3`)
- Locale-aware spoken-number variants (initial focus: German)

### 8.2 Resolution Rules

- Prefer direct row IDs.
- Map subquestion IDs to parent row IDs when needed.
- Resolve overlaps deterministically (earliest, longest, row-first).

### 8.3 Segment Output

For each resolved mention:
- capture text between this mention and next mention
- trim separator noise
- group by `row_id`
- build sorted segment preview in UI.

## 9) Draft Card Contract

Per row card stores:
- `rowId`
- `title`
- current sidecar values:
  - `currentFinding`
  - `currentRecommendation`
- proposed values:
  - `proposedFinding`
  - `proposedRecommendation`
- apply actions:
  - `actionFinding`
  - `actionRecommendation`
- `takes[]` history with timestamp, source note, parsed output, raw output.

## 10) Sidecar Integration Contract

### 10.1 Read Contract

- Read `project_sidecar.json` and extract report project.
- Build stable `row_id -> row` map.
- Never mutate on load.

### 10.2 Apply Contract (v0.1)

- Re-read sidecar before apply.
- Write backup first:
  - `backup/project_sidecar_<timestamp>.json`
- Apply only selected non-empty fields.
- Merge policy:
  - `skip` -> no change
  - `append` -> existing + `\n\n` + new
  - `replace` -> new value
- Update `meta.updatedAt` on successful write.

### 10.3 Draft Export Contract

- Save standalone draft file:
  - `voice_draft_<timestamp>.json`
- Include:
  - transcript raw
  - parsed segmentation
  - card proposals and actions
  - model metadata (ASR + extraction model IDs)

## 11) UI Spec (Experiment)

Panels:
- Header + runtime controls
- Recording controls (global)
- Transcript and segmented output
- Draft cards (review/apply)
- Log panel

Main actions:
- `Load AI`
- `Pick Project Folder`
- `Load project_sidecar.json`
- `Start/Stop recording`
- `Parse IDs`
- `Extract Finding + Recommendation`
- `Save draft JSON`
- `Apply Selected to Sidecar`

Card-level actions:
- `Mic this card` (single-row fast path)
- per-field action radios (`skip`, `append`, `replace`)

## 12) Runtime and Performance Policy

- Default device preference: WebGPU if available, else WASM.
- Keep heavy compute off the UI thread where feasible in future refactor.
- Use micro-batch extraction (target batch size 4-8 rows).
- Show explicit status transitions for long operations.
- Cache model sessions/tokenizers during one page lifecycle.

## 13) Safety, Privacy, and Auditability

- Local-only model loading by default.
- No implicit network calls in offline mode.
- Explicit apply button required for sidecar writes.
- Backup always created before sidecar mutation.
- Keep operation log with timestamps.
- Preserve raw model output in take history for review.

## 14) Risks and Mitigations

- ASR mishears row IDs:
  - Mitigation: deterministic ID parser + warnings + card-level mic fallback.
- Model output misses IDs:
  - Mitigation: strict output parser, missing-ID warnings, manual card edits.
- Browser/runtime mismatch for WebGPU:
  - Mitigation: ORT provider fallback to WASM.
- Large-project latency:
  - Mitigation: batch extraction, staged status, save/reload draft artifacts.

## 15) Milestones

### M0 - Foundation (Current)
- Standalone UI shell
- Local ASR + extraction model wiring
- Sidecar load/apply loop with backup

### M1 - Robust Parsing
- Better locale adapters (de/fr/it)
- mention-confidence UI
- unresolved-span assistant hints

### M2 - Throughput + UX
- extraction worker split
- faster per-card loop
- draft diff preview vs current sidecar text

### M3 - Integration Bridge
- optional patch preview JSON (no direct apply)
- adapter contract to main AutoBericht flow

## 16) Acceptance Criteria (v0.1)

- Can load sidecar and index rows from local project folder.
- Can transcribe microphone audio offline.
- Can parse spoken row IDs into row-scoped segments.
- Can generate finding/recommendation drafts for parsed rows.
- Can save draft JSON artifact.
- Can apply selected edits with backup-first behavior.

## 17) Recommended File Layout

`AutoBericht/experiments/ai-voice-report/`
- `README.md`
- `SPEC.md`
- `TASKS.md`
- `index.html`
- `style.css`
- `app.js`

## 18) Open Questions for Product Alignment

- Should v0.2 default to patch-preview export only (no direct sidecar write)?
- Should per-card voice become the primary UX and global transcript become optional advanced mode?
- Which locale should receive first-class spoken-number parsing after German (`fr` or `it`)?
