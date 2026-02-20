# AI Evidence Lab (Standalone Experiment) - Spec v0.1

Date: 2026-02-20  
Status: Draft for implementation  
Scope: Standalone experiment in `AutoBericht/experiments/`, no dependency on main app runtime.

## 1) Purpose

Build an offline-first browser experiment that ingests customer documents (PDF, DOCX, XLSX, images), extracts searchable text + evidence citations, and proposes report-relevant evidence for AutoBericht questions.

Core user value:
- "Show me proof for this question."
- Example: "Is safety part of job descriptions?"
- Output: short evidence card with file + page citation and snippet.

## 2) Product Principles

- Offline by default.
- Citation-first (no uncited claims).
- Slow is acceptable; correctness and traceability first.
- Standalone experiment now, clean adapter path into AutoBericht later.

## 3) Non-Goals (v0.1)

- No automatic full sidecar overwrite.
- No hidden autonomous edits to report text.
- No cloud APIs.
- No dependency on VBA/Word pipeline.

## 4) Deployment Mode

- Separate web page(s) under `AutoBericht/experiments/`.
- User picks project folder via File System Access.
- Reads from:
  - `inputs/` documents
  - `project_sidecar.json` (read-only in first pass)
- Writes to:
  - `outputs/evidence_lab/` JSON artifacts
  - optional patch proposal JSON for controlled apply

## 5) High-Level Workflow

1. Pick project directory.
2. Scan documents in `inputs/` (or user-selected subset).
3. For each file:
   - Extract text directly where possible.
   - OCR pages/images when text is missing or weak.
4. Chunk and index with multilingual embeddings.
5. Generate evidence matches for:
   - row-level questions (`1.1.1`, `2.1.3`, etc.)
   - chapter-level prompts (optional later).
6. Display ranked evidence cards with citations.
7. Export:
   - `evidence_index.json` (index metadata)
   - `evidence_matches.json` (question -> citations)
   - optional `sidecar_patch.preview.json` (not auto-applied in v0.1).

## 6) Architecture (Standalone)

### 6.1 Frontend Components

- `index.html`: main shell (folder picker, run buttons, review panel)
- `ingest.worker.js`: document parse + OCR pipeline
- `embed.worker.js`: embeddings + retrieval index build
- `match.worker.js`: question matching + optional summarization
- `storage.js`: IndexedDB/OPFS persistence layer
- `adapters/sidecar.js`: read-only sidecar projection + optional patch export

### 6.2 Data Flow

- Files -> Parsed pages -> Normalized blocks -> Chunks -> Embeddings -> Search index
- Question prompt -> retrieval -> rerank -> evidence set -> optional generated summary

### 6.3 Storage

- IndexedDB for chunk metadata + vectors (v0.1)
- OPFS for large model/data cache (v0.2 candidate)
- Persistent storage request (`navigator.storage.persist`) on first run

## 7) Model Strategy

Use task-specialized models; avoid using one model for everything.

### 7.1 Text Embeddings (multilingual retrieval)

Primary:
- `Xenova/paraphrase-multilingual-MiniLM-L12-v2`
  - web-ready ONNX repo
  - strong practical size/speed balance
  - multilingual sentence embeddings

Optional heavy candidate:
- `BAAI/bge-m3`
  - stronger long-context/multilingual retrieval
  - much heavier; likely slower for browser-first workflow

### 7.2 OCR / Document Vision

Primary OCR engine:
- `tesseract.js` (WASM)
  - supports browser and many languages
  - note: no native PDF ingestion; rasterize PDF pages first

Secondary OCR path (optional experiment):
- Transformers.js `image-to-text` OCR-capable models (e.g. `onnx-community/mgp-str-base`)
  - useful for targeted scenarios
  - not default until validated on DE/FR/IT business docs

VLM usage:
- Keep Liquid VLM for "understand page/image context" and extraction cleanup, not as first OCR layer.

### 7.3 LLM Reasoning / Summarization

Primary:
- Existing local Liquid ONNX pipeline (already proven in repo)

CPU fallback option:
- `wllama` (WASM llama.cpp binding)
  - no backend/GPU required
  - use only if WebGPU path unavailable or unstable

## 8) Vision/OCR Feasibility Decision

Recommended hybrid:

- Step A: Try direct text extraction first (PDF text layer, DOCX, XLSX).
- Step B: OCR only where extraction quality is low or missing.
- Step C: Use VLM/LLM to structure findings and recommendations from cited evidence.

Why this path:
- Cheaper than VLM-on-every-page.
- Better traceability.
- Better multilingual robustness for business docs where many PDFs already have text layers.

## 9) Retrieval and Matching Design

### 9.1 Chunk Schema

Each chunk stores:
- `chunk_id`
- `file_path`
- `page_number` (if applicable)
- `source_type` (`pdf_text`, `pdf_ocr`, `docx`, `xlsx`, `image_ocr`)
- `language_guess`
- `text`
- `embedding`
- `char_span` or block offsets

### 9.2 Query Templates

Per row/question ID:
- Query from question text + chapter context + optional synonyms by locale.
- Example:
  - Base: "Zustaendigkeit und Verantwortung schriftlich geregelt"
  - Expansion: "Stellenbeschreibung, Verantwortlichkeit, Sicherheit, Gesundheit"

### 9.3 Ranking

- Top-k vector retrieval.
- Lightweight rerank score:
  - cosine score
  - term overlap with key tokens
  - citation diversity bonus (prefer distinct files/pages).

### 9.4 Output Object

For each row id:
- `row_id`
- `status` (`evidence_found`, `weak`, `none`)
- `citations[]`:
  - `file`
  - `page`
  - `snippet`
  - `score`
- optional:
  - `proposed_finding_note`
  - `proposed_recommendation_note`

## 10) Sidecar Integration Contract (Future-Safe)

Experiment remains standalone, but defines a strict adapter contract.

### 10.1 Read Contract

- Read current `project_sidecar.json`.
- Extract row IDs and existing workstate fields.
- Never mutate in-place during first phase.

### 10.2 Write Contract

- Write proposal file only:
  - `outputs/evidence_lab/sidecar_patch.preview.json`
- Patch format:
  - list of operations per `row_id`
  - `append` vs `replace`
  - strict provenance: citation IDs attached to each proposed text block.

### 10.3 Apply Strategy (later)

- Manual review screen with per-row accept/reject.
- Backup before apply:
  - `backup/project_sidecar.before_evidence_<timestamp>.json`

## 11) UI Spec (Experiment)

Panels:
- Left: corpus + ingestion status
- Center: selected question row + evidence cards
- Right: draft proposal (`finding`, `recommendation`) with citation chips

Main actions:
- `Pick Folder`
- `Scan Inputs`
- `Build/Refresh Index`
- `Find Evidence for Selected Row`
- `Find Evidence for Included Rows`
- `Export Proposal JSON`
- `Save Debug Log`

## 12) Runtime and Performance Policy

- Default device preference: WebGPU if available, else WASM.
- All heavy work in Web Workers.
- Batch/overnight mode allowed:
  - "Run all ingestion + indexing + row matching"
  - resumable progress checkpoints.

## 13) Safety and Auditability

- Every generated statement must be traceable to one or more citations.
- Add `uncited` warning state for any generated text.
- Keep raw extracted text for audit diffing.
- Log pipeline decisions:
  - extraction mode chosen (text vs OCR)
  - model used
  - elapsed time.

## 14) Risks and Mitigations

- OCR quality variability:
  - Mitigation: image preprocessing + confidence thresholds + manual review.
- Model size / browser memory limits:
  - Mitigation: chunked processing, external ONNX data handling, cached artifacts.
- WebGPU incompatibility across machines:
  - Mitigation: tested fallback to WASM/CPU path.
- Retrieval misses:
  - Mitigation: query expansion + hybrid lexical rerank + per-row manual retry.

## 15) Milestones

### M0 - Foundation (Standalone Skeleton)
- Page shell + folder picker + artifact writer
- Sidecar read adapter (read-only)
- Logging and persistent storage request

### M1 - Ingestion + Index
- PDF text extraction
- DOCX/XLSX extraction
- OCR fallback for scanned pages/images
- Embedding index build + save/load

### M2 - Evidence Matching
- Row-level retrieval
- Evidence card UI with citation previews
- Proposal JSON export

### M3 - Draft Generation
- Liquid-assisted finding/recommendation draft from citations
- citation-linked generation blocks
- stricter "no citation -> no text" mode

### M4 - Integration Bridge
- Patch preview format stable
- Optional apply pipeline (with backup + conflict checks)

## 16) Acceptance Criteria (v0.1)

- Can ingest mixed folder of PDF/DOCX/XLSX/images fully offline.
- Can produce evidence for at least selected rows with file/page snippets.
- Can export standalone proposal JSON without touching sidecar.
- Can resume from cached index without full re-ingestion.
- All evidence items are traceable and reproducible from artifacts.

## 17) Recommended Initial File Layout

`AutoBericht/experiments/ai-evidence-lab/`
- `README.md` (how to run)
- `SPEC.md` (this spec copied/adapted)
- `index.html`
- `style.css`
- `app.js`
- `workers/ingest.worker.js`
- `workers/embed.worker.js`
- `workers/match.worker.js`
- `lib/` (local browser libs)
- `schemas/evidence_matches.schema.json`
- `schemas/sidecar_patch.schema.json`

## 18) Sources Reviewed (2026-02-20)

- Transformers.js custom local/offline settings (`env.localModelPath`, `allowRemoteModels`, local wasm paths):  
  https://huggingface.co/docs/transformers.js/en/custom_usage
- Transformers.js WebGPU usage (`device: 'webgpu'`):  
  https://huggingface.co/docs/transformers.js/en/guides/webgpu
- ONNX Runtime Web support matrix and import modes (`wasm` + `webgpu`):  
  https://onnxruntime.ai/docs/get-started/with-javascript/web.html
- ONNX Runtime WebGPU guidance (lightweight model -> WASM; compute-intensive -> WebGPU):  
  https://onnxruntime.ai/docs/tutorials/web/ep-webgpu.html
- ONNX Runtime large model constraints (ArrayBuffer, protobuf 2GB, wasm 4GB, external data):  
  https://onnxruntime.ai/docs/tutorials/web/large-models.html
- Tesseract.js browser OCR scope (100+ languages, no direct PDF support):  
  https://github.com/naptha/tesseract.js
- PDF.js APIs (`getDocument`, `getTextContent`, `render`, `streamTextContent`):  
  https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib.html  
  https://mozilla.github.io/pdf.js/api/draft/module-pdfjsLib-PDFPageProxy.html
- Mammoth browser raw text extraction from DOCX `arrayBuffer`:  
  https://github.com/mwilliamson/mammoth.js
- SheetJS browser parsing (`XLSX.read`, `readFile` not supported in browser):  
  https://docs.sheetjs.com/docs/api/parse-options
- WebLLM (WebGPU-focused in-browser LLM engine):  
  https://webllm.mlc.ai/docs/user/get_started.html
- wllama (WASM CPU browser inference, worker-based, no WebGPU currently):  
  https://github.com/ngxson/wllama
- File System Access API (`showDirectoryPicker`, handles, worker availability):  
  https://developer.mozilla.org/en-US/docs/Web/API/File_System_API  
  https://developer.mozilla.org/en-US/docs/Web/API/Window/showDirectoryPicker
- Web Workers (background processing):  
  https://developer.mozilla.org/en-US/docs/Web/API/Web_Workers_API/Using_web_workers
- Service Workers (offline caching architecture):  
  https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API/Using_Service_Workers
- Storage persistence and eviction behavior (`navigator.storage.persist`):  
  https://developer.mozilla.org/en-US/docs/Web/API/Storage_API/Storage_quotas_and_eviction_criteria  
  https://developer.mozilla.org/en-US/docs/Web/API/StorageManager/persist
- Liquid model notes / ONNX variant guidance:  
  https://huggingface.co/LiquidAI/LFM2.5-VL-1.6B-ONNX  
  https://huggingface.co/LiquidAI/LFM2.5-VL-1.6B
