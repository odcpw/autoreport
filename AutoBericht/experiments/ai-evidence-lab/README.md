# AI Evidence Lab (Experiment)

Standalone offline-first experiment area for document evidence retrieval and sidecar-safe patch proposals.

## Files

- Main spec: `SPEC.md`
- Dependency map: `DEPENDENCIES.md`
- Atomic tasks: `TASKS.md`
- Integration boundary: `INTEGRATION.md`
- App shell: `index.html`, `app.js`, `style.css`
- Workers: `workers/*.worker.js`
- Schemas: `schemas/*.schema.json`

## Run

1. Start local server from repo root (same server used for AutoBericht).
2. Open:
   - `http://127.0.0.1:<port>/AutoBericht/experiments/ai-evidence-lab/index.html`
3. Click `Pick Folder` and choose a project root containing:
   - `inputs/`
   - `project_sidecar.json` (optional but recommended)
4. Run:
   - `Scan Inputs`
   - `Build / Refresh Index`
   - `Find Evidence (Selected Row)` or `Find Evidence (Included Rows)`
   - `Generate Draft (Cited)`
   - `Export Proposal JSON`
   - optional: `Run Smoke E2E`

Outputs are written to:
- `outputs/evidence_lab/`

## Local Asset Placement

Keep all browser dependencies local (no CDN in this experiment):
- `lib/pdfjs/*`
- `lib/mammoth/*`
- `lib/sheetjs/*`
- `lib/tesseract/*`
- local model/runtime assets under project AI folder paths

## Troubleshooting

- If folder picking fails, use Chrome/Edge over `http://127.0.0.1`.
- If storage is evicted, click `Request Persistent Storage`.
- If WebGPU is unavailable, runtime falls back to WASM mode.
- If no citations appear, verify `Scan Inputs` and `Build / Refresh Index` were run.

## Current Limitations

- PDF/DOCX/OCR extraction is still fallback-level until local pdf.js, mammoth and tesseract assets are added.
- Current retrieval vectors are hashed local vectors, not multilingual model embeddings yet.
- Patch output is preview-only; no sidecar auto-apply in this experiment.

## Artifact Examples

- `evidence_index_2026-02-20T16-00-00-000Z.json`
- `evidence_matches_selected_2026-02-20T16-00-00-000Z.json`
- `evidence_matches_included_2026-02-20T16-00-00-000Z.json`
- `sidecar_patch.preview_2026-02-20T16-00-00-000Z.json`
- `export_manifest_2026-02-20T16-00-00-000Z.json`

## QA Verification (v0.1)

Use the local deterministic QA harness:

```bash
node AutoBericht/experiments/ai-evidence-lab/scripts/run_qa_e2e.mjs
node AutoBericht/experiments/ai-evidence-lab/scripts/verify_offline_mode.mjs
```

What this validates:
- 10 known row/evidence pairs (`qa-corpus/expected_pairs.json`)
- reproducibility across reruns
- citation traceability from matches -> draft -> patch preview
- offline/local guardrails (no remote URLs in runtime files)
- end-to-end artifact generation (selected + included matching paths)

## Roadmap

- v0.2: Real PDF/DOCX/XLSX extraction and OCR pipeline in workers.
- v0.3: Multilingual embeddings and retrieval quality pass.
- v0.4: Citation-grounded draft generation layer.
- v1.0 candidate: Main AutoBericht import bridge with review-first apply flow.

This experiment is intentionally isolated from the main AutoBericht runtime for now.
