# QA Corpus (AI Evidence Lab)

This folder contains a deterministic offline verification pack for v0.1.

## Contents

- `project-fixture/`:
  - `project_sidecar.json` (10 included rows)
  - `inputs/` with one DOCX, one XLSX, one text-PDF, one scanned-PDF
- `ground_truth.json`: curated extracted text blocks used by the local QA harness
- `expected_pairs.json`: 10 known row-to-evidence expectations
- `results/`: generated QA outputs and verification reports

## Run QA Harness

From repo root:

```bash
node AutoBericht/experiments/ai-evidence-lab/scripts/run_qa_e2e.mjs
node AutoBericht/experiments/ai-evidence-lab/scripts/verify_offline_mode.mjs
```

## What It Verifies

- End-to-end artifact generation (index + matches + patch preview + manifest)
- 10 known row/evidence checks
- weak vs strong match notes
- reproducibility across reruns
- citation traceability in draft/patch outputs
- offline/local-mode guardrails (no remote URLs in experiment runtime)
