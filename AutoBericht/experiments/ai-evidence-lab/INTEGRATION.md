# AI Evidence Lab - Integration Boundary (Future Hook)

## Purpose

Define a stable integration contract between the standalone evidence lab and the main AutoBericht app.

## Boundary

Evidence Lab responsibilities:
- Read project folder inputs and sidecar in read-only mode.
- Build evidence artifacts under `outputs/evidence_lab/`.
- Produce patch preview JSON only.

Main AutoBericht responsibilities (future):
- Import selected evidence rows.
- Offer row-by-row accept/reject controls.
- Apply accepted changes into live sidecar with backup.

## Artifact Import Path

Expected input files for main app import:
- `outputs/evidence_lab/evidence_matches_selected_*.json`
- `outputs/evidence_lab/evidence_matches_included_*.json`
- `outputs/evidence_lab/sidecar_patch.preview_*.json`

Suggested import flow:
1. User picks artifact file from `outputs/evidence_lab/`.
2. Main app validates schema + sidecar row existence.
3. Main app shows preview diff per row.
4. Main app writes accepted edits only.

## Sidecar Compatibility Note

Current compatibility assumption:
- Sidecar rows are addressable by `row.id` string.
- Workstate fields used by downstream logic include:
  - `include`
  - `done`
  - `priority`
  - `findingText` (optional)
  - `recommendationText` (optional)

If row id semantics change, artifact row references must be mapped before apply.

## Migration Notes

If schema changes in Evidence Lab:
- Increment `spec_version`.
- Keep old parser for one previous minor version.
- Add migration table in this document.

Planned migration table:
- `0.1 -> 0.2`: TBD

## Experimental Guard

UI/runtime should keep an explicit experimental marker until integrated:
- Label: `Experimental - standalone`.
- No auto-apply to sidecar.
- Backup required before any future apply mode.
