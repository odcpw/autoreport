# AGENTS.md — AutoBericht (autoreport)

## Purpose (Project Context)
AutoBericht is an offline, policy-safe reporting system for safety culture assessments. It streamlines ingestion of self-assessments and photos, editing of findings/recommendations, and export to Word/PPT/PDF. The current redesign is browser-first, with a minimal editor that saves state locally and a thin Word-template VBA export layer.

Key constraints from the spec:
- Locked-down corporate Windows/Edge environment.
- No installs, no unsafe browser flags, no online services.
- Everything works offline from a local project folder.
- File System Access API is preferred; if blocked, the app cannot operate (no fallback by design).

## Greenfield + Solo Development
- This is a greenfield project: treat the redesign as new build work.
- No legacy fallback or compatibility is required; legacy assets are reference-only.
- We are working alone until release is ready (single-author workflows; no multi-user features).

## Current Direction (Authoritative)
- Minimal editor: AutoBericht/mini/ (current UI direction).
- Canonical state: project_sidecar.json (editor writes/reads this).
- Docs: AutoBericht/docs/design-spec.md, AutoBericht/docs/system-overview.md, docs/STATUS.md.
- Legacy UI/docs: legacy/ (reference only, do not extend).

## Scope (What We Build)
Inputs:
- Self-assessment Excel (fixed question IDs).
- Photos and customer documents.

Outputs:
- Word report, PPT deck, PDF (templates remain fixed; exports use latest templates).

Non-goals:
- Online/cloud features.
- Real-time collaboration.
- Changing corporate IT/security or report templates.
- Backwards compatibility with legacy UI or data formats.

## Architecture Snapshot
Single project folder contains:
- project_sidecar.json (working state, canonical)
- project_db.xlsx (optional readable archive; future)
- self_assessment.xlsx
- photos/, Outputs/, cache/ (optional)

Data flow:
Self-assessment + Photos -> Minimal Editor -> project_sidecar.json -> (optional) Word-template VBA export -> Word/PPT/PDF.

## Domain Rules (High-Value)
- Question IDs are stable; some are "tiroir" sub-items consolidated into one report line.
- Field observations (chapter 4.8) are handled separately from standard findings.
- Report includes only improvement opportunities; positives are omitted.
- Score rule: 1=0%, 2=25%, 3=50%, 4=75%; excluded findings count as 100%.
- Reports are single-language (DE/FR/IT); templates are fixed but updated over time.

## Developer Notes
- Keep changes offline-first and no-install.
- VBA lives in the Word template (export layer), not Excel.
- Prefer edits in AutoBericht/mini/ and AutoBericht/docs/.
- Treat legacy/ as read-only reference.
- Avoid bundlers or new dependencies unless explicitly required.
