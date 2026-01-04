# AutoBericht: Current State vs Target

This document tracks the **redesign** direction (2026).

## Operating Constraints

- Offline-only (no CDN, no telemetry, no remote fonts/images).
- No unsafe browser flags.
- Locked-down corporate Edge environment.
- Prefer File System Access API for direct folder writes.

## Current State (Redesign)

- Minimal editor exists in `AutoBericht/mini/` with sidecar load/save.
- File System Access + SheetJS spike exists in `AutoBericht/experiments/`.
- Shared debug log exporter is available across UI pages.

## Target End-to-End Flow

1. User opens the minimal editor.
2. Selects a project folder via File System Access API.
3. Loads and edits `project_sidecar.json`.
4. Exports to Word/PPT via Excel (thin macro layer) or browser (if feasible).
5. Optional: materialize photo folders from tags.

## Immediate Next Steps

1. Validate File System Access API on the work laptop.
2. Define and lock the `project_sidecar.json` schema.
3. Add self-assessment import (SheetJS) into sidecar.
4. Add recommendation library load/save.
5. Build a photo-tagging spike once folder access is verified.
