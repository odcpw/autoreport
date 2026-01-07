# AutoBericht: Current State vs Target

This document tracks the **redesign** direction (2026).

## Operating Constraints

- Offline-only (no CDN, no telemetry, no remote fonts/images).
- No unsafe browser flags.
- Locked-down corporate Edge environment.
- Prefer File System Access API for direct folder writes.

## Current State (Redesign)

- Minimal editor exists in `AutoBericht/mini/` with sidecar load/save and autosave.
- PhotoSorter exists with stacked/tab layouts, observation tag add/remove, and filename display.
- 4.8 Beobachtungen is a special chapter (reorderable rows, tag‑driven cards, photo overlay per tag).
- Management Summary (Chapter 0) added with 8 placeholder cards.
- Markdown-lite preview supports bold/italic/bullets/links; tooltip cheatsheet added.
- I18n scaffold added (markdown tooltip wired; locale set from project meta).
- Shared debug log exporter is available across UI pages.
- File System Access + SheetJS spike exists in `AutoBericht/experiments/`.

## Target End-to-End Flow

1. User opens the minimal editor.
2. Selects a project folder via File System Access API.
3. Loads and edits `project_sidecar.json`.
4. Exports to Word/PPT via Word macro (content controls) reading sidecar JSON.
5. Optional: materialize photo folders from tags.

## Immediate Next Steps

1. Word template: add bookmarks/content controls + styles per chapter.
2. Word VBA importer: read sidecar JSON, inject content, apply styles, renumber.
3. Export rules: handle 4.8 ordering + photo lists in Word output.
4. Spider chart data export (per‑chapter scores).
5. Optional: File System Access fallback (manual import/export).
