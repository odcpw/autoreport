# AutoBericht: Current State vs Target

This document tracks the **redesign** direction (2026).

## Operating Constraints

- Offline-only (no CDN, no telemetry, no remote fonts/images).
- No unsafe browser flags.
- Locked-down corporate Edge environment.
- Prefer File System Access API for direct folder writes.

## Current State (Redesign)

- Minimal editor exists in `AutoBericht/mini/` with sidecar load/save and autosave.
- No-VBA Word export is available from the Project page (DOCX template markers, logos, spider image, thermo bars, priority values, chapter table payloads).
- PhotoSorter exists with stacked/tab layouts, observation tag add/remove, and filename display.
- 4.8 Beobachtungen is a special chapter (reorderable rows, tag‑driven cards, photo overlay per tag).
- Management Summary (Chapter 0) added with 8 placeholder cards.
- Markdown-lite preview supports bold/italic/bullets/links; tooltip cheatsheet added.
- I18n scaffold added (markdown tooltip wired; locale set from project meta).
- Shared debug log exporter is available across UI pages.
- File System Access + SheetJS spike exists in `AutoBericht/experiments/`.
- Seed recommendation guidance locked in `docs/research_recommendations/seed_prompt_template_v4.md` (natural prose, standalone paragraphs).
- Chapter 1 recommendations generated via Claude (one finding at a time) and written to `docs/research_recommendations/chapter1_recommendations_v5.md`.
- Chapter 1 recommendations applied to `AutoBericht/data/seed/knowledge_base_de.json` (library entries 1.1.1–1.5.7 updated).
- Chapters 2–14 recommendations generated via Claude and applied to `AutoBericht/data/seed/knowledge_base_de.json` (QA pass enforces 4 paragraphs, 6–9 sentences per paragraph).
- Chapter 11 recommendations generated in DE and applied to `AutoBericht/data/seed/knowledge_base_de.json` (4 paragraphs, 6–9 sentences each).
- FR-CH translation completed and applied to `AutoBericht/data/seed/knowledge_base_fr.json` (library aligned to DE; 261 entries).
- IT-CH translation completed and applied to `AutoBericht/data/seed/knowledge_base_it.json` (library aligned to DE; 261 entries).

## Target End-to-End Flow

1. User opens the minimal editor.
2. Selects a project folder via File System Access API.
3. Loads and edits `project_sidecar.json`.
4. Exports directly to Word (`.docx`) from the browser using the project template placeholders.
5. Optional: uses legacy Word/VBA pipeline where still required.
6. Optional: materialize photo folders from tags.

## Immediate Next Steps

- [x] No-VBA Word export (template marker replacement, chapter payload injection, logo insertion, spider image, thermo bars, priority column).
- [ ] Template hardening and governance (final placeholder map, style locks, and template QA pass).
- [ ] Explicit UI locale policy (lock English UI vs project-locale-driven UI text).
- [ ] Optional legacy VBA pipeline cleanup/retirement plan.
- [x] No File System Access fallback (removed by design).
- [x] Seed recommendations expanded for chapters 1–14 in DE/FR/IT libraries.
