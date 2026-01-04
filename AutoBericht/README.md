# AutoBericht UI (Redesign)

This folder now defaults to the minimal, folder-first editor in `mini/`. The previous
three-tab MVP UI lives under `legacy/` for reference.

> **Tip (recommended):** Use the repo launcher (`start-autobericht.cmd`). It starts
> a local, offline-only server and opens the minimal editor without browser flags.

## Directory Layout

- `index.html` — landing page linking to the new editor, experiments, and legacy UI.
- `mini/` — minimal chapter editor (current direction).
- `experiments/` — focused spikes (FS Access, SheetJS, etc.).
- `shared/` — shared utilities (debug logger).
- `legacy/` — legacy MVP UI (three-panel, import/export tabs).
- `libs/` — bundled third-party libraries (offline only).

## Getting Started (New)

1. Launch the local server:
   ```
   start-autobericht.cmd
   ```
2. The server opens the minimal editor at `/mini/`.
3. Use **Open Project Folder** and **Load sidecar**.
4. Save with **Save sidecar**. Use **Save debug log** for troubleshooting.

## Legacy UI

The old MVP UI remains under `legacy/` for reference and fallback. It is not the
active direction and may diverge from the redesign spec.

## Related Documentation

- **[Redesign Spec](../docs/architecture/redesign-spec-2026-01-04.md)**
- **[Redesign Workflow](../docs/guides/redesign-workflow.md)**
- **[System Overview (Legacy)](../docs/architecture/system-overview.md)**
