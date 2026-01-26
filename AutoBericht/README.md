# AutoBericht UI (Redesign)

This folder defaults to the minimal, folder-first editor in `mini/`.

> **Tip (recommended):** Use the launcher in this folder (`start-autobericht.cmd`). It starts
> a local, offline-only server and opens the minimal editor without browser flags.

## Directory Layout

- `index.html` — landing page linking to the editor and experiments.
- `mini/` — minimal chapter editor (current direction).
- `experiments/` — focused spikes (FS Access, SheetJS, etc.).
- `shared/` — shared utilities (debug logger).
- `vba/` — Word/PPT VBA export modules.
- `data/` — seeds, checklists, weights.
- `libs/` — bundled third-party libraries (offline only).

## Getting Started (New)

1. Launch the local server:
   ```
   start-autobericht.cmd
   ```
2. The server opens the minimal editor at `/mini/`.
3. Use **Open Project Folder** (sidecar loads or is created automatically).
4. Save with **Save sidecar**. Use **Save debug log** for troubleshooting.

## Related Documentation

- **Design Spec:** `docs/design-spec.md`
- **System Overview:** `docs/system-overview.md`
- **Workflow:** `docs/workflow.md`
