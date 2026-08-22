# AutoBericht UI

This folder contains the offline web app.

## Main Areas

- `mini/` — production UI (`index.html` for AutoBericht, `photosorter.html` for PhotoSorter)
- `data/` — seeds, checklists, weights
- `project-template/` — scaffold copied into a new empty project folder
- `libs/` — bundled dependencies (offline)
- `experiments/` — non-production spikes

## Start

Recommended: run from repo root:

```bash
start-autobericht.cmd
```

This calls `AutoBericht/start-autobericht.ps1`, starts a local server, and opens `mini/`.

## Test

The production browser modules have dependency-free regression tests:

```bash
cd AutoBericht
npm test
```

The suite covers sidecar persistence/conflicts, startup and locale bootstrap, self-assessment validation, safe Markdown, ZIP limits, and the real DOCX export/template boundary. The DOCX checks require the `ooxml` CLI on `PATH`.

## Related Docs

- `../docs/autobericht/design-spec.md`
- `../docs/autobericht/system-overview.md`
- `../docs/autobericht/workflow.md`
- `../docs/onboarding/`
