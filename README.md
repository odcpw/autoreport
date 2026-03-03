# AutoBericht System

Offline safety-culture reporting tool with a browser UI (`AutoBericht/mini`) and project-local data (`project_sidecar.json`, library JSON, photo folders).

## Quick Start

1. Place this repo on local disk.
2. Run `start-autobericht.cmd` from repo root.
3. In the UI, click **Open Project Folder** and choose an empty project folder.
4. On **Project**, choose `Locale` to bootstrap seed content.
5. Use:
   - **PhotoSorter** for import/tag/export of photos
   - **AutoBericht** for findings/recommendations + chapter workflow
   - **Project** for Word/PPT export + library update/export

For onboarding guides, see:
- `instructions.txt`
- `docs/onboarding/AutoBericht-Schnellstart-DE.docx`
- `docs/onboarding/AutoBericht-Guide-Rapide-FR.docx`
- `docs/onboarding/AutoBericht-Guida-Rapida-IT.docx`

## Documentation

Primary redesign docs are in `docs/autobericht/`:
- `docs/autobericht/design-spec.md`
- `docs/autobericht/system-overview.md`
- `docs/autobericht/workflow.md`
- `docs/autobericht/project-template.md`

Additional research/docs:
- `docs/research_recommendations/`
- `docs/oracle-browser.md`

## Repo Structure

```text
autoreport/
├── README.md
├── instructions.txt
├── start-autobericht.cmd
├── AutoBericht/
│   ├── mini/                 # main web app (AutoBericht + PhotoSorter)
│   ├── data/                 # seeds, weights, checklists
│   ├── project-template/     # scaffold copied into new project folders
│   ├── libs/                 # bundled third-party libs (offline)
│   ├── tools/
│   └── experiments/
├── docs/
│   ├── autobericht/          # architecture/workflow docs
│   ├── onboarding/           # DE/FR/IT onboarding guides + screenshots
│   └── research_recommendations/
├── sources/
├── tools/
└── sync-autobericht.ps1
```

## Current Data Contract

- `project_sidecar.json`: project state (chapters, edits, filters, library action per row, photos state)
- `library_user_*.json`: reusable knowledge library (cross-project text/tag base)
- Library updates are applied when **Generate / Update Library** is run on Project page.

## Support

- Use **Save debug log** in the UI when troubleshooting.
- See `STATUS.md` for current state notes.
