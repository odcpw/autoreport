# AutoBericht System

> Automated safety culture reporting system combining Excel VBA and an offline web UI

## Overview

AutoBericht streamlines the creation of safety culture assessment reports by automating data ingestion, photo management, content editing, and multi-format export (PDF, PowerPoint). The system operates entirely offline to meet corporate security requirements.

## Redesign Status (2026)

- The minimal editor in `AutoBericht/mini/` is the current direction.
- The legacy MVP UI now lives in `AutoBericht/legacy/` (reference only).
- See `AutoBericht/docs/design-spec.md` for the interview-based spec.

### System Components

```
┌─────────────────────────────────────────────────────────────┐
│                    AutoBericht (Redesign)                   │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌──────────────┐    ┌──────────────────┐    ┌──────────────┐ │
│  │  Browser UI  │───▶│ project_sidecar  │───▶│  Exports     │ │
│  │  (mini)      │◀───│ (canonical JSON) │    │ Word/PPT/PDF │ │
│  └──────────────┘    └──────────────────┘    └──────────────┘ │
│             ▲                    │                             │
│             │ optional           ▼                             │
│        ┌──────────────┐   ┌──────────────┐                      │
│        │  SheetJS DB  │   │ Excel/VBA    │                      │
│        │ project_db   │   │ (thin export)│                      │
│        └──────────────┘   └──────────────┘                      │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### Key Features

- **Browser-first editing** with sidecar JSON stored in the project folder.
- **Offline, policy-safe** workflow (no flags, no installs).
- **Thin Excel/VBA export layer** for Word/PPT templates (optional).
- **File System Access API** preferred for direct folder writes.

## Quick Start

### For Report Creators

1. **Prepare Data**
   - Open Excel workbook (`project.xlsm`)
   - Import master findings (Word document)
   - Import customer self-assessment (Excel file)

2. **Edit Content**
   - Start the offline UI with `AutoBericht/start-autobericht.cmd` (opens the minimal editor)
   - Open the project folder and load `project_sidecar.json`
   - Edit findings and recommendations by chapter
   - Save the sidecar back to the project folder

3. **Export Reports**
   - Generate PDF report (print from browser)
   - Export PowerPoint presentations
   - Save project snapshot as JSON

**See**: [Workflow](AutoBericht/docs/workflow.md)

### For Developers

1. **Clone Repository**
   ```bash
   git clone https://github.com/odcpw/autoreport.git
   cd autoreport
   ```

2. **VBA Development**
   - Install VBA modules from `/legacy/macros/` folder
   - See legacy docs in `legacy/docs/guides/vba-workflow.md`

3. **Web UI Development**
   - Open `AutoBericht/mini/index.html` (current) or `AutoBericht/legacy/index.html` (legacy)
   - Libraries bundled in `AutoBericht/libs/`
   - See [AutoBericht README](AutoBericht/README.md)

**See**: [Design Spec](AutoBericht/docs/design-spec.md)

## Documentation

📚 **[Status & Docs →](STATUS.md)**

### Key Documents

| Document | Description |
|----------|-------------|
| [System Architecture](AutoBericht/docs/system-overview.md) | Current redesign architecture |
| [Design Spec](AutoBericht/docs/design-spec.md) | Interview-based requirements |
| [Workflow](AutoBericht/docs/workflow.md) | Minimal editor workflow |
| [VBA Modules Reference](legacy/macros/README.md) | VBA module documentation (legacy export layer) |

Legacy documentation is archived under `legacy/docs/`.

## Project Structure

```
autoreport/
├── .beads/                     # Beads task tracker (issues.jsonl)
├── AGENTS.md                   # Agent playbook for this repo
├── README.md                    # You are here
├── AutoBericht/start-autobericht.cmd   # Windows launcher for the web UI
├── AutoBericht/start-autobericht.ps1   # PowerShell launcher for the web UI
├── sync-autobericht.ps1         # Macro sync helper
├── docs/                        # Documentation
│   ├── INDEX.md                # Documentation roadmap (current)
│   ├── STATUS.md               # Current state vs target
│   └── (canonical docs live in AutoBericht/docs/)
├── legacy/                      # Legacy UI + docs archive
│   ├── docs/                   # Legacy documentation
│   └── macros/                 # Legacy VBA modules (.bas, .cls, .frm)
│       └── README.md           # VBA module overview
├── AutoBericht/                 # Offline web UI
│   ├── index.html              # Entry point
│   ├── css/                    # Styles
│   ├── js/                     # JavaScript modules
│   ├── libs/                   # Third-party libraries
│   └── assets/                 # Images, icons
└── tools/                       # Local tooling scripts
    └── serve-autobericht.ps1    # Local server (offline-only)
```

## Tech Stack

### Excel VBA
- **Purpose**: Data import, structured sheet management, macro synchronization
- **Key Libraries**: VBA-JSON, Microsoft Scripting Runtime
- **Entry Point**: `modMacroSync.RefreshProjectMacros`

### HTML/JavaScript
- **Purpose**: Offline editing interface
- **Key Libraries**: markdown-it, CodeMirror, Paged.js, PptxGenJS, SheetJS
- **Entry Point**: `AutoBericht/mini/index.html` (current) or `AutoBericht/index.html` (landing)

## Data Flow

```
Word/Excel Sources
       │
       ▼
   VBA Import ──────────────┐
       │                    │
       ▼                    ▼
Structured Sheets      project.json
       │                    │
       ▼                    ▼
   VBA Export          Web UI Load
       │                    │
       │◀───────────────────┘
       │
       ▼
   Exports (PDF/PPTX/JSON)
```

1. **Import**: VBA reads Word/Excel sources, populates structured sheets
2. **Edit**: Web UI loads `project.json` or Excel file directly
3. **Export**: Generate reports and update `project.json`

**See**: [Design Spec](AutoBericht/docs/design-spec.md)

**Canonical contract**: redesign uses `project_sidecar.json` as the working state (see `AutoBericht/docs/design-spec.md` and `docs/STATUS.md`). Legacy `project.json` docs are archived in `legacy/docs/architecture/data-model.md`.

## Contributing

This is a private corporate project. For development guidelines:

- VBA: Follow Rubberduck style guide
- JavaScript: ES modules, no bundler
- Documentation: Keep it simple and actionable

## Support

- **Issues**: Check documentation first, then consult team
- **VBA Help**: See `legacy/docs/guides/vba-workflow.md`
- **Web UI Help**: See [AutoBericht README](AutoBericht/README.md)

## License

Internal use only. All rights reserved.

---

**Next Steps**: Read the [Status](STATUS.md) or jump to [Workflow](AutoBericht/docs/workflow.md)
