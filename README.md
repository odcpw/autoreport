# AutoBericht System

> Automated safety culture reporting system combining Word VBA and an offline web UI

## Overview

AutoBericht streamlines the creation of safety culture assessment reports by automating data ingestion, photo management, content editing, and multi-format export (PDF, PowerPoint). The system operates entirely offline to meet corporate security requirements.

## Redesign Status (2026)

- The minimal editor in `AutoBericht/mini/` is the current direction.
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
│        │  SheetJS DB  │   │ Word/VBA    │                      │
│        │ project_db   │   │ (thin export)│                      │
│        └──────────────┘   └──────────────┘                      │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### Key Features

- **Browser-first editing** with sidecar JSON stored in the project folder.
- **Offline, policy-safe** workflow (no flags, no installs).
- **Thin Word/VBA export layer** for Word/PPT templates (optional).
- **File System Access API** preferred for direct folder writes.

## Quick Start

### For Report Creators

1. **Prepare Data**
   - Open Excel workbook (`project.xlsm`)
   - Import master findings (Word document)
   - Import customer self-assessment (Excel file)

2. **Edit Content**
   - Start the offline UI with `AutoBericht/start-autobericht.cmd` (opens the minimal editor)
   - Open the project folder (sidecar loads or is created automatically)
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
   - VBA modules live in `AutoBericht/vba/`
   - See `AutoBericht/docs/ribbon-notes.md` for template notes

3. **Web UI Development**
   - Open `AutoBericht/mini/index.html` (current)
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
| [Project Template](AutoBericht/docs/project-template.md) | Recommended project folder layout |
| [Ribbon Notes](AutoBericht/docs/ribbon-notes.md) | Notes on Word ribbon + template setup |

## Project Structure

```
autoreport/
├── README.md                    # You are here
├── AutoBericht/start-autobericht.cmd   # Windows launcher for the web UI
├── AutoBericht/start-autobericht.ps1   # PowerShell launcher for the web UI
├── sync-autobericht.ps1         # Macro sync helper
├── STATUS.md                    # Current state vs target
├── docs/                        # Research / prompts / tooling notes
│   └── research_recommendations
├── AutoBericht/                 # Offline web UI + docs
│   ├── index.html              # Entry point
│   ├── mini/                   # Minimal editor (current)
│   ├── vba/                    # Word/PPT VBA export modules
│   ├── data/                   # Seeds, checklists, weights
│   ├── libs/                   # Bundled third-party libraries (SheetJS)
│   └── docs/                   # Canonical redesign documentation
├── ProjectTemplate/             # Sample project folder (templates + structure)
└── tools/                       # Local tooling scripts
    └── serve-autobericht.ps1    # Local server (offline-only)
```

## Tech Stack

### Word VBA
- **Purpose**: Data import, structured sheet management, macro synchronization
- **Key Libraries**: VBA-JSON, Microsoft Scripting Runtime
- **Modules**: `AutoBericht/vba/`

### HTML/JavaScript
- **Purpose**: Offline editing interface
- **Key Libraries**: SheetJS (bundled in `AutoBericht/libs/`)
- **Entry Point**: `AutoBericht/mini/index.html` (current) or `AutoBericht/index.html` (landing)

## Data Flow

```
Word/Excel Sources
       │
       ▼
   VBA Import ──────────────┐
       │                    │
       ▼                    ▼
Structured Sheets      project_sidecar.json
       │                    │
       ▼                    ▼
   VBA Export          Web UI Load (project_sidecar.json)
       │                    │
       │◀───────────────────┘
       │
       ▼
   Exports (PDF/PPTX/JSON)
```

1. **Import**: VBA reads Word/Excel sources, populates structured sheets
2. **Edit**: Web UI loads `project_sidecar.json` and edits content
3. **Export**: Generate reports and update `project_sidecar.json`

**See**: [Design Spec](AutoBericht/docs/design-spec.md)

**Canonical contract**: redesign uses `project_sidecar.json` as the working state (see `AutoBericht/docs/design-spec.md` and `STATUS.md`).

## Contributing

This is a private corporate project. For development guidelines:

- VBA: Follow Rubberduck style guide
- JavaScript: ES modules, no bundler
- Documentation: Keep it simple and actionable

## Support

- **Issues**: Check documentation first, then consult team
- **VBA Help**: See `AutoBericht/docs/ribbon-notes.md` or `AutoBericht/vba/`
- **Web UI Help**: See [AutoBericht README](AutoBericht/README.md)

## License

Internal use only. All rights reserved.

---

**Next Steps**: Read the [Status](STATUS.md) or jump to [Workflow](AutoBericht/docs/workflow.md)
