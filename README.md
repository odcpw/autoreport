# AutoBericht System

> Automated safety culture reporting system combining Excel VBA and an offline web UI

## Overview

AutoBericht streamlines the creation of safety culture assessment reports by automating data ingestion, photo management, content editing, and multi-format export (PDF, PowerPoint). The system operates entirely offline to meet corporate security requirements.

## Redesign Status (2026)

- The minimal editor in `AutoBericht/mini/` is the current direction.
- The legacy MVP UI now lives in `AutoBericht/legacy/` (reference only).
- See `docs/architecture/redesign-spec-2026-01-04.md` for the interview-based spec.

### System Components

```
┌─────────────────────────────────────────────────────────────┐
│                    AutoBericht System                        │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐  │
│  │   Excel VBA  │───▶│  HTML/JS UI  │───▶│   Exports    │  │
│  │              │    │              │    │              │  │
│  │ • Data       │    │ • Import     │    │ • PDF Report │  │
│  │   Import     │    │ • Editor     │    │ • PPTX Deck  │  │
│  │ • Macro      │    │ • Export     │    │ • JSON       │  │
│  │   Sync       │    │              │    │              │  │
│  └──────────────┘    └──────────────┘    └──────────────┘  │
│         │                    │                               │
│         └────────────────────┴───────────────┐              │
│                                               │              │
│                                      ┌────────▼────────┐     │
│                                      │  project.json   │     │
│                                      │  (Data Model)   │     │
│                                      └─────────────────┘     │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### Key Features

- **Excel VBA Backend**: Import master findings and customer assessments, manage structured data
- **Offline Web UI**: Content editing and report generation (photo tagging happens in Excel PhotoSorter)
- **Multi-format Export**: Generate PDF reports and PowerPoint presentations
- **Unified Data Model**: Single `project.json` format shared between Excel and web UI
- **Photo Management (Excel)**: Organize and tag photos by chapter, category, and training topic (PhotoSorterForm)

## Quick Start

### For Report Creators

1. **Prepare Data**
   - Open Excel workbook (`project.xlsm`)
   - Import master findings (Word document)
   - Import customer self-assessment (Excel file)

2. **Edit Content**
   - Start the offline UI with `start-autobericht.cmd` (opens the minimal editor)
   - Open the project folder and load `project_sidecar.json`
   - Edit findings and recommendations by chapter
   - Save the sidecar back to the project folder

3. **Export Reports**
   - Generate PDF report (print from browser)
   - Export PowerPoint presentations
   - Save project snapshot as JSON

**See**: [Redesign Workflow](docs/guides/redesign-workflow.md)

### For Developers

1. **Clone Repository**
   ```bash
   git clone https://github.com/odcpw/autoreport.git
   cd autoreport
   ```

2. **VBA Development**
   - Install VBA modules from `/macros/` folder
   - See [VBA Workflow Guide](docs/guides/vba-workflow.md)

3. **Web UI Development**
   - Open `AutoBericht/mini/index.html` (current) or `AutoBericht/legacy/index.html` (legacy)
   - Libraries bundled in `AutoBericht/libs/`
   - See [AutoBericht README](AutoBericht/README.md)

**See**: [Redesign Spec](docs/architecture/redesign-spec-2026-01-04.md)

## Documentation

📚 **[Complete Documentation Index →](docs/INDEX.md)**

### Key Documents

| Document | Description |
|----------|-------------|
| [System Architecture](docs/architecture/system-overview.md) | How all components work together |
| [Data Model](docs/architecture/data-model.md) | JSON schema and Excel sheets structure |
| [VBA Workflow](docs/guides/vba-workflow.md) | Excel macro usage and development |
| [VBA Modules Reference](macros/README.md) | Detailed module documentation |
| [HTML Workflow](docs/guides/html-workflow.md) | Web UI usage and features |

## Project Structure

```
autoreport/
├── .beads/                     # Beads task tracker (issues.jsonl)
├── AGENTS.md                   # Agent playbook for this repo
├── README.md                    # You are here
├── start-autobericht.cmd        # Windows launcher for the web UI
├── start-autobericht.ps1        # PowerShell launcher for the web UI
├── sync-autobericht.ps1         # Macro sync helper
├── docs/                        # Documentation
│   ├── INDEX.md                # Documentation roadmap
│   ├── architecture/           # System design
│   ├── guides/                 # How-to guides
│   ├── reference/              # Technical reference
│   └── legacy/                 # Historical/planning docs
├── macros/                      # VBA modules (.bas, .cls, .frm)
│   └── README.md               # VBA module overview
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

**See**: [Data Model Documentation](docs/architecture/data-model.md)

**Canonical contract**: legacy flow uses `project.json` as documented in `docs/architecture/data-model.md`. The redesign shifts canonical state to `project_sidecar.json` (see `docs/architecture/redesign-spec-2026-01-04.md` and `docs/STATUS.md`).

## Contributing

This is a private corporate project. For development guidelines:

- VBA: Follow Rubberduck style guide
- JavaScript: ES modules, no bundler
- Documentation: Keep it simple and actionable

## Support

- **Issues**: Check documentation first, then consult team
- **VBA Help**: See [VBA Workflow](docs/guides/vba-workflow.md)
- **Web UI Help**: See [AutoBericht README](AutoBericht/README.md)

## License

Internal use only. All rights reserved.

---

**Next Steps**: Read the [Documentation Index](docs/INDEX.md) or jump to [Getting Started](docs/guides/getting-started.md)
