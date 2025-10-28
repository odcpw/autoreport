# AutoBericht System

> Automated safety culture reporting system combining Excel VBA, offline web UI, and AI-powered quality analysis

## Overview

AutoBericht streamlines the creation of safety culture assessment reports by automating data ingestion, photo management, content editing, and multi-format export (PDF, PowerPoint). The system operates entirely offline to meet corporate security requirements.

### System Components

```
┌─────────────────────────────────────────────────────────────┐
│                    AutoBericht System                        │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐  │
│  │   Excel VBA  │───▶│  HTML/JS UI  │───▶│   Exports    │  │
│  │              │    │              │    │              │  │
│  │ • Data       │    │ • PhotoSorter│    │ • PDF Report │  │
│  │   Import     │    │ • Editor     │    │ • PPTX Deck  │  │
│  │ • Macro      │    │ • Settings   │    │ • JSON       │  │
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
│  ┌──────────────────────────────────────────────────────┐   │
│  │              vision-critique (Optional)              │   │
│  │  AI-powered UI quality analysis for development     │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### Key Features

- **Excel VBA Backend**: Import master findings and customer assessments, manage structured data
- **Offline Web UI**: Photo tagging, content editing, report preview (runs in browser without server)
- **Multi-format Export**: Generate PDF reports and PowerPoint presentations
- **Unified Data Model**: Single `project.json` format shared between Excel and web UI
- **Photo Management**: Organize and tag photos by chapter, category, and training topic
- **AI Quality Analysis**: Optional vision-based UI critique tool for development

## Quick Start

### For Report Creators

1. **Prepare Data**
   - Open Excel workbook (`project.xlsm`)
   - Import master findings (Word document)
   - Import customer self-assessment (Excel file)

2. **Edit Content**
   - Open `AutoBericht/index.html` in Chrome/Edge
   - Tag photos in PhotoSorter tab
   - Edit findings and recommendations in AutoBericht tab
   - Configure export settings

3. **Export Reports**
   - Generate PDF report (print from browser)
   - Export PowerPoint presentations
   - Save project snapshot as JSON

**See**: [Getting Started Guide](docs/guides/getting-started.md)

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
   - Open `AutoBericht/index.html` in browser
   - Libraries bundled in `AutoBericht/libs/`
   - See [AutoBericht README](AutoBericht/README.md)

4. **UI Quality Analysis** (optional)
   ```bash
   cd vision-critique
   uv sync
   playwright install chromium
   uv run vision-critique capture --tab photosorter
   ```

**See**: [System Architecture](docs/architecture/system-overview.md)

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
├── README.md                    # You are here
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
└── vision-critique/             # UI quality analysis tool
    └── README.md               # Tool documentation
```

## Tech Stack

### Excel VBA
- **Purpose**: Data import, structured sheet management, macro synchronization
- **Key Libraries**: VBA-JSON, Microsoft Scripting Runtime
- **Entry Point**: `modMacroSync.RefreshProjectMacros`

### HTML/JavaScript
- **Purpose**: Offline editing interface
- **Key Libraries**: markdown-it, CodeMirror, Paged.js, PptxGenJS, SheetJS
- **Entry Point**: `AutoBericht/index.html`

### Vision Critique (Optional)
- **Purpose**: AI-powered UI quality analysis
- **Tech**: Python, Playwright, Claude/GPT-4V/Ollama
- **Entry Point**: `vision-critique` CLI tool

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

## Contributing

This is a private corporate project. For development guidelines:

- VBA: Follow Rubberduck style guide
- JavaScript: ES modules, no bundler
- Documentation: Keep it simple and actionable

## Support

- **Issues**: Check documentation first, then consult team
- **VBA Help**: See [VBA Workflow](docs/guides/vba-workflow.md)
- **Web UI Help**: See [AutoBericht README](AutoBericht/README.md)
- **Vision Critique**: See [Tool README](vision-critique/README.md)

## License

Internal use only. All rights reserved.

---

**Next Steps**: Read the [Documentation Index](docs/INDEX.md) or jump to [Getting Started](docs/guides/getting-started.md)
