# AutoBericht Documentation Index

Welcome to the AutoBericht documentation. This index helps you find the right information for your role and task.

## Project Status

- **Current state vs target**: [STATUS](STATUS.md)

## Documentation by Role

### 👤 Report Creators (Non-Technical)

Start here if you create safety assessment reports:

1. **[Getting Started](guides/getting-started.md)** - Your first report from start to finish
2. **[VBA Workflow](guides/vba-workflow.md)** - Using Excel macros for data import
3. **[HTML Workflow](guides/html-workflow.md)** - Using the web interface for editing

### 👨‍💻 Developers (Technical)

Start here if you're developing or maintaining the system:

1. **[System Overview](architecture/system-overview.md)** - How all components connect
2. **[Data Model](architecture/data-model.md)** - JSON schema and Excel structure
3. **[VBA Architecture](architecture/vba-architecture.md)** - Excel macro design
4. **[VBA Modules Reference](../macros/README.md)** - Detailed module documentation
5. **[PowerShell Sync](reference/powershell-sync.md)** - Macro distribution system

### 🔧 UI Developers

Start here if you're working on the web interface:

1. **[AutoBericht UI](../AutoBericht/README.md)** - Web interface overview
2. **[System Overview](architecture/system-overview.md)** - Integration points
3. **[Data Model](architecture/data-model.md)** - `project.json` contract

## Documentation by Topic

### 📐 Architecture & Design

| Document | What You'll Learn |
|----------|-------------------|
| [System Overview](architecture/system-overview.md) | How VBA, web UI, and exports work together |
| [VBA Architecture](architecture/vba-architecture.md) | Excel macro system design principles |
| [Data Model](architecture/data-model.md) | JSON schema, Excel sheets, and data flow |

### 📖 User Guides

| Document | What You'll Learn |
|----------|-------------------|
| [Getting Started](guides/getting-started.md) | Create your first report end-to-end |
| [VBA Workflow](guides/vba-workflow.md) | Import data, manage macros, export JSON |
| [HTML Workflow](guides/html-workflow.md) | Edit content, tag photos, generate exports |

### 📚 Technical Reference

| Document | What You'll Learn |
|----------|-------------------|
| [VBA Modules](../macros/README.md) | All VBA modules, functions, and their purpose |
| [JSON Schema](reference/json-schema.md) | Complete data structure specification |
| [PowerShell Sync](reference/powershell-sync.md) | Macro distribution and versioning |
| [Round-Trip Smoke Checklist](reference/roundtrip-smoke-checklist.md) | Quick verification of Excel ↔ Web UI workflow |
| [project.schema.json](reference/project.schema.json) | Machine-readable schema (VBA-canonical) |

### 📦 Component Documentation

| Component | Documentation |
|-----------|---------------|
| Excel VBA | [VBA Modules Reference](../macros/README.md) |
| Web UI | [AutoBericht README](../AutoBericht/README.md) |

### 🗄️ Legacy & Planning

Historical documents and execution plans:

| Document | Purpose |
|----------|---------|
| [Excel Migration Notes](legacy/excel-migration-notes.md) | Legacy workbook analysis and migration strategy |
| [Execution Plan](legacy/execution-plan.md) | Original MVP implementation roadmap |

## Common Tasks

### "I need to..."

#### Create a report
→ Start with [Getting Started Guide](guides/getting-started.md)

#### Import master findings
→ See [VBA Workflow - Section 2.1](guides/vba-workflow.md#21-master-findings-word--rows)

#### Tag photos
→ Use Excel PhotoSorter (see [VBA Workflow](guides/vba-workflow.md#4-photosorter))

#### Export to PDF/PowerPoint
→ See [HTML Workflow - Export Section](guides/html-workflow.md#export-overview)

#### Understand the data structure
→ See [Data Model](architecture/data-model.md)

#### Develop VBA modules
→ See [VBA Architecture](architecture/vba-architecture.md) and [VBA Modules Reference](../macros/README.md)

#### Sync macros from GitHub
→ See [PowerShell Sync](reference/powershell-sync.md)

#### Understand how everything connects
→ See [System Overview](architecture/system-overview.md)

## Documentation Structure

```
docs/
├── INDEX.md (you are here)          # Documentation roadmap
│
├── architecture/                     # System design
│   ├── system-overview.md           # Component integration
│   ├── vba-architecture.md          # VBA design principles
│   └── data-model.md                # JSON and Excel structure
│
├── guides/                           # How-to guides
│   ├── getting-started.md           # First-time user guide
│   ├── vba-workflow.md              # Excel macro usage
│   └── html-workflow.md             # Web UI usage
│
├── reference/                        # Technical reference
│   ├── json-schema.md               # Complete schema spec
│   ├── project.schema.json          # Machine-readable schema (VBA-canonical)
│   ├── powershell-sync.md           # Macro distribution
│   └── roundtrip-smoke-checklist.md # Excel ↔ Web UI smoke checklist
│
└── legacy/                           # Historical docs
    ├── excel-migration-notes.md     # Legacy analysis
    └── execution-plan.md            # Original roadmap
```

## Quick Reference

### Key Concepts

- **Structured Sheets**: Excel sheets that map directly to JSON (Meta, Chapters, Rows, Photos, Lists)
- **project.json**: Unified data format shared between VBA and web UI
- **Master Findings**: Template content from Word document
- **Customer Assessment**: Self-evaluation data from Excel
- **Overrides**: Custom consultant edits to template content
- **PhotoSorter**: Photo tagging and organization interface
  - Canonical photo tagging UI lives in Excel (PhotoSorterForm).

### Key Files

- `project.xlsm` - Excel workbook with VBA macros
- `project.json` - Unified data export/import format
- `AutoBericht/index.html` - Web UI entry point
- `macros/*.bas` - VBA source modules

## Getting Help

1. **Check this index** for the right document
2. **Read the relevant guide or reference**
3. **Search for your specific topic** in the docs
4. **Consult your team** if still unclear

## Contributing to Documentation

When adding or updating docs:

- Keep language simple and actionable
- Include examples and code samples
- Add cross-references to related topics
- Update this index when adding new docs
- Follow the structure: guides (how-to) vs reference (what is)

---

**Ready to start?** Choose your path:
- 👤 [Report Creators → Getting Started](guides/getting-started.md)
- 👨‍💻 [Developers → System Overview](architecture/system-overview.md)
- 🔧 [UI Developers → AutoBericht UI](../AutoBericht/README.md)
