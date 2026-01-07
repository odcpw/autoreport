# AutoBericht System Architecture Overview (Legacy)

> This overview reflects the pre-redesign architecture. For the 2026 redesign
> direction, see `AutoBericht/docs/design-spec.md`.

## Purpose

This document explains how all AutoBericht components work together to create a complete safety assessment reporting system. It bridges the VBA backend, the offline web UI, and the `project.json` data model.

## System Diagram

```
┌────────────────────────────────────────────────────────────────────┐
│                     AutoBericht Ecosystem                           │
└────────────────────────────────────────────────────────────────────┘

┌─────────────────┐              ┌──────────────────┐
│  Data Sources   │              │   Excel Workbook │
├─────────────────┤              │  (project.xlsm)  │
│ • Word Master   │─────────────▶│                  │
│ • Excel Self-   │   VBA Import │ Structured       │
│   Assessment    │              │ Sheets:          │
│ • Photos Folder │              │ • Meta           │
└─────────────────┘              │ • Chapters       │
                                 │ • Rows           │
                                 │ • Photos         │
                                 │ • Lists          │
                                 │ • History        │
                                 └────────┬─────────┘
                                          │
                                          │ VBA Export
                                          ▼
                                 ┌────────────────────┐
                                 │   project.json     │
                                 │  (Unified Format)  │
                                 └────────┬───────────┘
                                          │
                                          │ Load/Save
                                          ▼
┌────────────────────────────────────────────────────────────┐
│              Legacy Offline Web UI (AutoBericht/legacy/)                  │
├────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐     │
│  │   Import     │  │ AutoBericht  │  │    Export    │     │
│  │              │  │              │  │              │     │
│  │ • Load       │  │ • Edit       │  │ • PDF/PPTX   │     │
│  │   workbook   │  │   findings   │  │ • Download   │     │
│  │ • Validate   │  │ • Overrides  │  │   project    │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │              Export Generators                        │  │
│  │  • PDF Report (Paged.js)                            │  │
│  │  • PowerPoint Decks (PptxGenJS)                     │  │
│  │  │  • Report PPTX                                    │  │
│  │  │  • Seminar PPTX                                   │  │
│  │  • project.json snapshot                            │  │
│  └──────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────┘
                          │
                          │ Generated Files
                          ▼
                 ┌────────────────────┐
                 │   Deliverables     │
                 │ • report.pdf       │
                 │ • report.pptx      │
                 │ • training.pptx    │
                 │ • project.json     │
                 └────────────────────┘
```

## Component Responsibilities

### 1. Excel VBA Backend

**Location**: `/macros/` folder, imported into `project.xlsm`

**Purpose**: Data ingestion, structured storage, macro synchronization

**Key Responsibilities**:
- Import master findings from Word documents
- Import customer self-assessment from Excel files
- Manage structured sheets (Meta, Chapters, Rows, Photos, Lists, History)
- Export unified `project.json` format
- Synchronize macro updates from GitHub

**Key Modules**:
- `modABProjectLoader` - Load JSON into structured sheets
- `modABProjectExport` - Export sheets to JSON
- `ImportMyMaster` - Word document parser
- `ImportSelbstbeurteilungKunde` - Customer assessment parser
- `modMacroSync` - GitHub sync and module refresh

**See**: [VBA Architecture](vba-architecture.md), [VBA Workflow Guide](../guides/vba-workflow.md)

### 2. Offline Web UI

**Location**: `/AutoBericht/legacy/` folder

**Purpose**: Content editing and report generation. Photo tagging is performed in Excel (PhotoSorterForm); the web UI consumes the resulting metadata via `project.json` / workbook import.

**Key Features**:
- **Import Tab**: Load `project.json`/workbook/photos, edit tag lists, run validation
- **AutoBericht Tab**: Edit findings/recommendations, levels, overrides
- **Export Tab**: Generate PDF/PPTX outputs and download `project.json`

**Technology**:
- Pure HTML/CSS/JavaScript (ES modules)
- No build step, no server required
- Runs from `file://` protocol or local server
- All libraries bundled offline

**Key Libraries**:
- markdown-it (Markdown rendering)
- CodeMirror 6 (code editor)
- Paged.js (PDF generation)
- PptxGenJS (PowerPoint generation)
- SheetJS (Excel file handling)

**See**: [AutoBericht README](../../AutoBericht/README.md), [HTML Workflow Guide](../guides/html-workflow.md)

### 3. Data Model (project.json)

**Purpose**: Unified data format shared between VBA and web UI

**Canonical contract**: `legacy/docs/architecture/data-model.md`. The Excel/VBA exporter/loader defines the canonical JSON shape; the web UI adapts to it. Current progress/mismatches live in `docs/STATUS.md`.

**Structure**:
```json
{
  "version": 1,
  "meta": { projectId, company, locale, createdAt, author },
  "chapters": [
    {
      "id": "1.1",
      "title": {...},
      "rows": [
        {
          "id": "1.1.1",
          "master": { "finding": "...", "levels": { "1": "...", "2": "...", "3": "...", "4": "..." } },
          "overrides": { "finding": { "text": "...", "enabled": false }, "levels": { "1": { "text": "...", "enabled": false } } },
          "customer": { answer, remark, priority },
          "workstate": { selectedLevel, includeFinding, includeRecommendation, done }
        }
      ]
    }
  ],
  "photos": { filename: { notes, tags } },
  "lists": { listName: [ items ] },
  "history": [ override changes ]
}
```

**Key Characteristics**:
- One-to-one mapping with Excel structured sheets
- Hierarchical chapter structure
- Separation of master content, customer input, and consultant edits
- Photo tagging by multiple dimensions

**See**: [Data Model Documentation](data-model.md)

## Data Flow

### Import Flow

```
1. Prepare Sources
   ├─ Word document with master findings
   ├─ Excel file with customer self-assessment
   └─ Folder with photos

2. VBA Import
   ├─ ImportMyMaster.bas reads Word tables
   ├─ ImportSelbstbeurteilungKunde.bas reads Excel
   └─ PhotoSorterForm scans/tags photo directory

3. Structured Sheets
   ├─ Meta: project metadata
   ├─ Chapters: report structure
   ├─ Rows: findings and recommendations
   ├─ Photos: filenames and metadata
   └─ Lists: tag vocabularies

4. Export to JSON
   └─ modABProjectExport.ExportProjectJson()
```

### Edit Flow

```
1. Load Project
   ├─ Web UI reads project.json
   └─ OR: Web UI reads project.xlsm directly (SheetJS)

2. Edit Content
   ├─ Excel PhotoSorter: tag photos (writes Photos/PhotoTags/Lists)
   ├─ Web UI AutoBericht: edit findings, set levels, add overrides
   └─ Web UI Import: validation + tag list edits

3. Validate
   ├─ Check required fields
   ├─ Validate tags against lists
   └─ Ensure export readiness

4. Save Changes
   ├─ Update project.json
   └─ OR: Update Excel structured sheets
```

### Export Flow

```
1. Generate Outputs
   ├─ PDF Report (Paged.js → browser print)
   ├─ Report PPTX (PptxGenJS)
   ├─ Seminar PPTX (PptxGenJS)
   └─ project.json snapshot

2. Optional: VBA Post-Processing
   ├─ Read ExportLog sheet
   ├─ Renumber photo files
   └─ Organize deliverables
```

## Integration Points

### Excel ↔ Web UI

**Bidirectional via `project.json`**:
- VBA exports JSON → Web UI imports
- Web UI exports JSON → VBA imports

**Direct Excel file handling**:
- Web UI can read `.xlsm` using SheetJS
- Web UI can write `.xlsm` preserving VBA macros

### VBA ↔ GitHub

**PowerShell sync script**:
- Downloads latest macros from repository
- VBA `modMacroSync` imports modules automatically
- Enables version control without Git client

## Offline Requirements

### Why Offline?

- Corporate security policy
- No server infrastructure
- No external network dependencies
- Portable across workstations

### How It's Achieved

**Web UI**:
- All libraries bundled in `/AutoBericht/libs/`
- Runs from file:// or local file server
- No CDN, no external resources
- File System Access API for local file handling

**VBA**:
- Self-contained within Excel workbook
- Optional: PowerShell sync requires one-time network access
- Macros persist in `.xlsm` file

**Data**:
- All data local (JSON, Excel, photos)
- No cloud storage required
- No API calls during operation

## Deployment Model

### For Report Creators

```
1. Install Excel with VBA enabled
2. Copy AutoBericht/ folder (including legacy) to workstation
3. Open project.xlsm
4. Import macros once (optional: PowerShell sync)
5. Ready to work
```

### For Developers

```
1. Clone repository
2. Import VBA modules for Excel development
3. Run the local server (recommended) and open the AutoBericht UI
```

## Extension Points

### Adding New VBA Modules

1. Create `.bas` file in `/macros/`
2. Follow naming convention: `modAB[Purpose]`
3. Use `modABConstants` for shared definitions
4. Import via `modMacroSync.RefreshProjectMacros`

### Adding Web UI Features

1. Add ES module in `AutoBericht/legacy/js/`
2. Keep modular, avoid global state
3. Use existing state management
4. Update tab navigation if needed

### Adding Export Formats

1. Web UI: Add generator in export pipeline
2. VBA: Add export function in new module
3. Update `project.json` schema if needed
4. Document in guides

## Performance Considerations

### VBA

- **Import**: ~1-5 seconds for 100-300 finding rows
- **Export**: ~2-10 seconds for full project
- **PhotoSorter (Excel)**: Handles 700+ photos with lazy loading

### Web UI

- **Load**: <2 seconds for typical project
- **PDF Export**: 5-15 seconds depending on content
- **PPTX Export**: 3-10 seconds per deck

## Security Model

### Data Isolation

- All data local to workstation
- No external network calls during operation
- No telemetry or analytics

### VBA Security

- Macros require "Trust access to VBA project object model"
- Code signing not currently implemented
- Source control via GitHub (audit trail)

### Web UI Security

- Content Security Policy for XSS protection
- DOMPurify for Markdown sanitization
- No eval(), no dynamic code execution
- All libraries vetted and bundled

## Troubleshooting

### Common Integration Issues

**"Project.json won't load in web UI"**
→ Check JSON schema version, validate with AJV

**"VBA import fails"**
→ Verify Word/Excel source format matches expected headers

**"Excel file won't open in web UI"**
→ Ensure SheetJS library loaded, check browser console

**"Photos not displaying"**
→ Verify file paths are relative, check browser file access permissions

## Next Steps

- **Understand VBA**: [VBA Architecture](vba-architecture.md)
- **Understand Data**: [Data Model](data-model.md)
- **Use the System**: [Getting Started Guide](../guides/getting-started.md)
- **Develop VBA**: [VBA Workflow](../guides/vba-workflow.md)
- **Develop UI**: [AutoBericht README](../../AutoBericht/README.md)

---

**Related Documentation**:
- [VBA Architecture](vba-architecture.md) - Macro system design
- [Data Model](data-model.md) - JSON and Excel structure
- [VBA Workflow Guide](../guides/vba-workflow.md) - Excel usage
- [HTML Workflow Guide](../guides/html-workflow.md) - Web UI usage
