# AutoBericht Web UI Workflow Guide

Complete guide to using the offline HTML/JavaScript interface for photo management, content editing, and report generation.

## Overview

The AutoBericht web UI is a standalone HTML application that runs entirely offline in your browser. No server, no internet connection, no installation required.

**Three Main Tabs**:
1. **PhotoSorter**: Tag and organize photos
2. **AutoBericht**: Edit findings and recommendations
3. **Settings**: Configure exports and manage project

## Getting Started

### Opening the Application

**Option A: Direct file open** (requires browser flag)
```bash
# Windows
Right-click AutoBericht/index.html → Open with → Chrome/Edge

# Or launch Chrome with:
chrome.exe --allow-file-access-from-files path\to\index.html
```

**Option B: Local server** (recommended for development)
```bash
cd AutoBericht
python -m http.server 8000

# Open: http://localhost:8000
```

### First Launch

1. You'll see three tabs: PhotoSorter, AutoBericht, Settings
2. Go to **Settings** tab first
3. Upload your `project.json` file
4. Wait for validation to complete (green status)
5. Navigate to other tabs

**See**: [Getting Started Guide](getting-started.md) for project creation

## PhotoSorter Tab

Tag and organize photos by chapter, category, and training topic.

### Interface Layout

```
┌────────────────────────────────────────────────────┐
│  PhotoSorter                                        │
├────────────────────────────────────────────────────┤
│                                                     │
│  [Add Photos] [Scan Folder]                       │
│                                                     │
│  ┌──────────────────────────────────┐             │
│  │  Photo Thumbnail Grid            │             │
│  │  (Virtual Scrolling)             │             │
│  │                                   │             │
│  │  [IMG] [IMG] [IMG] [IMG] [IMG]   │             │
│  │  [IMG] [IMG] [IMG] [IMG] [IMG]   │             │
│  │  ...                              │             │
│  └──────────────────────────────────┘             │
│                                                     │
│  Selected Photo:                                   │
│  ┌──────────────────────────────────┐             │
│  │  [Large Preview]                  │             │
│  └──────────────────────────────────┘             │
│                                                     │
│  Notes: [___________________________]              │
│                                                     │
│  Chapter Tags:                                     │
│  [1.1] [1.2] [1.3] [4.8] ...                      │
│                                                     │
│  Categories:                                       │
│  [Gefährdung] [Brandschutz] ...                   │
│                                                     │
│  Training:                                         │
│  [PSA Basics] [Führungskräfte] ...                │
│                                                     │
└────────────────────────────────────────────────────┘
```

### Adding Photos

**Method 1: Upload individual photos**
```
1. Click "Add Photos"
2. Select one or more image files
3. Photos appear in grid
```

**Method 2: Scan folder** (Chrome/Edge only)
```
1. Click "Scan Folder"
2. Choose directory (supports nested folders)
3. All images imported automatically
```

**Supported formats**: JPG, JPEG, PNG, GIF, WEBP

### Tagging Photos

1. **Select photo**: Click thumbnail in grid
2. **View full size**: Large preview appears below
3. **Add notes**: Type in notes field (optional)
4. **Tag by chapter**: Click chapter button(s) - multiple allowed
5. **Tag by category**: Click category button(s) - multiple allowed
6. **Tag by training**: Click training button(s) - multiple allowed

**Visual feedback**:
- Active tags are highlighted
- Button shows count (e.g., "1.1 (5)" = 5 photos tagged)
- Changes save automatically to project state

### Navigation

**Keyboard**:
- `Arrow Keys`: Move between thumbnails
- `Space`: Toggle selected photo's first tag
- `Enter`: Focus notes field
- `Tab`: Navigate tag buttons

**Mouse**:
- Click thumbnail to select
- Click buttons to toggle tags
- Scroll grid to browse more photos

### Photo Metadata

Each photo stores:
```json
{
  "fileName": "DSC00123.jpg",
  "displayName": "Leitbild im Empfang",
  "notes": "Neue Tafel, gut sichtbar",
  "tags": {
    "chapters": ["1.1.3"],
    "categories": ["Gefährdung"],
    "training": ["Führungskräftetraining"]
  },
  "preferredLocale": "de-CH",
  "capturedAt": "2025-02-11T10:30:00Z"
}
```

### Bulk Operations

**Tag multiple photos at once**:
```
1. Select first photo
2. Hold Shift, click last photo (range select)
3. OR: Hold Ctrl, click individual photos (multi-select)
4. Click tag button → applies to all selected
```

**Remove tags**:
```
1. Select photo(s)
2. Click active tag button to toggle off
3. OR: Clear all tags: Right-click → Clear All
```

### Export Integration

Tagged photos are automatically included in:
- **Report PPTX**: Photos appear with their tagged chapter findings
- **Training PPTX**: Photos grouped by training tag
- **PDF Report**: Photos embedded in chapter sections (future)

**See**: [Export Formats](#export-formats) below

## AutoBericht Tab

Edit findings, recommendations, and work state for each report item.

### Interface Layout

```
┌──────────────────────────────────────────────────────────┐
│  AutoBericht                                              │
├──────────────────────────────────────────────────────────┤
│                                                           │
│  ┌────────────┐  ┌──────────────────────────────────┐   │
│  │  Chapter   │  │  Finding Editor                   │   │
│  │  Tree      │  │                                    │   │
│  │            │  │  Finding: 1.1.1                   │   │
│  │  ▼ 1       │  │  Unternehmensleitbild...          │   │
│  │    ▼ 1.1   │  │                                    │   │
│  │      1.1.1 │  │  Master Finding:                  │   │
│  │      1.1.2 │  │  [_____________________________]  │   │
│  │      1.1.3 │  │                                    │   │
│  │    ▶ 1.2   │  │  ☐ Use adjusted finding           │   │
│  │  ▶ 2       │  │  [_____________________________]  │   │
│  │  ▶ 3       │  │                                    │   │
│  │  ▶ 4       │  │  Level: (•) 1 ( ) 2 ( ) 3 ( ) 4  │   │
│  │    ▶ 4.8   │  │                                    │   │
│  │            │  │  Recommendation:                  │   │
│  └────────────┘  │  [_____________________________]  │   │
│                  │                                    │   │
│                  │  ☐ Use adjusted recommendation     │   │
│                  │  [_____________________________]  │   │
│                  │                                    │   │
│                  │  Include:                          │   │
│                  │  ☑ Finding  ☑ Recommendation      │   │
│                  │                                    │   │
│                  │  Notes: [_______________]          │   │
│                  │  ☐ Done                            │   │
│                  └──────────────────────────────────────┘
└──────────────────────────────────────────────────────────┘
```

### Chapter Tree Navigation

**Expand/collapse chapters**:
- Click ▶/▼ icon or chapter name
- Shows hierarchical structure (1 → 1.1 → 1.1.1)

**Select finding**:
- Click finding ID (e.g., "1.1.1")
- Editor loads finding details
- Previous changes auto-saved

**Visual indicators**:
- ✓ Green: Finding marked done
- ⚠ Orange: Has overrides
- ○ Gray: Not started
- Numbers show completion (5/10)

### Editing a Finding

**1. Review master content**
```
Master Finding: Template text from Word import
Master Recommendations (Level 1-4): Template recommendation for each level
Customer Answer: Their self-assessment score (0-4)
Customer Remark: Their notes
Customer Priority: Optional priority flag
```

**2. Select recommendation level**
```
Choose level (1-4) based on customer answer and your assessment
This determines which recommendation text to include
```

**3. Add overrides (optional)**
```
☐ Use adjusted finding
   → Enables Markdown editor for custom finding text
   → Replaces master finding in export

☐ Use adjusted recommendation
   → Enables Markdown editor for custom recommendation
   → Replaces master level text in export
```

**4. Set inclusion flags**
```
☑ Include finding
   → Finding appears in report

☑ Include recommendation
   → Recommendation appears in report

Uncheck to exclude from exports
```

**5. Add notes and mark done**
```
Notes: Internal notes (not exported)
☐ Done: Mark complete (visual indicator in tree)
```

### Markdown Editors

**Features**:
- Live preview pane (side-by-side)
- Syntax highlighting (CodeMirror)
- Markdown support: **bold**, *italic*, lists, links
- Undo/redo (Ctrl+Z / Ctrl+Y)
- Auto-save on blur

**Common Markdown**:
```markdown
**Bold text**
*Italic text*
- Bullet point
1. Numbered list
[Link text](url)
```

**Preview**:
- Right pane shows rendered output
- DOMPurify sanitizes HTML
- Updates as you type

### Chapter 4.8 (Audit Observations)

Special chapter for custom findings not in master:

**Adding custom rows**:
```
1. Navigate to chapter 4.8
2. Click "+ Add Observation"
3. New row created with ID "4.8.custom-001"
4. Edit finding and recommendation from scratch
5. Use overrides for all content
```

**Renumbering**:
- At export, 4.8.custom-* becomes 4.8.1, 4.8.2, etc.
- Original IDs preserved in `exportHints.renumberedId`

### Bulk Operations

**Mark multiple as done**:
```
Select chapter in tree → Right-click → Mark all done
```

**Clear overrides**:
```
Select finding → Right-click → Reset to master
```

**Copy finding**:
```
Select finding → Right-click → Duplicate (chapter 4.8 only)
```

### Keyboard Shortcuts

- `↑` / `↓`: Navigate findings in tree
- `Enter`: Expand/collapse chapter
- `Tab`: Navigate form fields
- `Ctrl+S`: Manual save (auto-saves already)
- `Ctrl+/`: Show keyboard shortcuts overlay

## Settings Tab

Configure exports, validate project, and manage data.

### Interface Layout

```
┌────────────────────────────────────────────────────┐
│  Settings                                           │
├────────────────────────────────────────────────────┤
│                                                     │
│  Project Management                                │
│  [Upload project.json] [Upload .xlsm]             │
│  [Download project.json]                           │
│                                                     │
│  ────────────────────────────────────────────────  │
│                                                     │
│  Export Configuration                               │
│  Company: [___________]                            │
│  Locale: [de-CH ▼]                                 │
│                                                     │
│  PDF Settings:                                     │
│  Margins (mm): Top [25] Bottom [25]                │
│                Left [25] Right [25]                │
│  ☑ Show header/footer                              │
│  ☑ Show logos                                      │
│                                                     │
│  ────────────────────────────────────────────────  │
│                                                     │
│  Validation Status: ✓ Export Ready                │
│  Issues: None                                      │
│                                                     │
│  ────────────────────────────────────────────────  │
│                                                     │
│  Export                                            │
│  [Generate PDF Report]                             │
│  [Export Report PPTX]                              │
│  [Export Training PPTX]                            │
│  [Download project.json]                           │
│                                                     │
└────────────────────────────────────────────────────┘
```

### Loading Projects

**Upload project.json**:
```
1. Click "Upload project.json"
2. Select JSON file
3. Wait for parsing and validation
4. Check validation status
```

**Upload .xlsm** (direct Excel file):
```
1. Click "Upload .xlsm"
2. Select Excel workbook
3. SheetJS parses structured sheets
4. Converts to in-memory project state
```

**Validation**:
- ✓ Green: Ready to export
- ⚠ Orange: Warnings (can still export)
- ✗ Red: Errors (fix before export)

### Export Configuration

**Project metadata**:
- **Company**: Appears in headers/footers
- **Locale**: Date/time formatting (de-CH, fr-CH, it-CH, en-CH)
- **Author**: Your name (optional)

**PDF settings**:
- **Margins**: Top, bottom, left, right (millimeters)
- **Header/Footer**: Toggle visibility
- **Logos**: Toggle logo display on first page
- **Footer text**: Template with tokens `{company}`, `{date}`, `{time}`

**PPTX settings**:
- **Slide size**: 16:9 (default) or 4:3
- **Template**: Choose color scheme
- **Image quality**: High, medium, low (affects file size)

### Validation

**Automatic checks**:
- Required fields present (meta.projectId, meta.company)
- Chapter IDs valid and unique
- Row IDs match pattern (e.g., "1.1.1")
- Photo tags reference existing chapters
- No orphaned rows (chapter exists)

**Issue types**:
- **Error**: Must fix before export
- **Warning**: Can proceed, but review recommended
- **Info**: Suggestions for improvement

**Fixing issues**:
```
1. Click issue in validation panel
2. Jumps to relevant tab (PhotoSorter or AutoBericht)
3. Fix problem
4. Return to Settings → revalidates automatically
```

## Export Formats

### PDF Report

**Generation**:
```
Settings → Export → Generate PDF Report
Browser print dialog opens
Choose "Save as PDF" or physical printer
```

**Content**:
- Cover page with logos and company name
- Table of contents (auto-generated)
- Chapters with findings and recommendations
- Photos embedded inline (if enabled)
- Page numbers and headers/footers

**Technology**: Paged.js (CSS Paged Media polyfill)

**Tips**:
- Preview before printing (browser print preview)
- Adjust margins in Settings if content is cut off
- Use landscape for wide tables (future feature)

### Report PPTX

**Generation**:
```
Settings → Export → Export Report PPTX
File downloads automatically: report.pptx
```

**Content**:
- Title slide with project info
- One slide per included finding
- Finding text as title
- Recommendation as bullet points (Markdown converted)
- Photos tagged with chapter ID embedded
- Master slide with branding

**Layout**:
```
┌─────────────────────────────────┐
│  Finding: 1.1.1                 │
│  Unternehmensleitbild           │
├─────────────────────────────────┤
│                                  │
│  • Recommendation Level 1:      │
│    Das Unternehmen sollte...   │
│                                  │
│  [Photo 1]  [Photo 2]           │
│                                  │
└─────────────────────────────────┘
```

**Technology**: PptxGenJS

### Training PPTX

**Generation**:
```
Settings → Export → Export Training PPTX
File downloads automatically: training.pptx
```

**Content**:
- Title slide per training tag
- Photos grouped by training assignment
- References to chapters where photos appear
- Notes field content as speaker notes

**Layout**:
```
┌─────────────────────────────────┐
│  Training: PSA Basics           │
├─────────────────────────────────┤
│                                  │
│  [Photo 1]  [Photo 2]  [Photo 3]│
│  [Photo 4]  [Photo 5]  [Photo 6]│
│                                  │
│  References:                     │
│  • Chapter 1.1.3                │
│  • Chapter 2.4.1                │
│                                  │
└─────────────────────────────────┘
```

### Project JSON

**Generation**:
```
Settings → Export → Download project.json
File downloads: project-YYYY-MM-DD.json
```

**Purpose**:
- Save work in progress
- Share with colleagues
- Version control
- Re-import for further editing

**Content**:
- Complete project state
- All findings, overrides, photos, tags
- Export configuration
- Override history

## Advanced Features

### Keyboard Navigation

**Global**:
- `Alt+1`: PhotoSorter tab
- `Alt+2`: AutoBericht tab
- `Alt+3`: Settings tab
- `Ctrl+/`: Show keyboard shortcuts overlay
- `Esc`: Close modals/overlays

**PhotoSorter**:
- `Arrow keys`: Navigate photos
- `Space`: Toggle first tag
- `Delete`: Remove selected photo

**AutoBericht**:
- `↑` / `↓`: Navigate findings
- `Enter`: Expand/collapse chapter
- `Ctrl+E`: Focus editor
- `Ctrl+S`: Save (auto-saves on blur anyway)

### Browser Compatibility

**Fully Supported**:
- Chrome 90+
- Edge 90+
- Firefox 88+ (file access limited)

**Partially Supported**:
- Safari 14+ (no File System Access API)
- Mobile browsers (limited file access)

**Required Features**:
- ES modules
- File API
- IndexedDB (for offline storage)
- Web Workers (for background processing)

### Offline Operation

**What works offline**:
- All editing features
- Photo tagging
- Report preview
- PDF generation (via browser print)
- PPTX generation
- JSON export

**What requires internet** (optional):
- Initial library download (if using CDN)
- Vision critique tool (uses AI APIs)
- GitHub macro sync

**Storage**:
- Project state stored in memory during session
- Use "Download project.json" to persist
- Future: IndexedDB auto-save

### Performance

**Large projects** (700+ photos, 300+ findings):
- Virtual scrolling for photo grid (smooth with 1000+ photos)
- Lazy loading for thumbnails
- Chapter tree optimized for deep hierarchies
- Validation runs in background Web Worker

**Export times**:
- PDF: 5-15 seconds (depends on content and photos)
- PPTX: 3-10 seconds per deck
- JSON: < 1 second

## Troubleshooting

### "Can't open index.html"
→ Use local server or launch browser with `--allow-file-access-from-files`

### "Project won't load"
→ Check JSON structure, validate schema, check browser console

### "Photos not displaying"
→ Ensure file paths relative, photos in accessible directory

### "PDF export looks wrong"
→ Adjust margins in Settings, check browser print preview settings

### "PPTX export fails"
→ Check for unsupported Markdown syntax, verify photos accessible

### "Changes not saving"
→ Auto-save on blur; manually use "Download project.json" to persist

### "Slow performance"
→ Close other tabs, reduce photo count, use Chrome/Edge

## Best Practices

### Workflow Tips

1. **Save frequently**: Export JSON after major changes
2. **Validate often**: Check Settings tab before big edits
3. **Tag photos early**: Easier to organize before hundreds accumulated
4. **Use overrides sparingly**: Master content is there for a reason
5. **Mark done**: Helps track progress across large projects

### Content Guidelines

1. **Findings**: Clear, objective, factual
2. **Recommendations**: Actionable, specific, leveled appropriately
3. **Notes**: Internal only, use for team communication
4. **Overrides**: Document why you deviated from master

### Performance

1. **Photo sizes**: Keep under 5MB each (optimize before upload)
2. **Project size**: Split very large projects (>500 findings)
3. **Browser**: Close other tabs while editing
4. **Export**: Test small project first, ensure settings correct

## Next Steps

- **Create your first report**: [Getting Started Guide](getting-started.md)
- **Learn VBA workflow**: [VBA Workflow Guide](vba-workflow.md)
- **Understand data model**: [Data Model](../architecture/data-model.md)
- **System architecture**: [System Overview](../architecture/system-overview.md)

## Related Documentation

- [Getting Started](getting-started.md) - End-to-end first report
- [VBA Workflow](vba-workflow.md) - Excel macro usage
- [Data Model](../architecture/data-model.md) - JSON and Excel structure
- [AutoBericht README](../../AutoBericht/README.md) - Technical details
- [Vision Critique](../../vision-critique/README.md) - UI quality analysis

---

**Need help?** Check the [Documentation Index](../INDEX.md) or consult your team.
