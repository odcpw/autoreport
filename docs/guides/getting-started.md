# Getting Started with AutoBericht

This guide walks you through creating your first safety assessment report from start to finish.

## Prerequisites

### Required
- Windows PC with Excel 2016 or later
- Modern web browser (Chrome or Edge recommended)
- VBA macros enabled in Excel

### Your Materials
- Master findings document (Word .docx)
- Customer self-assessment (Excel file)
- Project photos in a folder

## Step 1: Set Up the Excel Workbook

### First Time Setup

1. **Open the workbook**
   ```
   Open: project.xlsm
   Enable macros when prompted
   ```

2. **Sync VBA modules** (optional, if using GitHub)
   - Run the PowerShell sync script (see [PowerShell Sync](../reference/powershell-sync.md))
   - Or manually import `.bas` files from `/macros/` folder

3. **Initialize structured sheets**
   ```vb
   ' In Excel VBA editor (Alt+F11), run:
   Sub SetupProject()
       EnsureAutoBerichtSheets clearExisting:=True
   End Sub
   ```

   This creates the structured sheets:
   - Meta
   - Chapters
   - Rows
   - Photos
   - Lists
   - ExportLog
   - OverridesHistory

## Step 2: Import Master Findings

The master findings are the template content for all assessment criteria.

1. **Prepare Word document**
   - Format: Table with headers
   - Required columns: `ReportItemID`, `Feststellung`, `Level1`, `Level2`, `Level3`, `Level4`
   - Example ID format: `1.1.1`, `1.1.2`, etc.

2. **Run the import macro**
   ```
   In Excel: Run ImportMyMaster
   Select your Word document
   Wait for import to complete
   ```

3. **Verify results**
   - Check `Rows` sheet for imported content
   - Review summary dialog (inserted vs updated counts)
   - All rows should have `masterFinding` and `masterLevel1-4` populated

**See**: [VBA Workflow - Master Import](vba-workflow.md#21-master-findings-word--rows)

## Step 3: Import Customer Self-Assessment

Customer self-assessment provides their answers to each criterion.

1. **Prepare Excel file**
   - Required columns: `ReportItemID`, answer column (D), remarks (E), priority (F)
   - IDs must match master findings

2. **Run the import macro**
   ```
   In Excel: Run ImportSelbstbeurteilungKunde
   Select customer Excel file
   Wait for import to complete
   ```

3. **Verify results**
   - Check `Rows` sheet for customer data
   - `customerAnswer`, `customerRemark`, `customerPriority` should be filled
   - Review summary dialog

**See**: [VBA Workflow - Customer Import](vba-workflow.md#22-customer-self-assessment-excel--rows)

## Step 4: Export to JSON

Export your data to the unified format that the web UI can read.

1. **Run export macro**
   ```
   In Excel: Run modABProjectExport.ExportProjectJson
   Choose output location (e.g., C:\Projects\myproject.json)
   ```

2. **Verify export**
   - Open the JSON file in a text editor
   - Should contain: version, meta, chapters, rows, photos, lists, history
   - Validate structure matches [Data Model](../architecture/data-model.md)

**See**: [VBA Workflow - JSON Export](vba-workflow.md#4-json-export)

## Step 5: Open Web UI

Now you can use the offline web interface for content editing.

1. **Open AutoBericht**
   ```
   Navigate to: AutoBericht/index.html

   Option A: Double-click (may require browser flag)
   Option B: Right-click → Open with → Chrome/Edge
   Option C: Use local server: python -m http.server
   ```

2. **Load your project**
   ```
   Go to Settings tab
   Click "Upload project.json"
   Select the JSON file you exported
   ```

3. **Verify loading**
   - Settings tab shows green validation
   - AutoBericht tab shows chapter tree
   - PhotoSorter tab shows (no photos yet)

**See**: [HTML Workflow Guide](html-workflow.md)

## Step 6: Add and Tag Photos

Organize project photos by chapter, category, and training topic.

### Option A: Using Excel PhotoSorter

1. **Scan photos into Excel**
   ```
   Open PhotoSorterForm in Excel
   Click "Scan Directory"
   Select your photos folder
   ```

2. **Tag photos**
   - Click chapter buttons to assign photos
   - Tag by category (Gefährdung, Brandschutz, etc.)
   - Tag by training topic
   - Photos sheet gets updated

3. **Export JSON again**
   ```
   Run modABProjectExport.ExportProjectJson
   ```

4. **Reload in web UI**
   ```
   Go to Settings tab
   Upload the new project.json
   ```

### Option B: Using Web UI PhotoSorter

1. **Upload photos**
   ```
   Go to PhotoSorter tab
   Click "Add Photos"
   Select photos or folder
   ```

2. **Tag photos**
   - Select photo thumbnail
   - Toggle chapter buttons
   - Add category tags
   - Add training tags
   - Add notes in textarea

3. **Auto-save**
   - Changes saved to in-memory project state
   - Export JSON when done

**See**: [HTML Workflow - PhotoSorter](html-workflow.md#photosorter-tab)

## Step 7: Edit Findings and Recommendations

Review and customize the report content.

1. **Navigate the chapter tree**
   ```
   AutoBericht tab → Left sidebar
   Click chapters to expand
   Click findings to edit
   ```

2. **For each finding:**
   - Review master finding text
   - Check customer answer/remark
   - Select appropriate recommendation level (1-4)
   - Add overrides if needed:
     - ☑️ "Use adjusted finding" → edit finding text
     - ☑️ "Use adjusted recommendation" → edit recommendation text
   - Toggle inclusion:
     - ☑️ "Include finding"
     - ☑️ "Include recommendation"
   - Mark done when complete

3. **Use Markdown editors**
   - Left pane: edit text (Markdown syntax)
   - Right pane: live preview
   - Support for bullets, bold, italic, links

4. **Save your work**
   ```
   Settings tab → Export → Download project.json
   ```

**See**: [HTML Workflow - AutoBericht Tab](html-workflow.md#autobericht-tab)

## Step 8: Configure Export Settings

Prepare for PDF and PowerPoint generation.

1. **Go to Settings tab**

2. **Set export options**
   - Company name
   - Report locale (de-CH, fr-CH, etc.)
   - PDF margins (mm)
   - Header/footer preferences
   - Logo images (optional)

3. **Validate project**
   - Check for validation errors
   - Fix any issues shown
   - Ensure "Export Ready" is green

**See**: [HTML Workflow - Settings Tab](html-workflow.md#settings-tab)

## Step 9: Generate Reports

Create the final deliverables.

### PDF Report

```
AutoBericht → Settings tab → Export section
Click "Generate PDF Report"
Browser print dialog opens
Choose "Save as PDF"
```

### PowerPoint Report

```
Click "Export Report PPTX"
File downloads automatically
One slide per finding with photos
```

### Training Deck

```
Click "Export Training PPTX"
File downloads automatically
Photos grouped by training tag
```

### Project Snapshot

```
Click "Download project.json"
Saves current state for later editing
```

**See**: [HTML Workflow - Exports](html-workflow.md#export-overview)

## Step 10: Finalize and Distribute

1. **Review outputs**
   - Open PDF and verify formatting
   - Open PPTX files and check content
   - Ensure photos display correctly

2. **Make adjustments**
   - Return to AutoBericht tab to edit
   - Re-export after changes
   - Iterate until satisfied

3. **Archive project**
   - Save final `project.json`
   - Keep Excel workbook
   - Organize photo files
   - Document any custom overrides

## Common Workflows

### Quick Report Update

```
1. Open AutoBericht/index.html
2. Load existing project.json
3. Make edits in AutoBericht tab
4. Export PDF (Ctrl+P)
5. Done!
```

### Adding New Photos

```
1. Add photos to folder
2. Excel: PhotoSorter → Scan Directory
3. Excel: Tag new photos
4. Excel: Export JSON
5. Web UI: Reload project.json
```

### Iterative Editing

```
while (not satisfied):
    Edit content in web UI
    Generate PDF preview
    Review with team
    Make adjustments
Save final project.json
```

## Keyboard Shortcuts

### Web UI
- `Alt+1` / `Alt+2` / `Alt+3`: Switch tabs
- `Ctrl+/`: Show/hide keyboard shortcuts
- `Esc`: Close modals
- `Tab` / `Shift+Tab`: Navigate form fields

### Excel VBA
- `Alt+F11`: Open VBA editor
- `F5`: Run macro
- `Ctrl+Break`: Stop running macro

## Troubleshooting

### "Import failed - headers not found"
→ Verify Word document has correct table headers (see Step 2)

### "JSON export is empty"
→ Ensure structured sheets are populated (check Rows sheet)

### "Web UI won't load JSON"
→ Validate JSON structure, check browser console for errors

### "Photos not displaying"
→ Verify file paths are relative, photos in same folder structure

### "PDF export looks wrong"
→ Check Settings tab for margin/header configuration

### "Can't edit in web UI"
→ Ensure project is loaded (Settings tab should show validation status)

## Next Steps

- **Learn VBA workflow**: [VBA Workflow Guide](vba-workflow.md)
- **Master web UI**: [HTML Workflow Guide](html-workflow.md)
- **Understand data model**: [Data Model](../architecture/data-model.md)
- **System architecture**: [System Overview](../architecture/system-overview.md)

## Getting Help

1. Check this guide for your specific task
2. Consult [Documentation Index](../INDEX.md)
3. Review error messages carefully
4. Ask your team for support

---

**Congratulations!** You've created your first AutoBericht report. The more you use it, the faster and easier it becomes.
