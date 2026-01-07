# AutoBericht Round-Trip Smoke Checklist (10 minutes)

Goal: verify the **VBA-exported unified `project.json`** round-trips through the web UI and back into Excel without manual fixes.

## Prereqs

- Excel workbook `project.xlsm` with current macros imported/enabled
- A small sample data set:
  - master findings Word file
  - customer self-assessment Excel file
  - 2–5 photos in a local folder
- AutoBericht UI runnable via `AutoBericht/start-autobericht.cmd` (recommended)

## Steps

1. In Excel: run the normal import workflow (master + customer) so the structured sheets are populated.
2. In Excel PhotoSorter: scan/tag the 2–5 photos and confirm tags are written to `Photos` / `PhotoTags` / `Lists` (as applicable).
3. In Excel: export `project.json` using the VBA export macro.
4. In AutoBericht (web UI):
   - Import the exported `project.json`
   - Confirm validation shows “ready” (exports enabled)
   - Open the AutoBericht tab and confirm you can see rows and edit text
5. Make a minimal edit:
   - Toggle one finding to excluded
   - Change selected level for one row
   - Add a small finding override text
6. In AutoBericht: export/download the updated `project.json`.
7. In Excel: import/load the updated `project.json` via the VBA loader macro.
8. Verify in Excel:
   - The edited row reflects updated workstate (level/override/include)
   - Photo tags and lists are still present
9. (Optional) Re-export `project.json` from Excel and compare with AutoBericht export:
   - Schema-valid
   - No missing required sections (`meta`, `lists`, `photos`, `chapters[].rows[]`)

## Expected Outcome

- AutoBericht accepts the VBA output directly (no conversion scripts).
- AutoBericht writes the same unified contract back out.
- Excel loader accepts the AutoBericht output and preserves edits.
