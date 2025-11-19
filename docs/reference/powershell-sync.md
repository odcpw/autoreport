# PowerShell Sync Script (Workstation Setup)

## Overview
Use this script on the work PC to pull the full repo from GitHub (*odcpw/autoreport*) into a local folder (ideally outside OneDrive/SharePoint). Afterwards, run the `RefreshProjectMacros` macro from Excel to re-import everything automatically. This version **does not require Git**; it downloads the ZIP and copies the entire repo.

## Requirements
- PowerShell 5.1+ with internet access.
- GitHub access to download `https://github.com/odcpw/autoreport/archive/refs/heads/main.zip`.
- Excel workbook is configured with `modMacroSync.RefreshProjectMacros` and “Trust access to the VBA project object model” enabled.

## Script (already in repo as `sync-autobericht.ps1`)
```powershell
# Recommended: run from the folder where you want the repo contents (e.g., C:\Autobericht)
PS C:\Autobericht> .\sync-autobericht.ps1

# Optional overrides
# PS C:\Autobericht> .\sync-autobericht.ps1 -TargetFolder "C:\Autobericht"
# PS C:\Autobericht> .\sync-autobericht.ps1 -CleanTarget   # wipe target before copy
```

This downloads the ZIP and copies the full repo into `<TargetFolder>`. It includes `AutoBericht`, `macros/`, `docs/`, and all supporting assets.

## Typical Workflow
1. Run the PowerShell script from a local path (preferably outside OneDrive/SharePoint, e.g., `C:\Autobericht`):
   ```powershell
   PS C:\Autobericht> .\sync-autobericht.ps1
   ```
   This downloads and copies the repo into `C:\Autobericht` (overwriting existing files unless you use `-CleanTarget` to wipe first).
2. Open `project.xlsm` and run `RefreshProjectMacros`. The loader removes previous modules and imports everything fresh from the current folder.
3. For the web app, serve `AutoBericht` locally (e.g., `python -m http.server 5501`) and open `http://localhost:5501/AutoBericht/index.html`.

## Notes
- The script uses a ZIP download (no Git required) to keep the full repo in sync. Set `$TargetFolder` to a non-SharePoint path for reliability.
- If you prefer Git, replace the download section with `git pull` (not needed for the default workflow).
