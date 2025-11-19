# PowerShell Sync Script (Workstation Setup)

## Overview
Use this script on the work PC to pull the full repo from GitHub (*odcpw/autoreport*) into a local folder (ideally outside OneDrive/SharePoint). Afterwards, run the `RefreshProjectMacros` macro from Excel to re-import everything automatically.

## Requirements
- PowerShell 5.1+ with internet access.
- Git installed and available in `PATH`.
- GitHub access to `https://github.com/odcpw/autoreport.git`.
- Excel workbook is configured with `modMacroSync.RefreshProjectMacros` and “Trust access to the VBA project object model” enabled.

## Script (already in repo as `sync-autobericht.ps1`)
```powershell
# Recommended: run from the folder where you want the repo (e.g., C:\Autobericht)
PS C:\Autobericht> .\sync-autobericht.ps1

# Optional overrides
# PS C:\Autobericht> .\sync-autobericht.ps1 -TargetFolder "C:\Autobericht" -Branch main
# PS C:\Autobericht> .\sync-autobericht.ps1 -FullClone   # remove depth=1
```

This clones (or updates) the repo into `<TargetFolder>\autoreport`. It includes `AutoBericht`, `macros/`, `docs/`, and all supporting assets.

## Typical Workflow
1. Run the PowerShell script from a local path (preferably outside OneDrive/SharePoint, e.g., `C:\Autobericht`):
   ```powershell
   PS C:\Autobericht> .\sync-autobericht.ps1
   ```
   This clones/updates the repo into `C:\Autobericht\autoreport`.
2. Open `project.xlsm` and run `RefreshProjectMacros`. The loader removes previous modules and imports everything fresh from the current folder.
3. For the web app, serve `AutoBericht` locally (e.g., `python -m http.server 5501`) and open `http://localhost:5501/AutoBericht/index.html`.

## Notes
- The script uses `git clone`/`git pull` to keep the full repo in sync. Use the `-FullClone` switch if you prefer a non-shallow clone.
- If you must avoid Git, you can still download `https://github.com/odcpw/autoreport/archive/refs/heads/main.zip` manually, but Git is more reliable on OneDrive/SharePoint.
