# PowerShell Sync Script (Workstation Setup)

## Overview
Use this script on the work PC to pull the latest macros and docs from GitHub (*odcpw/autoreport*) and drop them into the local working directory. Afterwards, run the `RefreshProjectMacros` macro from Excel to re-import everything automatically.

## Requirements
- PowerShell 5.1+ with internet access.
- GitHub Personal Access Token (PAT) with `repo` scope stored in the credential manager or as environment variable (optional when using `Invoke-WebRequest` against public repo).
- Excel workbook is configured with `modMacroSync.RefreshProjectMacros` and “Trust access to the VBA project object model” enabled.

## Script (Save as `sync-autobericht.ps1`)
```powershell
param(
    [string]$TargetFolder = "$PSScriptRoot",
    [string]$RepoArchiveUrl = "https://github.com/odcpw/autoreport/archive/refs/heads/main.zip"
)

$zipPath = Join-Path $env:TEMP "autobericht_macros.zip"
$extractRoot = Join-Path $env:TEMP "autobericht_macros"

Write-Host "Downloading latest macros..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $RepoArchiveUrl -OutFile $zipPath

if (Test-Path $extractRoot) { Remove-Item $extractRoot -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

$sourceInner = Join-Path $extractRoot "autobericht-main"  # repo default folder name
if (-not (Test-Path $sourceInner)) {
    $sourceInner = (Get-ChildItem $extractRoot | Where-Object {$_.PSIsContainer}).FullName
}

$sourceMacros = Join-Path $sourceInner "macros"
$sourceDocs   = Join-Path $sourceInner "docs"

if (-not (Test-Path $sourceMacros)) {
    throw "Macros folder not found in extracted archive."
}

Write-Host "Syncing macros into" $TargetFolder -ForegroundColor Cyan
Copy-Item -Path $sourceMacros -Destination $TargetFolder -Recurse -Force
Copy-Item -Path $sourceDocs -Destination $TargetFolder -Recurse -Force

Write-Host "Sync complete. Run Excel macro 'RefreshProjectMacros' to import." -ForegroundColor Green
```

> Adjust `$RepoArchiveUrl` if using a different branch or tag.

## Typical Workflow
1. Run the PowerShell script from the local macro working directory (`.os	oolseferenceideli`):
   ```powershell
   PS C:\Autobericht\macros> .\sync-autobericht.ps1
   ```
2. Open `project.xlsm` and run `RefreshProjectMacros`. The loader removes previous modules and imports everything fresh from the current folder.
3. Proceed with PhotoSorter/AutoBericht editing as usual.

## Notes
- The script copies both `macros/` and `docs/` into the target folder for offline reference.
- If you prefer Git instead of `Invoke-WebRequest`, replace the download block with `git pull` inside the repo and keep the rest of the workflow unchanged.
