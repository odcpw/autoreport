param(
    [string]$TargetFolder   = (Get-Location).Path,  # where to drop the repo contents
    [string]$RepoArchiveUrl = "https://github.com/odcpw/autoreport/archive/refs/heads/main.zip",
    [switch]$CleanTarget                       # optional: wipe target folder before copying
)

$ErrorActionPreference = 'Stop'

function Ensure-Folder([string]$Path) {
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Assert-FileExists([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Expected file missing during sync verification: $Path"
    }
}

function Clear-AndVerifyMarkOfTheWeb([string]$Path) {
    Assert-FileExists -Path $Path
    Unblock-File -LiteralPath $Path -ErrorAction Stop
    $streams = @(Get-Item -LiteralPath $Path -Stream * -ErrorAction Stop)
    if ($streams.Stream -contains 'Zone.Identifier') {
        throw "Mark-of-the-Web still present after Unblock-File: $Path. Remove the Zone.Identifier stream and rerun sync."
    }
}

function Clear-AndVerifyExtractedPayload([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Expected extracted payload missing during sync verification: $Path"
    }

    $markedFiles = @(Get-ChildItem -LiteralPath $Path -Recurse -File | Where-Object {
        $streams = @(Get-Item -LiteralPath $_.FullName -Stream * -ErrorAction Stop)
        $streams.Stream -contains 'Zone.Identifier'
    })

    foreach ($file in $markedFiles) {
        Clear-AndVerifyMarkOfTheWeb -Path $file.FullName
    }

    $remainingMarkedFiles = @(Get-ChildItem -LiteralPath $Path -Recurse -File | Where-Object {
        $streams = @(Get-Item -LiteralPath $_.FullName -Stream * -ErrorAction Stop)
        $streams.Stream -contains 'Zone.Identifier'
    })
    if ($remainingMarkedFiles.Count -gt 0) {
        throw "Extracted payload still contains Zone.Identifier streams: $($remainingMarkedFiles[0].FullName). Rerun sync after clearing the extracted files."
    }
}

$zipPath     = Join-Path $env:TEMP "autobericht_zip_download.zip"
$extractRoot = Join-Path $env:TEMP "autobericht_zip_extract"

try {
    Ensure-Folder -Path $TargetFolder
    $ResolvedTarget = (Resolve-Path $TargetFolder).Path

    Write-Host "Downloading archive..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $RepoArchiveUrl -OutFile $zipPath
    Write-Host "Clearing archive Internet marker..." -ForegroundColor Cyan
    Clear-AndVerifyMarkOfTheWeb -Path $zipPath

    if (Test-Path $extractRoot) {
        Remove-Item $extractRoot -Recurse -Force
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

    # Locate the extracted repo root (usually autoreport-main)
    $sourceInner = Join-Path $extractRoot "autoreport-main"
    if (-not (Test-Path $sourceInner)) {
        $sourceInner = (Get-ChildItem $extractRoot | Where-Object { $_.PSIsContainer }).FullName | Select-Object -First 1
    }
    if (-not $sourceInner -or -not (Test-Path $sourceInner)) {
        throw "Could not locate extracted repo root under $extractRoot"
    }

    Write-Host "Clearing extracted payload Internet markers..." -ForegroundColor Cyan
    Clear-AndVerifyExtractedPayload -Path $sourceInner

    if ($CleanTarget) {
        Write-Host "Cleaning target folder $ResolvedTarget..." -ForegroundColor Yellow
        Get-ChildItem -Path $ResolvedTarget -Force | Remove-Item -Recurse -Force
    }

    Write-Host "Copying repo into $ResolvedTarget ..." -ForegroundColor Cyan
    Copy-Item -Path (Join-Path $sourceInner '*') -Destination $ResolvedTarget -Recurse -Force

    foreach ($relativePath in @(
        'start-autobericht.cmd',
        'sync-autobericht.cmd',
        'sync-autobericht.ps1',
        'AutoBericht\start-autobericht.ps1',
        'AutoBericht\tools\serve-autobericht.ps1'
    )) {
        Clear-AndVerifyMarkOfTheWeb -Path (Join-Path $ResolvedTarget $relativePath)
    }

    Write-Host "Verified launcher files are marker-free." -ForegroundColor Green
    Write-Host "Sync complete. Start AutoBericht with start-autobericht.cmd." -ForegroundColor Green
}
catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
finally {
    foreach ($p in @($zipPath, $extractRoot)) {
        if (Test-Path $p) { Remove-Item $p -Recurse -Force }
    }
}
