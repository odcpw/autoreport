param(
    [string]$TargetFolder   = (Get-Location).Path,  # where to drop the repo contents
    [string]$RepoArchiveUrl = "https://github.com/odcpw/autoreport/archive/refs/heads/main.zip",
    [switch]$CleanTarget                       # optional: wipe target folder before copying
)

function Ensure-Folder([string]$Path) {
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Unblock-InstalledScript([string]$BasePath, [string]$RelativePath) {
    $scriptPath = Join-Path $BasePath $RelativePath
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Expected installed script missing after sync: $scriptPath"
    }
    Unblock-File -LiteralPath $scriptPath -ErrorAction Stop
}

$zipPath     = Join-Path $env:TEMP "autobericht_zip_download.zip"
$extractRoot = Join-Path $env:TEMP "autobericht_zip_extract"

try {
    Ensure-Folder -Path $TargetFolder
    $ResolvedTarget = (Resolve-Path $TargetFolder).Path

    Write-Host "Downloading archive..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $RepoArchiveUrl -OutFile $zipPath

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

    if ($CleanTarget) {
        Write-Host "Cleaning target folder $ResolvedTarget..." -ForegroundColor Yellow
        Get-ChildItem -Path $ResolvedTarget -Force | Remove-Item -Recurse -Force
    }

    Write-Host "Copying repo into $ResolvedTarget ..." -ForegroundColor Cyan
    Copy-Item -Path (Join-Path $sourceInner '*') -Destination $ResolvedTarget -Recurse -Force

    foreach ($relativeScript in @(
        'sync-autobericht.ps1',
        'AutoBericht\start-autobericht.ps1',
        'AutoBericht\tools\serve-autobericht.ps1'
    )) {
        Unblock-InstalledScript -BasePath $ResolvedTarget -RelativePath $relativeScript
    }

    Write-Host "Sync complete. Start AutoBericht with start-autobericht.cmd." -ForegroundColor Green
}
catch {
    Write-Error $_
    exit 1
}
finally {
    foreach ($p in @($zipPath, $extractRoot)) {
        if (Test-Path $p) { Remove-Item $p -Recurse -Force }
    }
}
