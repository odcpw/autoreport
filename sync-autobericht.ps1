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

    Write-Host "Sync complete. Open AutoBericht via http://localhost:5501/AutoBericht/index.html (serve locally)." -ForegroundColor Green
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
