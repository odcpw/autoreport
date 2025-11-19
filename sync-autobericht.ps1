param(
    [string]$TargetFolder = (Get-Location).Path,
    [string]$RepoUrl      = "https://github.com/odcpw/autoreport.git",
    [string]$Branch       = "main",
    [switch]$FullClone
)

function Ensure-Git {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git is required but was not found in PATH."
    }
}

function Ensure-Folder([string]$Path) {
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

try {
    Ensure-Git

    Ensure-Folder -Path $TargetFolder
    $ResolvedTarget = (Resolve-Path $TargetFolder).Path

    $RepoName = [IO.Path]::GetFileNameWithoutExtension((Split-Path $RepoUrl -Leaf))
    if (-not $RepoName) { $RepoName = "autoreport" }

    $RepoPath = Join-Path $ResolvedTarget $RepoName
    $CloneArgs = @("--branch", $Branch, $RepoUrl, $RepoPath)
    if (-not $FullClone) {
        $CloneArgs = @("--depth", "1") + $CloneArgs
    }

    if (-not (Test-Path $RepoPath)) {
        Write-Host "Cloning $RepoUrl ($Branch) into $RepoPath..." -ForegroundColor Cyan
        git clone @CloneArgs
    }
    else {
        Write-Host "Updating existing repo at $RepoPath..." -ForegroundColor Cyan
        git -C "$RepoPath" fetch --all --prune
        git -C "$RepoPath" checkout "$Branch"
        git -C "$RepoPath" pull --ff-only origin "$Branch"
    }

    Write-Host "Sync complete. Repo available at:`n$RepoPath" -ForegroundColor Green
    Write-Host "Open AutoBericht via http://localhost:5501/AutoBericht/index.html (serve locally) or file:// if you start Edge with file access enabled."
}
catch {
    Write-Error $_
    exit 1
}
