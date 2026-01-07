param(
  [int]$Port = 0,
  [switch]$NoOpen,
  [string]$StartPath = 'mini/'
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$server = Join-Path $PSScriptRoot 'tools\serve-autobericht.ps1'

& $server -Root $root -Port $Port -NoOpen:$NoOpen -StartPath $StartPath
