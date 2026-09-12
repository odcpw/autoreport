param(
  [int]$Port = 0,
  [switch]$NoOpen,
  [string]$StartPath = 'mini/'
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$server = Join-Path $PSScriptRoot 'tools\serve-autobericht.ps1'
$templates = Join-Path (Split-Path -Parent $PSScriptRoot) 'templates'

& $server -Root $root -TemplateRoot $templates -Port $Port -NoOpen:$NoOpen -StartPath $StartPath
