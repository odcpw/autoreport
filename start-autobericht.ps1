param(
  [int]$Port = 0,
  [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'

$root = Join-Path $PSScriptRoot 'AutoBericht'
$server = Join-Path $PSScriptRoot 'tools\serve-autobericht.ps1'

& $server -Root $root -Port $Port -NoOpen:$NoOpen
