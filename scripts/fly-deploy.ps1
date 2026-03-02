[CmdletBinding()]
param(
  [string]$App,
  [switch]$SkipPull,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Get-FlyAppFromToml {
  param([string]$TomlPath)

  if (-not (Test-Path -LiteralPath $TomlPath)) {
    throw "fly.toml not found at: $TomlPath"
  }

  $content = Get-Content -LiteralPath $TomlPath -Raw

  # Match lines like: app = 'pv-thermal-inspection-gi'
  $pattern = '(?m)^\s*app\s*=\s*["''](?<app>[^"''\r\n]+)["'']\s*$'
  $m = [regex]::Match($content, $pattern)
  if (-not $m.Success) {
    throw "Could not find app name in fly.toml (expected: app = '...')"
  }

  return $m.Groups['app'].Value
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
Set-Location -LiteralPath $repoRoot

if (-not $SkipPull) {
  if ($DryRun) {
    Write-Output "Would run: git pull"
  } else {
    git pull
  }
}

if (-not $App) {
  $App = Get-FlyAppFromToml -TomlPath (Join-Path $repoRoot 'fly.toml')
}

if ($DryRun) {
  Write-Output "Would run: fly deploy -a $App"
} else {
  fly deploy -a $App
}
