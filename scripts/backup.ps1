param(
    [string]$ProjectRoot = "$(Resolve-Path (Get-Location))",
    [string]$OutputDir = "backups",
    [switch]$IncludeData,
    [switch]$IncludeVendorSdk,
    [switch]$RepoOnly
)

$ErrorActionPreference = 'Stop'

$projectRootPath = (Resolve-Path $ProjectRoot).Path
$repoName = Split-Path -Leaf $projectRootPath
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$outputPath = Join-Path $projectRootPath $OutputDir
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

Write-Host "Project root: $projectRootPath"
Write-Host "Backup output: $outputPath"

# 1) Offline Git backup (includes full git history)
$bundlePath = Join-Path $outputPath "${repoName}_${timestamp}.bundle"
Write-Host "Creating git bundle: $bundlePath"
Push-Location $projectRootPath
try {
    git bundle create "$bundlePath" --all | Out-Null
} finally {
    Pop-Location
}

if ($RepoOnly) {
    Write-Host "RepoOnly specified; skipping full zip." 
    Write-Host "Done."
    exit 0
}

# 2) Full zip backup (optionally includes data/vendor SDK)
# We build a temporary staging folder that excludes heavy caches by default.
$staging = Join-Path $outputPath "_staging_${timestamp}"
$zipPath = Join-Path $outputPath "${repoName}_${timestamp}.zip"

New-Item -ItemType Directory -Force -Path $staging | Out-Null

$excludeDirs = @(
    '.git',
    '.venv',
    'venv',
    '__pycache__',
    '.pytest_cache',
    '.mypy_cache',
    '.ruff_cache',
    'node_modules',
    'dist',
    '.vite',
    'backups'
)

if (-not $IncludeData) {
    $excludeDirs += 'data'
    $excludeDirs += (Join-Path 'backend' 'thermal_out')
    $excludeDirs += (Join-Path 'backend' 'static' 'orthomosaic_sessions')
}

if (-not $IncludeVendorSdk) {
    $excludeDirs += (Join-Path 'backend' 'dji_thermal_sdk' 'doc')
    $excludeDirs += (Join-Path 'backend' 'dji_thermal_sdk' 'dataset')
    $excludeDirs += (Join-Path 'backend' 'dji_thermal_sdk' 'sample')
    $excludeDirs += (Join-Path 'backend' 'dji_thermal_sdk' 'tsdk-core' 'lib')
    $excludeDirs += (Join-Path 'backend' 'dji_thermal_sdk' 'utility' 'bin')
}

# Normalize excludes to full paths
$excludeFull = $excludeDirs | ForEach-Object { Join-Path $projectRootPath $_ }

Write-Host "Creating staging copy..."
$allFiles = Get-ChildItem -Path $projectRootPath -Recurse -File -Force -ErrorAction SilentlyContinue

foreach ($file in $allFiles) {
    $full = $file.FullName

    # Skip anything under excluded directories
    $skip = $false
    foreach ($ex in $excludeFull) {
        if ($full.StartsWith($ex, [System.StringComparison]::OrdinalIgnoreCase)) {
            $skip = $true
            break
        }
    }
    if ($skip) { continue }

    $relative = $full.Substring($projectRootPath.Length).TrimStart('\','/')
    $dest = Join-Path $staging $relative
    $destDir = Split-Path -Parent $dest
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item -LiteralPath $full -Destination $dest -Force
}

Write-Host "Creating zip: $zipPath"
if (Test-Path $zipPath) { Remove-Item -Force $zipPath }
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal

Remove-Item -Recurse -Force $staging

Write-Host "Done."
Write-Host "- Git bundle: $bundlePath"
Write-Host "- Zip backup: $zipPath"
Write-Host "Tip: Copy these files to an external drive or a different cloud account."