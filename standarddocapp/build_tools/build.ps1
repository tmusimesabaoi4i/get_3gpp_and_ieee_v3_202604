param(
    [switch]$SkipDeps,
    [switch]$Clean
)

# Build StandardDocApp.exe via PyInstaller using a fresh venv.
# Run this from any folder; the script normalizes paths.

$ErrorActionPreference = "Stop"

$AppDir   = Split-Path -Parent $PSScriptRoot
$RepoRoot = Split-Path -Parent $AppDir
$VenvDir  = Join-Path $AppDir ".venv-build"
$Python   = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $Python) { $Python = (Get-Command py -ErrorAction SilentlyContinue).Source }
if (-not $Python) { throw "Python interpreter not found on PATH (need 'python' or 'py')." }

Write-Host "AppDir   = $AppDir"
Write-Host "RepoRoot = $RepoRoot"
Write-Host "Python   = $Python"

if (-not (Test-Path $VenvDir)) {
    Write-Host "Creating venv at $VenvDir"
    & $Python -m venv $VenvDir
}

$VenvPy = Join-Path $VenvDir "Scripts\python.exe"
if (-not (Test-Path $VenvPy)) { throw "Failed to create venv: $VenvPy not found" }

$BuildScript = Join-Path $PSScriptRoot "build_app.py"
$ExtraArgs = @()
if ($SkipDeps) { $ExtraArgs += "--skip-deps" }
if ($Clean)    { $ExtraArgs += "--clean" }

& $VenvPy $BuildScript @ExtraArgs
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)" }

$Exe = Join-Path $AppDir "dist\StandardDocApp.exe"
if (Test-Path $Exe) {
    Write-Host ""
    Write-Host "Built: $Exe"
} else {
    throw "Build appeared to succeed but $Exe is missing."
}
