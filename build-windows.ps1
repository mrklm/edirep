<# build-windows.ps1 - Edirep (Windows) portable build script (ASCII only)

Usage:
  .\build-windows.ps1 -Version 3.10.0
  .\build-windows.ps1 -Version 3.10.0 -Arch x86_64
  .\build-windows.ps1 -Version 3.10.0 -AppsDir "$env:USERPROFILE\Apps"
  .\build-windows.ps1 -Version 3.10.0 -KeepBuildDirs

Notes:
- ASCII only: no accents, no emojis, no fancy dashes, no ellipsis.
- Designed for Windows PowerShell 5.1 and PowerShell 7+.
- This script does NOT run edirep.py (you test before building).
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d+(\.\d+){1,3}$')]
  [string]$Version,

  [ValidateSet("x86_64", "arm64")]
  [string]$Arch = "x86_64",

  [switch]$KeepBuildDirs,

  [string]$AppsDir = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Allow script execution for this process only (best effort).
try { Set-ExecutionPolicy -Scope Process Bypass -Force } catch {}

function Info([string]$msg) { Write-Host ("[INFO]  " + $msg) }
function Ok([string]$msg)   { Write-Host ("[OK]    " + $msg) }
function Warn([string]$msg) { Write-Host ("[WARN]  " + $msg) }
function Fail([string]$msg) { Write-Host ("[ERROR] " + $msg); exit 1 }

# --- 0) Context / paths ---
$Root = (Get-Location).Path

# Adjust these if your entrypoint or layout differs.
$AppName     = "Edirep"
$MainPy      = Join-Path $Root "edirep.py"
$AssetsDir   = Join-Path $Root "assets"
$ReqFile     = Join-Path $Root "requirements.txt"

$IconIco     = Join-Path $AssetsDir "logo.ico"
$LogoPng     = Join-Path $AssetsDir "logo.png"

$VenvDir     = Join-Path $Root ".venv"
$VenvPy      = Join-Path $VenvDir "Scripts\python.exe"
$Activate    = Join-Path $VenvDir "Scripts\Activate.ps1"

$DistExeRel  = "dist\$AppName.exe"
$DistExe     = Join-Path $Root $DistExeRel

$OutDir      = Join-Path $Root "releases"
$ZipName     = "$AppName-v$Version-windows-$Arch.zip"
$ZipPath     = Join-Path $OutDir $ZipName
$HashPath    = "$ZipPath.sha256"

Info "Build Windows - $AppName v$Version ($Arch)"
Info "Repo: $Root"

# --- 1) Required checks ---
Info "Checking required files..."
if (!(Test-Path $MainPy))    { Fail "Missing entrypoint: $MainPy" }
if (!(Test-Path $AssetsDir)) { Fail "Missing folder: $AssetsDir" }

if (!(Test-Path $ReqFile))   { Warn "Missing requirements.txt (will only install pyinstaller)." }

if (!(Test-Path $LogoPng))   { Warn "Missing assets\logo.png (recommended)." }
if (!(Test-Path $IconIco))   { Warn "Missing assets\logo.ico (exe will use default icon)." }

Ok "Base structure OK."

# --- 2) Venv + deps ---
if (!(Test-Path $VenvPy)) {
  Info "Creating venv .venv..."
  python -m venv $VenvDir
  if (!(Test-Path $VenvPy)) { Fail "Venv creation failed (missing .venv\Scripts\python.exe)." }
  Ok "Venv created."
} else {
  Info "Venv already exists."
}

# Activation is optional because we always call $VenvPy, but convenient for the shell.
if (Test-Path $Activate) {
  Info "Activating venv..."
  . $Activate
} else {
  Warn "Activate.ps1 not found, continuing using venv python directly."
}

Info "Upgrading pip..."
& $VenvPy -m pip install --upgrade pip | Out-Host

if (Test-Path $ReqFile) {
  Info "Installing dependencies from requirements.txt..."
  & $VenvPy -m pip install -r $ReqFile | Out-Host
}

Info "Installing PyInstaller..."
& $VenvPy -m pip install pyinstaller | Out-Host
Ok "Dependencies OK."

# --- 3) Clean build artifacts ---
if (-not $KeepBuildDirs) {
  Info "Cleaning build/dist/cache/spec..."
  Remove-Item -Recurse -Force (Join-Path $Root "build")      -ErrorAction SilentlyContinue
  Remove-Item -Recurse -Force (Join-Path $Root "dist")       -ErrorAction SilentlyContinue
  Remove-Item -Recurse -Force (Join-Path $Root "__pycache__")-ErrorAction SilentlyContinue
  Get-ChildItem -Path $Root -Filter "*.spec" -ErrorAction SilentlyContinue | `
    Remove-Item -Force -ErrorAction SilentlyContinue
  Ok "Clean OK."
} else {
  Warn "KeepBuildDirs enabled: build/dist/spec will be kept."
}

# --- 4) PyInstaller build ---
Info "Running PyInstaller..."

# NOTE: On Windows, --add-data uses "src;dest"
$pyiArgs = @(
  "--name", $AppName,
  "--onefile",
  "--windowed",
  "--add-data", "assets;assets"
)

if (Test-Path $IconIco) {
  $pyiArgs += @("--icon", $IconIco)
}

$pyiArgs += @($MainPy)

& $VenvPy -m PyInstaller @pyiArgs | Out-Host

if (!(Test-Path $DistExe)) { Fail "Build finished but exe not found: $DistExeRel" }
Ok "EXE generated: $DistExeRel"

# --- 5) Zip + sha256 ---
if (!(Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

Remove-Item -Force $ZipPath  -ErrorAction SilentlyContinue
Remove-Item -Force $HashPath -ErrorAction SilentlyContinue

Info "Creating ZIP: $ZipName"
Compress-Archive -Path $DistExe -DestinationPath $ZipPath -Force
if (!(Test-Path $ZipPath)) { Fail "ZIP not created: $ZipPath" }
Ok "ZIP created: releases\$ZipName"

Info "Computing SHA256..."
$hash = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToLower()
"$hash  $ZipName" | Set-Content -Encoding ASCII -Path $HashPath
Ok "SHA256 file created: releases\$ZipName.sha256"

# --- 6) Optional: copy exe to AppsDir ---
if ($AppsDir.Trim() -ne "") {
  if (!(Test-Path $AppsDir)) {
    Info "Creating AppsDir: $AppsDir"
    New-Item -ItemType Directory -Path $AppsDir | Out-Null
  }
  $DestExe = Join-Path $AppsDir "$AppName.exe"
  Copy-Item -Force $DistExe $DestExe
  Ok "EXE copied to: $DestExe"
}

# --- 7) Summary ---
Write-Host ""
Write-Host "==================== SUMMARY ===================="
Write-Host ("App         : " + $AppName)
Write-Host ("Version     : v" + $Version)
Write-Host ("EXE         : " + $DistExeRel)
Write-Host ("ZIP release : releases\" + $ZipName)
Write-Host ("SHA256      : releases\" + $ZipName + ".sha256")
if ($AppsDir.Trim() -ne "") { Write-Host ("AppsDir exe  : " + (Join-Path $AppsDir "$AppName.exe")) }
Write-Host "================================================"
Write-Host ""
Ok "Build Windows ok mon pote."
