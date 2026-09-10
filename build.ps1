# Build a portable onedir bundle of the Walmart Shelf Assistant.
#
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
#
# Output: dist\WalmartShelfAssistant\   <- copy this WHOLE FOLDER to the target PC
#
# The target PC needs Microsoft Excel installed. Nothing else - Python is
# bundled (verify: the running exe loads python314.dll from its own _internal).
#
# excel_mapper.ps1 is shipped as a plain file inside _internal\, so mapping
# rules can be tweaked after packaging without rebuilding.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "==> Cleaning previous build"
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
Remove-Item -Force WalmartShelfAssistant.spec -ErrorAction SilentlyContinue

Write-Host "==> Building"
py -m PyInstaller --noconfirm --clean --windowed `
    --name WalmartShelfAssistant `
    --add-data "excel_mapper.ps1;." `
    app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

$exe = Join-Path $PSScriptRoot "dist\WalmartShelfAssistant\WalmartShelfAssistant.exe"
if (-not (Test-Path $exe)) { throw "Build reported success but $exe is missing" }

# The mapper must sit next to app.py's frozen APP_DIR (= _internal), or the GUI
# will launch fine and then fail at "开始填充" with a missing-script error.
$mapper = Join-Path $PSScriptRoot "dist\WalmartShelfAssistant\_internal\excel_mapper.ps1"
if (-not (Test-Path $mapper)) { throw "excel_mapper.ps1 was not bundled into _internal\" }

Write-Host ""
Write-Host "==> Done"
Write-Host "    $exe"
Write-Host "    Copy the whole dist\WalmartShelfAssistant folder to the target PC."
Write-Host "    Target PC requirement: Microsoft Excel (for COM automation)."
Write-Host "    First launch on another PC will show a SmartScreen warning (unsigned exe)."
