# Install the Walmart Shelf Assistant into the current user's profile and put a
# shortcut on the Desktop.
#
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
#
# Per-user install (no administrator rights needed):
#   program  -> %LOCALAPPDATA%\Programs\WalmartShelfAssistant\
#   shortcut -> <Desktop>\沃尔玛上架助手.lnk
#
# Deliberately does NOT install from dist\ in place. build.ps1 deletes and
# recreates dist\ on every run, which would break a shortcut pointing into it.
# Copying to %LOCALAPPDATA% decouples the installed app from the build output.
#
# To remove: delete the folder above and the Desktop shortcut.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$source = Join-Path $PSScriptRoot "dist\WalmartShelfAssistant"
$target = Join-Path $env:LOCALAPPDATA "Programs\WalmartShelfAssistant"
$exeName = "WalmartShelfAssistant.exe"

if (-not (Test-Path (Join-Path $source $exeName))) {
    Write-Host "==> No build found, running build.ps1 first"
    & (Join-Path $PSScriptRoot "build.ps1")
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $source $exeName))) {
        throw "Build did not produce $source\$exeName"
    }
}

# Fail early rather than install a bundle that will die at the first mapping.
if (-not (Test-Path (Join-Path $source "_internal\excel_mapper.ps1"))) {
    throw "excel_mapper.ps1 is missing from the build - rerun build.ps1"
}

Write-Host "==> Installing to $target"
if (Test-Path $target) {
    Remove-Item -Recurse -Force $target
}
New-Item -ItemType Directory -Path $target -Force | Out-Null
Copy-Item -Path (Join-Path $source "*") -Destination $target -Recurse -Force

$exe = Join-Path $target $exeName
if (-not (Test-Path $exe)) { throw "Copy failed: $exe not found" }

# Shortcut on the real Desktop (may be redirected, e.g. to D:\ or OneDrive).
$desktop = [Environment]::GetFolderPath("Desktop")
$link = Join-Path $desktop "沃尔玛上架助手.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = $target
$shortcut.IconLocation = "$exe,0"
$shortcut.Description = "沃尔玛上架助手 - 把文件 A 的商品信息填入文件 B 的沃尔玛上架模板"
$shortcut.Save()

Write-Host ""
Write-Host "==> Installed"
Write-Host "    Program : $exe"
Write-Host "    Shortcut: $link"
Write-Host ""
Write-Host "    Requires Microsoft Excel on this machine (COM automation)."
Write-Host "    Path settings live in %LOCALAPPDATA%\WalmartShelfAssistant\settings.json"
