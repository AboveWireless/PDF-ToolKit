$ErrorActionPreference = "Stop"

$releaseDir = Join-Path $PSScriptRoot "..\dist\release"
$portableDir = Join-Path $PSScriptRoot "..\dist\pdf-toolkit-gui"
$zipPath = Join-Path $releaseDir "pdf-toolkit-windows-x64.zip"
$checksumPath = Join-Path $releaseDir "pdf-toolkit-windows-x64-checksums.txt"

if (-not $env:PDF_TOOLKIT_SKIP_BUILD) {
    powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "build_gui.ps1")
}

if (-not (Test-Path $portableDir)) {
    throw "Portable app folder not found at $portableDir"
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
if (Test-Path $checksumPath) { Remove-Item $checksumPath -Force }

Push-Location (Join-Path $PSScriptRoot "..\dist")
try {
    tar.exe -a -cf $zipPath "pdf-toolkit-gui"
}
finally {
    Pop-Location
}

$hash = Get-FileHash $zipPath -Algorithm SHA256
"$($hash.Hash)  $([System.IO.Path]::GetFileName($zipPath))" | Set-Content -Path $checksumPath -Encoding utf8

Write-Host "Created release artifacts:"
Write-Host " - $zipPath"
Write-Host " - $checksumPath"
