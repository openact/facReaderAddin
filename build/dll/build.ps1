Set-Location $PSScriptRoot

$ErrorActionPreference = "Stop"

$BuildDir = Split-Path $PSScriptRoot -Parent
$AppReaderDir = Split-Path $BuildDir -Parent
$OutputDir = Join-Path $AppReaderDir "dist"

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$env:CGO_ENABLED = "1"
$env:GOOS        = "windows"
$env:GOARCH      = "amd64"

Write-Host ">>> Building facReaderAddin.dll (windows/amd64) ..." -ForegroundColor Cyan
go build -buildmode=c-shared -o facReaderAddin.dll .
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Copy-Item -Path (Join-Path $PSScriptRoot "facReaderAddin.dll") -Destination $OutputDir -Force

Write-Host ">>> Built:" -ForegroundColor Green
Write-Host "    $(Join-Path $OutputDir 'facReaderAddin.dll')"
