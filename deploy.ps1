param(
    [ValidateSet("dll", "xlam", "xll", "all")]
    [string]$Target = "all"
)

Set-Location $PSScriptRoot
$ErrorActionPreference = "Stop"

if ($Target -eq "dll" -or $Target -eq "all") {
    & (Join-Path $PSScriptRoot "build\dll\build.ps1")
    Set-Location $PSScriptRoot
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if ($Target -eq "xlam" -or $Target -eq "all") {
    & (Join-Path $PSScriptRoot "build\excel_xlam\deploy.ps1")
    Set-Location $PSScriptRoot
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if ($Target -eq "xll" -or $Target -eq "all") {
    & (Join-Path $PSScriptRoot "build\excel_xll\build.ps1")
    Set-Location $PSScriptRoot
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
