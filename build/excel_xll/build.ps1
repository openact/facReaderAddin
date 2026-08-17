Set-Location $PSScriptRoot

$ErrorActionPreference = "Stop"

$BuildDir = Split-Path $PSScriptRoot -Parent
$AppReaderDir = Split-Path $BuildDir -Parent
$DistDir = Join-Path $AppReaderDir "dist"
if (-not (Test-Path $DistDir)) {
    New-Item -ItemType Directory -Path $DistDir | Out-Null
}

$gcc = Get-Command gcc -ErrorAction Stop
$src = Join-Path $PSScriptRoot "src\appReaderX.c"
$xll = Join-Path $DistDir "facReaderAddinX.xll"
$controllerXlam = Join-Path $DistDir "facReaderAddinXControl.xlam"
$appReaderDll = Join-Path $AppReaderDir "dist\facReaderAddin.dll"
$machineDllPath = "C:\Tools\OpenAct\Addins\facReaderAddin.dll"

if (-not (Test-Path $appReaderDll)) {
    throw "facReaderAddin.dll not found: $appReaderDll. Run app\facReaderAddin\build\dll\build.ps1 first."
}

Write-Host ">>> Building facReaderAddinX.xll ..." -ForegroundColor Cyan
& $gcc.Source -shared -O2 -Wall -Wextra -Wno-cast-function-type -static -static-libgcc -mwindows -o $xll $src
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host ">>> Building facReaderAddinXControl.xlam ..." -ForegroundColor Cyan
$rxDefault = [regex]::Escape('C:\Tools\OpenAct\Addins\facReaderAddin.dll')
$controller = Get-Content (Join-Path $PSScriptRoot "appReaderXControl.bas") -Raw
$controller = $controller -replace $rxDefault, $machineDllPath
$outController = Join-Path ([System.IO.Path]::GetTempPath()) ("facReaderAddinXControl_" + [guid]::NewGuid().ToString("N") + ".bas")
$controller | Set-Content $outController -Encoding Default
$excel = $null
$wb = $null
$controllerBuilt = $false
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $null = $wb.VBProject.VBComponents.Import($outController)
    $thisWorkbook = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
    $thisWorkbookCode = @'
Private Sub Workbook_Open()
    AppReaderXInstallF9Refresh
End Sub

Private Sub Workbook_AddinInstall()
    AppReaderXInstallF9Refresh
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    AppReaderXUninstallF9Refresh
End Sub

Private Sub Workbook_AddinUninstall()
    AppReaderXUninstallF9Refresh
End Sub
'@
    $thisWorkbook.InsertLines(1, $thisWorkbookCode)

    $xlOpenXMLAddIn = 55
    $wb.SaveAs($controllerXlam, $xlOpenXMLAddIn)
    $controllerBuilt = $true
}
catch {
    Write-Warning "自动打包 facReaderAddinXControl.xlam 失败：$($_.Exception.Message)"
    Write-Host "    可手工 fallback：导入 build\excel_xll\appReaderXControl.bas 后另存为 facReaderAddinXControl.xlam。" -ForegroundColor Yellow
    Write-Host "    若报 VBA 工程访问错误，请在 Excel 信任中心启用“信任对 VBA 工程对象模型的访问”。" -ForegroundColor Yellow
}
finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if (Test-Path $outController) {
        Remove-Item $outController -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$oldReadme = Join-Path $DistDir "facReaderAddinX_README.md"
if (Test-Path $oldReadme) {
    Remove-Item $oldReadme -Force
}

Write-Host ">>> Built:" -ForegroundColor Green
Write-Host "    $xll"
if ($controllerBuilt) {
    Write-Host "    $controllerXlam"
}
Write-Host "    $(Join-Path $DistDir 'facReaderAddin.dll')"
Write-Host "    $(Join-Path $DistDir 'facReaderAddin_README.md')"
Write-Host ""
Write-Host "Load facReaderAddinX.xll and facReaderAddinXControl.xlam in Excel Add-ins. Keep facReaderAddin.dll in the same folder." -ForegroundColor Yellow
