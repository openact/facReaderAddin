Attribute VB_Name = "facReaderAddinXControl"
' F9 refresh controller for facReaderAddinX.xll.
' This module intentionally does not expose worksheet UDFs.
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function AppReaderXClearCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacClearCache" _
        () As Long
    Private Declare PtrSafe Function AppReaderXRefreshChangedCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRefreshChangedCache" _
        () As Long
#Else
    Private Declare Function AppReaderXClearCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacClearCache" _
        () As Long
    Private Declare Function AppReaderXRefreshChangedCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRefreshChangedCache" _
        () As Long
#End If

Public Sub AppReaderXClearCache()
    On Error Resume Next
    AppReaderXClearCacheDLL
End Sub

Public Sub AppReaderXRefreshNow()
    AppReaderXRefreshChangedCacheDLL
    Application.Calculate
End Sub

Public Sub AppReaderXRefreshSheet()
    AppReaderXRefreshChangedCacheDLL
    On Error GoTo Fallback
    Application.CommandBars.ExecuteMso "CalculateSheet"
    Exit Sub
Fallback:
    ActiveSheet.Calculate
End Sub

Public Sub AppReaderXRefreshFull()
    AppReaderXRefreshChangedCacheDLL
    Application.CalculateFull
End Sub

Public Sub AppReaderXRefreshFullRebuild()
    AppReaderXRefreshChangedCacheDLL
    Application.CalculateFullRebuild
End Sub

Public Sub AppReaderXInstallF9Refresh()
    Dim macroPrefix As String
    macroPrefix = "'" & ThisWorkbook.Name & "'!"
    Application.OnKey "{F9}", macroPrefix & "AppReaderXRefreshNow"
    Application.OnKey "+{F9}", macroPrefix & "AppReaderXRefreshSheet"
    Application.OnKey "^%{F9}", macroPrefix & "AppReaderXRefreshFull"
    Application.OnKey "^+%{F9}", macroPrefix & "AppReaderXRefreshFullRebuild"
End Sub

Public Sub AppReaderXUninstallF9Refresh()
    Application.OnKey "{F9}"
    Application.OnKey "+{F9}"
    Application.OnKey "^%{F9}"
    Application.OnKey "^+%{F9}"
End Sub
