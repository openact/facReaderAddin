Attribute VB_Name = "facReaderAddin"
' .fac query module.
' Call EProj_Result with business parameters, or ERead_Result with table-order coordinates.
Option Explicit

Private gFacDimsCache As Object

#If VBA7 Then
    Private Declare PtrSafe Function AppReaderProjResultDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "EProj_Result" _
        (ByVal filePath As String, ByVal spCode As String, ByVal resultType As String, _
         ByVal variable As String, ByVal timePeriod As String, ByVal simID As String, _
         ByRef valueOut As Double) As Long
    Private Declare PtrSafe Function AppReaderReadResultDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "ERead_Result" _
        (ByVal filePath As String, ByVal coordCount As Long, _
         ByVal key1 As String, ByVal key2 As String, ByVal key3 As String, ByVal key4 As String, _
         ByVal key5 As String, ByVal key6 As String, ByVal key7 As String, ByVal key8 As String, _
         ByVal key9 As String, ByVal key10 As String, ByVal key11 As String, ByVal key12 As String, _
         ByRef valueOut As Double) As Long
    Private Declare PtrSafe Function AppReaderReadBatchDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "ERead_Batch" _
        (ByVal filePath As String, ByVal rowKeysPacked As String, ByVal colKeysPacked As String, _
         ByVal outputPath As String) As Long
    Private Declare PtrSafe Function AppReaderProjBatchDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "EProj_Batch" _
        (ByVal filePath As String, ByVal product As String, ByVal simID As String, _
         ByVal projKeysPacked As String, ByVal colKeysPacked As String, ByVal outputPath As String) As Long
    Private Declare PtrSafe Function AppReaderNumDimsDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacNumDims" _
        (ByVal filePath As String) As Long
    Private Declare PtrSafe Function AppReaderLoadErrorDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacLoadError" _
        (ByVal filePath As String, ByVal buffer As String, ByVal capacity As Long) As Long
    Private Declare PtrSafe Function AppReaderReleaseDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRelease" _
        (ByVal filePath As String) As Long
    Private Declare PtrSafe Function AppReaderClearCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacClearCache" _
        () As Long
    Private Declare PtrSafe Function AppReaderRefreshChangedCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRefreshChangedCache" _
        () As Long
#Else
    Private Declare Function AppReaderProjResultDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "EProj_Result" _
        (ByVal filePath As String, ByVal spCode As String, ByVal resultType As String, _
         ByVal variable As String, ByVal timePeriod As String, ByVal simID As String, _
         ByRef valueOut As Double) As Long
    Private Declare Function AppReaderReadResultDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "ERead_Result" _
        (ByVal filePath As String, ByVal coordCount As Long, _
         ByVal key1 As String, ByVal key2 As String, ByVal key3 As String, ByVal key4 As String, _
         ByVal key5 As String, ByVal key6 As String, ByVal key7 As String, ByVal key8 As String, _
         ByVal key9 As String, ByVal key10 As String, ByVal key11 As String, ByVal key12 As String, _
         ByRef valueOut As Double) As Long
    Private Declare Function AppReaderReadBatchDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "ERead_Batch" _
        (ByVal filePath As String, ByVal rowKeysPacked As String, ByVal colKeysPacked As String, _
         ByVal outputPath As String) As Long
    Private Declare Function AppReaderProjBatchDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "EProj_Batch" _
        (ByVal filePath As String, ByVal product As String, ByVal simID As String, _
         ByVal projKeysPacked As String, ByVal colKeysPacked As String, ByVal outputPath As String) As Long
    Private Declare Function AppReaderNumDimsDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacNumDims" _
        (ByVal filePath As String) As Long
    Private Declare Function AppReaderLoadErrorDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacLoadError" _
        (ByVal filePath As String, ByVal buffer As String, ByVal capacity As Long) As Long
    Private Declare Function AppReaderReleaseDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRelease" _
        (ByVal filePath As String) As Long
    Private Declare Function AppReaderClearCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacClearCache" _
        () As Long
    Private Declare Function AppReaderRefreshChangedCacheDLL Lib "C:\Tools\OpenAct\Addins\facReaderAddin.dll" Alias "FacRefreshChangedCache" _
        () As Long
#End If

Public Sub AppReaderClearCache()
    On Error Resume Next
    Set gFacDimsCache = Nothing
    AppReaderClearCacheDLL
End Sub

Public Sub AppReaderRefreshNow()
    AppReaderRefreshChangedCacheDLL
    Application.Calculate
End Sub

Public Sub AppReaderRefreshSheet()
    AppReaderRefreshChangedCacheDLL
    On Error GoTo Fallback
    Application.CommandBars.ExecuteMso "CalculateSheet"
    Exit Sub
Fallback:
    ActiveSheet.Calculate
End Sub

Public Sub AppReaderRefreshFull()
    AppReaderRefreshChangedCacheDLL
    Application.CalculateFull
End Sub

Public Sub AppReaderRefreshFullRebuild()
    AppReaderRefreshChangedCacheDLL
    Application.CalculateFullRebuild
End Sub

Public Sub AppReaderInstallF9Refresh()
    Dim macroPrefix As String
    macroPrefix = "'" & ThisWorkbook.Name & "'!"
    Application.OnKey "{F9}", macroPrefix & "AppReaderRefreshNow"
    Application.OnKey "+{F9}", macroPrefix & "AppReaderRefreshSheet"
    Application.OnKey "^%{F9}", macroPrefix & "AppReaderRefreshFull"
    Application.OnKey "^+%{F9}", macroPrefix & "AppReaderRefreshFullRebuild"
End Sub

Public Sub AppReaderUninstallF9Refresh()
    Application.OnKey "{F9}"
    Application.OnKey "+{F9}"
    Application.OnKey "^%{F9}"
    Application.OnKey "^+%{F9}"
End Sub

Private Function ArgSeparator() As String
    ArgSeparator = Chr$(31)
End Function

Private Function RowSeparator() As String
    RowSeparator = Chr$(30)
End Function

Private Function CachedFacNumDims(ByVal filePath As String) As Long
    CachedFacNumDims = AppReaderNumDimsDLL(filePath)
End Function

Private Function FacPathError(ByVal resVault As String, ByVal resID As String, ByVal filePath As String) As String
    Dim vaultPath As String
    vaultPath = CStr(resVault)
    If Right$(vaultPath, 1) = "\" Then vaultPath = Left$(vaultPath, Len(vaultPath) - 1)
    If Dir(vaultPath, vbDirectory) = "" Then
        FacPathError = "#ERROR: vault dir not found."
        Exit Function
    End If

    Dim resultDir As String
    resultDir = vaultPath & "\" & CStr(resID)
    If Dir(resultDir, vbDirectory) = "" Then
        FacPathError = "#ERROR: result dir not found."
        Exit Function
    End If

    FacPathError = "#ERROR: file not found."
End Function

Private Function BuildFacPath(resVault As String, resID As String, fileName As String) As String
    Dim base As String: base = CStr(resVault)
    Dim facName As String: facName = CStr(fileName)
    If Right(base, 1) <> "\" Then base = base & "\"
    If LCase$(Right(facName, 4)) <> ".fac" Then facName = facName & ".fac"
    BuildFacPath = base & CStr(resID) & "\" & facName
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

' Normalize Excel input to .fac period keys.
' Supported examples:
'   2025/12/31 -> 202512
'   Excel serial date (e.g. 46022) -> 202512
'   202512 -> 202512
'   2025 -> 2025
'   * -> "" (lookup mode does not support wildcard period)
Private Function NormalizeTimePeriod(ByVal raw As Variant) As String
    Dim s As String
    s = Trim$(CStr(raw))

    If s = "" Then
        NormalizeTimePeriod = ""
        Exit Function
    End If

    If s = "*" Then
        NormalizeTimePeriod = ""
        Exit Function
    End If

    If IsNumeric(raw) Then
        Dim n As Double
        n = CDbl(raw)

        If Abs(n - Fix(n)) < 0.0000001 Then
            Dim whole As Long
            whole = CLng(Fix(n))
            If whole >= 190001 And whole <= 299912 Then
                NormalizeTimePeriod = CStr(whole)
                Exit Function
            End If
            If whole >= 1900 And whole <= 2999 Then
                NormalizeTimePeriod = CStr(whole)
                Exit Function
            End If
        End If

        If IsDate(raw) Then
            NormalizeTimePeriod = Format$(CDate(raw), "yyyymm")
            Exit Function
        End If
    End If

    If IsDate(raw) Then
        NormalizeTimePeriod = Format$(CDate(raw), "yyyymm")
        Exit Function
    End If

    Dim digits As String
    digits = DigitsOnly(s)
    If Len(digits) = 6 Or Len(digits) = 4 Then
        NormalizeTimePeriod = digits
    Else
        NormalizeTimePeriod = s
    End If
End Function

Private Function StatusMessage(ByVal status As Long) As String
    Select Case status
        Case 0: StatusMessage = ""
        Case 1: StatusMessage = "lookup coordinates not found"
        Case 2: StatusMessage = "dimension mismatch"
        Case 3: StatusMessage = "invalid argument"
        Case 4: StatusMessage = "numeric parse error"
        Case 5: StatusMessage = "failed to write batch output"
        Case 6: StatusMessage = "file not found"
        Case 7: StatusMessage = "FAC load error"
        Case 8: StatusMessage = "wildcard is not allowed"
        Case Else: StatusMessage = "status " & CStr(status)
    End Select
End Function

Private Function LoadErrorMessage(ByVal filePath As String) As String
    Dim buffer As String
    buffer = String$(768, vbNullChar)
    Dim length As Long
    length = AppReaderLoadErrorDLL(filePath, buffer, Len(buffer))
    If length > 0 Then
        LoadErrorMessage = "#ERROR: " & Left$(buffer, length)
    Else
        LoadErrorMessage = "#ERROR: failed to load FAC " & filePath & "."
    End If
End Function

Private Function ReadErrorMessage(ByVal status As Long, ByVal filePath As String, _
                                  ByVal expectedCoords As Long, ByVal gotCoords As Long) As String
    Select Case status
        Case 1
            ReadErrorMessage = "#ERROR: no match found in " & filePath & "."
        Case 2
            If expectedCoords > 0 Then
                ReadErrorMessage = "#ERROR: expected " & CStr(expectedCoords) & " lookup coordinates, got " & CStr(gotCoords) & " for " & filePath & "."
            Else
                ReadErrorMessage = "#ERROR: dimension mismatch for " & filePath & " (got " & CStr(gotCoords) & " coordinates)."
            End If
        Case 3
            ReadErrorMessage = "#ERROR: invalid argument for " & filePath & "."
        Case 4
            ReadErrorMessage = "#ERROR: numeric parse error in " & filePath & "."
        Case 6
            ReadErrorMessage = "#ERROR: file not found: " & filePath & "."
        Case 7
            ReadErrorMessage = LoadErrorMessage(filePath)
        Case 8
            ReadErrorMessage = "#ERROR: invalid argument for " & filePath & " (wildcard ""*"" is not allowed)."
        Case Else
            ReadErrorMessage = "#ERROR: " & StatusMessage(status) & " for " & filePath & "."
    End Select
End Function

Private Function ProjErrorMessage(ByVal status As Long, ByVal filePath As String, _
                                  ByVal spCode As String, ByVal variable As String, _
                                  ByVal timePeriod As String, ByVal simID As String) As String
    Select Case status
        Case 1
            ProjErrorMessage = "#ERROR: no match found in " & filePath & "."
        Case 2
            ProjErrorMessage = "#ERROR: dimension mismatch for " & filePath & "."
        Case 3
            ProjErrorMessage = "#ERROR: invalid argument for " & filePath & "."
        Case 4
            ProjErrorMessage = "#ERROR: numeric parse error in " & filePath & "."
        Case 6
            ProjErrorMessage = "#ERROR: file not found: " & filePath & "."
        Case 7
            ProjErrorMessage = LoadErrorMessage(filePath)
        Case Else
            ProjErrorMessage = "#ERROR: " & StatusMessage(status) & " for " & filePath & "."
    End Select
End Function

Private Function ContainsReservedSeparator(ByVal v As String) As Boolean
    ContainsReservedSeparator = (InStr(1, v, ArgSeparator(), vbBinaryCompare) > 0 Or _
                                 InStr(1, v, RowSeparator(), vbBinaryCompare) > 0)
End Function

Private Function CellValue2(ByVal values As Variant, ByVal r As Long, ByVal c As Long, _
                            ByVal rowCount As Long, ByVal colCount As Long) As Variant
    If rowCount = 1 And colCount = 1 Then
        CellValue2 = values
    Else
        CellValue2 = values(r, c)
    End If
End Function

Private Function BuildPackedRows(ByVal rng As Range, ByVal normalizePeriods As Boolean, ByRef errText As String) As String
    Dim rowCount As Long: rowCount = rng.Rows.Count
    Dim colCount As Long: colCount = rng.Columns.Count
    Dim values As Variant: values = rng.Value2
    Dim rowParts() As String
    ReDim rowParts(1 To rowCount)

    Dim r As Long
    Dim c As Long
    For r = 1 To rowCount
        Dim fields() As String
        ReDim fields(1 To colCount)
        For c = 1 To colCount
            Dim v As String
            If normalizePeriods Then
                v = NormalizeTimePeriod(CellValue2(values, r, c, rowCount, colCount))
                If v = "" Then
                    errText = "invalid period key at row " & CStr(r) & ", column " & CStr(c) & "."
                    Exit Function
                End If
            Else
                v = CStr(CellValue2(values, r, c, rowCount, colCount))
            End If
            If ContainsReservedSeparator(v) Then
                errText = "lookup coordinate contains a reserved separator character."
                Exit Function
            End If
            If Trim$(v) = "*" Then
                errText = "invalid lookup coordinate at row " & CStr(r) & ", column " & CStr(c) & "."
                Exit Function
            End If
            fields(c) = v
        Next c
        rowParts(r) = Join(fields, ArgSeparator())
    Next r

    BuildPackedRows = Join(rowParts, RowSeparator())
End Function

Private Function BuildPackedVector(ByVal rng As Range, ByVal normalizePeriods As Boolean, ByRef errText As String) As String
    Dim rowCount As Long: rowCount = rng.Rows.Count
    Dim colCount As Long: colCount = rng.Columns.Count
    Dim values As Variant: values = rng.Value2
    Dim parts() As String
    ReDim parts(1 To rowCount * colCount)

    Dim r As Long
    Dim c As Long
    Dim k As Long: k = 1
    For r = 1 To rowCount
        For c = 1 To colCount
            Dim v As String
            If normalizePeriods Then
                v = NormalizeTimePeriod(CellValue2(values, r, c, rowCount, colCount))
                If v = "" Then
                    errText = "invalid period key at row " & CStr(r) & ", column " & CStr(c) & "."
                    Exit Function
                End If
            Else
                v = CStr(CellValue2(values, r, c, rowCount, colCount))
            End If
            If ContainsReservedSeparator(v) Then
                errText = "lookup coordinate contains a reserved separator character."
                Exit Function
            End If
            If Trim$(v) = "*" Then
                errText = "invalid lookup coordinate at row " & CStr(r) & ", column " & CStr(c) & "."
                Exit Function
            End If
            parts(k) = v
            k = k + 1
        Next c
    Next r

    BuildPackedVector = Join(parts, ArgSeparator())
End Function

Private Function TempOutputPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    TempOutputPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, fso.GetTempName())
End Function

Private Function ParseResultToken(ByVal token As String) As Variant
    Select Case token
        Case "#N/A"
            ParseResultToken = CVErr(xlErrNA)
        Case "#VALUE!"
            ParseResultToken = CVErr(xlErrValue)
        Case Else
            ParseResultToken = CDbl(token)
    End Select
End Function

Private Function ReadTsvMatrix(ByVal outputPath As String, ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim fileNo As Integer
    fileNo = FreeFile
    Open outputPath For Binary Access Read As #fileNo
    Dim text As String
    text = Input$(LOF(fileNo), #fileNo)
    Close #fileNo

    Dim lines() As String
    lines = Split(text, vbLf)

    Dim out() As Variant
    ReDim out(1 To rowCount, 1 To colCount)

    Dim r As Long
    Dim c As Long
    For r = 1 To rowCount
        Dim line As String
        If r - 1 <= UBound(lines) Then
            line = lines(r - 1)
        Else
            line = ""
        End If
        If Right$(line, 1) = vbCr Then line = Left$(line, Len(line) - 1)
        Dim fields() As String
        fields = Split(line, vbTab)
        For c = 1 To colCount
            If c - 1 <= UBound(fields) Then
                out(r, c) = ParseResultToken(fields(c - 1))
            Else
                out(r, c) = CVErr(xlErrNA)
            End If
        Next c
    Next r

    ReadTsvMatrix = out
End Function

' Single worksheet UDF.
' Parameters:
'   resVault   - full path of vault folder (e.g. ...\rbc2512_4)
'   resID      - sub folder such as E_1
'   product    - file name with or without .fac
'   spCode     - SP_CODE exact value (e.g. 0/1/2)
'   resultType - reserved (kept for compatibility; not used)
'   variable   - VAR_NAME exact value
'   timePeriod - supports yyyyMM, yyyy, Excel date cell
'   simID      - SIM_ID exact value (optional, default 0)
' Returns #REF! when file does not exist, #VALUE! on other errors.
Public Function EProj_Result(resVault As String, resID As String, product As String, _
                             spCode As Variant, resultType As String, variable As String, _
                             timePeriod As Variant, Optional simID As Variant = "0") As Variant
    On Error GoTo ErrHandler
    Application.Volatile True

    Dim filePath As String: filePath = BuildFacPath(resVault, resID, product)
    Dim periodKey As String: periodKey = NormalizeTimePeriod(timePeriod)

    If periodKey = "" Then
        EProj_Result = "#ERROR: invalid timePeriod for " & filePath & " (expected yyyyMM, yyyy, or a date cell)."
        Exit Function
    End If

    Dim valueOut As Double
    Dim status As Long
    status = AppReaderProjResultDLL(filePath, CStr(spCode), CStr(resultType), CStr(variable), periodKey, CStr(simID), valueOut)
    If status = 0 Then
        EProj_Result = valueOut
    Else
        EProj_Result = ProjErrorMessage(status, filePath, CStr(spCode), CStr(variable), periodKey, CStr(simID))
    End If
    Exit Function
ErrHandler:
    EProj_Result = CVErr(xlErrValue)
End Function

' Generic worksheet UDF.
' Parameters:
'   resVault - full path of vault folder (e.g. ...\rbc2512_4)
'   resID    - sub folder such as E_1
'   filename - .fac file name with or without .fac
'   keys     - table coordinates in .fac dimension order:
'              all row dimensions first, then the column dimension
' Returns #REF! when file does not exist, "#ERROR: ..." on argument errors.
Public Function ERead_Result(resVault As String, resID As String, filename As String, _
                             Optional key1 As Variant, Optional key2 As Variant, _
                             Optional key3 As Variant, Optional key4 As Variant, _
                             Optional key5 As Variant, Optional key6 As Variant, _
                             Optional key7 As Variant, Optional key8 As Variant, _
                             Optional key9 As Variant, Optional key10 As Variant, _
                             Optional key11 As Variant, Optional key12 As Variant) As Variant
    On Error GoTo ErrHandler
    Application.Volatile True

    Dim filePath As String
    filePath = BuildFacPath(resVault, resID, filename)

    Dim gotCoords As Long
    If Not IsMissing(key12) Then
        gotCoords = 12
    ElseIf Not IsMissing(key11) Then
        gotCoords = 11
    ElseIf Not IsMissing(key10) Then
        gotCoords = 10
    ElseIf Not IsMissing(key9) Then
        gotCoords = 9
    ElseIf Not IsMissing(key8) Then
        gotCoords = 8
    ElseIf Not IsMissing(key7) Then
        gotCoords = 7
    ElseIf Not IsMissing(key6) Then
        gotCoords = 6
    ElseIf Not IsMissing(key5) Then
        gotCoords = 5
    ElseIf Not IsMissing(key4) Then
        gotCoords = 4
    ElseIf Not IsMissing(key3) Then
        gotCoords = 3
    ElseIf Not IsMissing(key2) Then
        gotCoords = 2
    ElseIf Not IsMissing(key1) Then
        gotCoords = 1
    Else
        gotCoords = 0
    End If

    If gotCoords = 0 Then
        ERead_Result = "#ERROR: expected " & CStr(CachedFacNumDims(filePath)) & " lookup coordinates, got 0 for " & filePath & "."
        Exit Function
    End If

    Dim k1 As String, k2 As String, k3 As String, k4 As String
    Dim k5 As String, k6 As String, k7 As String, k8 As String
    Dim k9 As String, k10 As String, k11 As String, k12 As String
    If gotCoords >= 1 Then k1 = CStr(key1)
    If gotCoords >= 2 Then k2 = CStr(key2)
    If gotCoords >= 3 Then k3 = CStr(key3)
    If gotCoords >= 4 Then k4 = CStr(key4)
    If gotCoords >= 5 Then k5 = CStr(key5)
    If gotCoords >= 6 Then k6 = CStr(key6)
    If gotCoords >= 7 Then k7 = CStr(key7)
    If gotCoords >= 8 Then k8 = CStr(key8)
    If gotCoords >= 9 Then k9 = CStr(key9)
    If gotCoords >= 10 Then k10 = CStr(key10)
    If gotCoords >= 11 Then k11 = CStr(key11)
    If gotCoords >= 12 Then k12 = CStr(key12)

    Dim valueOut As Double
    Dim status As Long
    status = AppReaderReadResultDLL(filePath, gotCoords, _
        k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, _
        valueOut)
    If status = 0 Then
        ERead_Result = valueOut
    ElseIf status = 1 Then
        ERead_Result = ReadErrorMessage(status, filePath, CachedFacNumDims(filePath), gotCoords)
    ElseIf status = 2 Then
        ERead_Result = ReadErrorMessage(status, filePath, CachedFacNumDims(filePath), gotCoords)
    ElseIf status = 6 Then
        ERead_Result = ReadErrorMessage(status, filePath, CachedFacNumDims(filePath), gotCoords)
    Else
        ERead_Result = ReadErrorMessage(status, filePath, CachedFacNumDims(filePath), gotCoords)
    End If
    Exit Function
ErrHandler:
    ERead_Result = "#ERROR: invalid ERead_Result arguments."
End Function

' Batch generic table-order lookup.
' rowKeys contains all row dimensions (NumDims - 1 columns); colKeys contains column dimension keys.
Public Function ERead_Results(resVault As String, resID As String, filename As String, _
                              rowKeys As Range, colKeys As Range) As Variant
    On Error GoTo ErrHandler
    Application.Volatile True

    Dim filePath As String
    filePath = BuildFacPath(resVault, resID, filename)

    Dim expectedCoords As Long
    expectedCoords = CachedFacNumDims(filePath)
    If expectedCoords <= 0 Then
        ERead_Results = FacPathError(resVault, resID, filePath)
        Exit Function
    End If
    If rowKeys.Columns.Count <> expectedCoords - 1 Then
        ERead_Results = "#ERROR: expected " & CStr(expectedCoords - 1) & " row-key columns, got " & CStr(rowKeys.Columns.Count) & "."
        Exit Function
    End If

    Dim errText As String
    Dim rowsPacked As String
    Dim colsPacked As String
    rowsPacked = BuildPackedRows(rowKeys, False, errText)
    If errText <> "" Then
        ERead_Results = "#ERROR: " & errText
        Exit Function
    End If
    colsPacked = BuildPackedVector(colKeys, False, errText)
    If errText <> "" Then
        ERead_Results = "#ERROR: " & errText
        Exit Function
    End If

    Dim outputPath As String
    outputPath = TempOutputPath()
    Dim status As Long
    status = AppReaderReadBatchDLL(filePath, rowsPacked, colsPacked, outputPath)
    If status <> 0 Then
        If status = 6 Then
            ERead_Results = "#ERROR: file not found: " & filePath & "."
        ElseIf status = 7 Then
            ERead_Results = LoadErrorMessage(filePath)
        Else
            ERead_Results = "#ERROR: " & StatusMessage(status) & " for " & filePath & "."
        End If
        On Error Resume Next
        Kill outputPath
        Exit Function
    End If

    ERead_Results = ReadTsvMatrix(outputPath, rowKeys.Rows.Count, colKeys.Cells.Count)
    On Error Resume Next
    Kill outputPath
    Exit Function
ErrHandler:
    ERead_Results = "#ERROR: invalid ERead_Results arguments."
End Function

' Batch projection lookup. projKeys must have two columns: SP_CODE, VAR_NAME.
Public Function EProj_Results(resVault As String, resID As String, product As String, _
                              projKeys As Range, timePeriods As Range, Optional simID As Variant = "0") As Variant
    On Error GoTo ErrHandler
    Application.Volatile True

    If projKeys.Columns.Count <> 2 Then
        EProj_Results = "#ERROR: EProj_Results expects projKeys with 2 columns: SP_CODE and VAR_NAME."
        Exit Function
    End If

    Dim filePath As String
    filePath = BuildFacPath(resVault, resID, product)

    Dim errText As String
    Dim rowsPacked As String
    Dim colsPacked As String
    rowsPacked = BuildPackedRows(projKeys, False, errText)
    If errText <> "" Then
        EProj_Results = "#ERROR: " & errText
        Exit Function
    End If
    colsPacked = BuildPackedVector(timePeriods, True, errText)
    If errText <> "" Then
        EProj_Results = "#ERROR: " & errText
        Exit Function
    End If

    Dim outputPath As String
    outputPath = TempOutputPath()
    Dim status As Long
    status = AppReaderProjBatchDLL(filePath, CStr(product), CStr(simID), rowsPacked, colsPacked, outputPath)
    If status <> 0 Then
        If status = 6 Then
            EProj_Results = "#ERROR: file not found: " & filePath & "."
        ElseIf status = 7 Then
            EProj_Results = LoadErrorMessage(filePath)
        Else
            EProj_Results = "#ERROR: " & StatusMessage(status) & " for " & filePath & "."
        End If
        On Error Resume Next
        Kill outputPath
        Exit Function
    End If

    EProj_Results = ReadTsvMatrix(outputPath, projKeys.Rows.Count, timePeriods.Cells.Count)
    On Error Resume Next
    Kill outputPath
    Exit Function
ErrHandler:
    EProj_Results = "#ERROR: invalid EProj_Results arguments."
End Function

' Release one cached table after related queries are done.
' Returns TRUE when a cached table was removed, FALSE otherwise.
Public Function FacRelease(resVault As String, resID As String, product As String) As Variant
    On Error GoTo ErrHandler
    FacRelease = (AppReaderReleaseDLL(BuildFacPath(resVault, resID, product)) <> 0)
    Exit Function
ErrHandler:
    FacRelease = CVErr(xlErrValue)
End Function
