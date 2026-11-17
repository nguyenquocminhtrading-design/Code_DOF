Option Explicit

Private Const FUZZY_THRESHOLD As Double = 85

Private Function SourceSheets() As Variant
    SourceSheets = Array("DOF", "ADF", "DOAU", "CAN")
End Function

Private Function CanonicalFields() As Variant
    CanonicalFields = Array("Entity", "Name", "Item No", "Description", _
        "Qty", "UOM", "Net Sales", "Total Cost", "GM", "Month", "Group", "IC check")
End Function

Private Function SynonymMap(ByVal canonicalField As String) As Variant
    Select Case canonicalField
        Case "Entity":       SynonymMap = Array("Entity", "Entity2")
        Case "Name":         SynonymMap = Array("Name")
        Case "Item No":      SynonymMap = Array("Item No.", "Item No", "Item 1")
        Case "Description":  SynonymMap = Array("Description", "Item 2", "Memo")
        Case "Qty":          SynonymMap = Array("Invoiced Qty.", "Inv Qty", "Qty")
        Case "UOM":          SynonymMap = Array("Base UOM", "U/M", "Pack Size")
        Case "Net Sales":    SynonymMap = Array("Net Sales", "Amount")
        Case "Total Cost":   SynonymMap = Array("Total Cost")
        Case "GM":           SynonymMap = Array("GM", "GP$")
        Case "Month":        SynonymMap = Array("Month")
        Case "Group":        SynonymMap = Array("Group")
        Case "IC check":     SynonymMap = Array("IC check")
        Case Else:           SynonymMap = Array()
    End Select
End Function

'==========================================================================
' MAIN ENTRY POINT
'==========================================================================
Sub RunConsolidation()

    Dim wb As Workbook
    Dim wsOut As Worksheet, wsLog As Worksheet
    Dim shtName As Variant
    Dim canonFields As Variant
    Dim outCol As Long
    Dim startTime As Double
    startTime = Timer

    Set wb = ThisWorkbook

    ' --- Speed switches: turn off everything that slows bulk writes ----
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Consolidating sales data..."

    On Error GoTo CleanFail

    Set wsOut = ResetSheet(wb, "Consolidated")
    Set wsLog = ResetSheet(wb, "Mapping Log")

    canonFields = CanonicalFields()
    For outCol = LBound(canonFields) To UBound(canonFields)
        wsOut.Cells(1, outCol + 1).Value = canonFields(outCol)
    Next outCol
    wsOut.Cells(1, UBound(canonFields) + 2).Value = "Source Sheet"
    FormatHeaderRow wsOut, UBound(canonFields) + 2

    wsLog.Range("A1:F1").Value = Array("Source Sheet", "Canonical Field", _
        "Matched Source Column", "Match Type", "Confidence %", "Column Letter")
    FormatHeaderRow wsLog, 6

    Dim logRow As Long: logRow = 2
    Dim outRow As Long: outRow = 2

    For Each shtName In SourceSheets()
        If SheetExists(wb, CStr(shtName)) Then
            Application.StatusBar = "Processing sheet: " & shtName & " ..."
            ProcessSheet wb, CStr(shtName), wsOut, wsLog, logRow, outRow
        Else
            Debug.Print "Sheet not found, skipped: " & shtName
        End If
    Next shtName

    wsOut.Columns.AutoFit
    wsLog.Columns.AutoFit
    wsOut.Activate

CleanExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Consolidation complete." & vbCrLf & _
           "Rows written: " & (outRow - 2) & vbCrLf & _
           "Time: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf & _
           "See 'Consolidated' sheet for data and 'Mapping Log' for the column-mapping audit trail.", _
           vbInformation, "Done"
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbCritical, "Consolidation failed"
    Resume CleanExit

End Sub

'==========================================================================
' Process a single source sheet using bulk array read/write (fast)
'==========================================================================
Private Sub ProcessSheet(ByVal wb As Workbook, ByVal shtName As String, _
                          ByVal wsOut As Worksheet, ByVal wsLog As Worksheet, _
                          ByRef logRow As Long, ByRef outRow As Long)

    Dim ws As Worksheet
    Set ws = wb.Sheets(shtName)

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then
        Debug.Print "No header row found on sheet: " & shtName
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    Dim numDataRows As Long
    numDataRows = lastRow - headerRow

    ' --- Build column map (cheap: only reads the header row, ~30 cells) --
    Dim canonFields As Variant: canonFields = CanonicalFields()
    Dim numCanon As Long: numCanon = UBound(canonFields) - LBound(canonFields) + 1

    Dim colMap() As Long
    ReDim colMap(LBound(canonFields) To UBound(canonFields))

    Dim f As Long
    For f = LBound(canonFields) To UBound(canonFields)
        colMap(f) = MapColumn(ws, headerRow, lastCol, CStr(canonFields(f)), _
                               shtName, wsLog, logRow)
    Next f

    ' --- BULK READ: whole data block into memory in ONE call ------------
    Dim srcArr As Variant
    If numDataRows = 1 Then
        ' Value2 on a single-row range returns a scalar/1-col array edge case;
        ' force a proper 2D array by reading a (1 To 1, 1 To lastCol) range.
        Dim tmp As Variant
        tmp = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(headerRow + 1, lastCol)).Value2
        ReDim srcArr(1 To 1, 1 To lastCol)
        Dim cc As Long
        For cc = 1 To lastCol
            srcArr(1, cc) = tmp(1, cc)
        Next cc
    Else
        srcArr = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, lastCol)).Value2
    End If

    ' --- Build output array IN MEMORY (no worksheet calls in this loop) --
    Dim outArr As Variant
    ReDim outArr(1 To numDataRows, 1 To numCanon + 1)

    Dim writtenRows As Long: writtenRows = 0
    Dim r As Long, c As Long
    Dim rowHasData As Boolean

    For r = 1 To numDataRows
        rowHasData = False
        For c = 1 To lastCol
            If Not IsEmpty(srcArr(r, c)) Then
                If Not (VarType(srcArr(r, c)) = vbString And srcArr(r, c) = "") Then
                    rowHasData = True
                    Exit For
                End If
            End If
        Next c

        If rowHasData Then
            writtenRows = writtenRows + 1
            For c = LBound(canonFields) To UBound(canonFields)
                If colMap(c) > 0 Then
                    outArr(writtenRows, c + 1) = srcArr(r, colMap(c))
                End If
            Next c
            outArr(writtenRows, numCanon + 1) = shtName
        End If
    Next r

    If writtenRows = 0 Then Exit Sub

    ' --- Trim array to actual written size if any blank rows were skipped
    Dim finalArr As Variant
    If writtenRows = numDataRows Then
        finalArr = outArr
    Else
        ReDim finalArr(1 To writtenRows, 1 To numCanon + 1)
        Dim rr As Long
        For rr = 1 To writtenRows
            For c = 1 To numCanon + 1
                finalArr(rr, c) = outArr(rr, c)
            Next c
        Next rr
    End If

    ' --- BULK WRITE: whole block to the Consolidated sheet in ONE call --
    wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow + writtenRows - 1, numCanon + 1)).Value2 = finalArr

    outRow = outRow + writtenRows

End Sub

'==========================================================================
' Find the source column index for a canonical field (unchanged logic)
'==========================================================================
Private Function MapColumn(ByVal ws As Worksheet, ByVal headerRow As Long, _
                            ByVal lastCol As Long, ByVal canonicalField As String, _
                            ByVal shtName As String, ByVal wsLog As Worksheet, _
                            ByRef logRow As Long) As Long

    Dim c As Long
    Dim headerText As String
    Dim synonyms As Variant
    Dim s As Long

    synonyms = SynonymMap(canonicalField)

    For c = 1 To lastCol
        headerText = CleanHeader(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then
            For s = LBound(synonyms) To UBound(synonyms)
                If StrComp(headerText, CleanHeader(CStr(synonyms(s))), vbTextCompare) = 0 Then
                    LogMapping wsLog, logRow, shtName, canonicalField, _
                        ws.Cells(headerRow, c).Value, "Exact/Synonym", 100, c
                    MapColumn = c
                    Exit Function
                End If
            Next s
        End If
    Next c

    Dim bestCol As Long: bestCol = 0
    Dim bestScore As Double: bestScore = 0
    Dim bestHeader As String

    For c = 1 To lastCol
        headerText = CleanHeader(ws.Cells(headerRow, c).Value)
        If Len(headerText) > 0 Then
            Dim score As Double
            score = SimilarityPct(headerText, CleanHeader(canonicalField))
            If score > bestScore Then
                bestScore = score: bestCol = c: bestHeader = ws.Cells(headerRow, c).Value
            End If
            For s = LBound(synonyms) To UBound(synonyms)
                score = SimilarityPct(headerText, CleanHeader(CStr(synonyms(s))))
                If score > bestScore Then
                    bestScore = score: bestCol = c: bestHeader = ws.Cells(headerRow, c).Value
                End If
            Next s
        End If
    Next c

    If bestScore >= FUZZY_THRESHOLD Then
        LogMapping wsLog, logRow, shtName, canonicalField, bestHeader, "Fuzzy", bestScore, bestCol
        MapColumn = bestCol
    Else
        LogMapping wsLog, logRow, shtName, canonicalField, "(not found)", "None", 0, 0
        MapColumn = 0
    End If

End Function

Private Sub LogMapping(ByVal wsLog As Worksheet, ByRef logRow As Long, _
                        ByVal shtName As String, ByVal canonicalField As String, _
                        ByVal matchedHeader As String, ByVal matchType As String, _
                        ByVal confidence As Double, ByVal colIdx As Long)
    wsLog.Cells(logRow, 1).Value = shtName
    wsLog.Cells(logRow, 2).Value = canonicalField
    wsLog.Cells(logRow, 3).Value = matchedHeader
    wsLog.Cells(logRow, 4).Value = matchType
    wsLog.Cells(logRow, 5).Value = Round(confidence, 1)
    If colIdx > 0 Then
        wsLog.Cells(logRow, 6).Value = ColLetter(colIdx)
    Else
        wsLog.Cells(logRow, 6).Value = ""
    End If
    If matchType = "None" Then
        wsLog.Range(wsLog.Cells(logRow, 1), wsLog.Cells(logRow, 6)).Interior.Color = RGB(255, 199, 206)
    End If
    logRow = logRow + 1
End Sub

Private Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long
    Dim lastCheckCol As Long
    lastCheckCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCheckCol < 40 Then lastCheckCol = 40

    For r = 1 To 5
        For c = 1 To lastCheckCol
            If StrComp(CleanHeader(ws.Cells(r, c).Value), "entity", vbTextCompare) = 0 Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

Private Function CleanHeader(ByVal v As Variant) As String
    Dim s As String
    s = Trim(CStr(v & ""))
    Do While Right(s, 1) = "."
        s = Left(s, Len(s) - 1)
    Loop
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanHeader = Trim(s)
End Function

Private Function SimilarityPct(ByVal s1 As String, ByVal s2 As String) As Double
    Dim maxLen As Long
    s1 = LCase(s1): s2 = LCase(s2)
    maxLen = Application.WorksheetFunction.Max(Len(s1), Len(s2))
    If maxLen = 0 Then
        SimilarityPct = 100
    Else
        SimilarityPct = (1 - (LevenshteinDistance(s1, s2) / maxLen)) * 100
    End If
End Function

Private Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Long
    Dim m As Long, n As Long, i As Long, j As Long, cost As Long
    m = Len(s): n = Len(t)

    Dim d() As Long
    ReDim d(0 To m, 0 To n)

    For i = 0 To m: d(i, 0) = i: Next i
    For j = 0 To n: d(0, j) = j: Next j

    For i = 1 To m
        For j = 1 To n
            If Mid(s, i, 1) = Mid(t, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            Dim v1 As Long, v2 As Long, v3 As Long
            v1 = d(i - 1, j) + 1
            v2 = d(i, j - 1) + 1
            v3 = d(i - 1, j - 1) + cost
            If v1 <= v2 And v1 <= v3 Then
                d(i, j) = v1
            ElseIf v2 <= v3 Then
                d(i, j) = v2
            Else
                d(i, j) = v3
            End If
        Next j
    Next i

    LevenshteinDistance = d(m, n)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function ResetSheet(ByVal wb As Workbook, ByVal shtName As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    If SheetExists(wb, shtName) Then
        wb.Sheets(shtName).Delete
    End If
    Application.DisplayAlerts = True
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = shtName
    Set ResetSheet = ws
End Function

Private Sub FormatHeaderRow(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    On Error Resume Next
    ws.Rows(1).AutoFilter
    On Error GoTo 0
End Sub

Private Function ColLetter(ByVal colIndex As Long) As String
    ColLetter = Split(Cells(1, colIndex).Address(True, False), "$")(0)
End Function
