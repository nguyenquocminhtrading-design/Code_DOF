Option Explicit

'============================================================
'  AR / AP ELIMINATION REPORT
'  - Sum range      : FX_adjustment  (fallback: Ending Balance)
'  - Before         : SUMIFS Kind + Entity
'  - Elimination    : SUMIFS + Entity check 2 = "TRUE" AND Cond 1 = "TRUE"
'  - After          : Before(row) - Elimination(row)
'============================================================

Sub GenerateARAPEliminationReport()
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    Dim headerRow As Long
    headerRow = 1

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column

    '--- 1. Map columns ---
    Dim cKind As Long
    Dim cEntity As Long
    Dim cBal As Long
    Dim cCond1 As Long
    Dim cCond2 As Long
    Dim cDateCol As Long
    Dim cSum As Long
    Dim c As Long
    Dim hName As String

    cKind = 0
    cEntity = 0
    cBal = 0
    cCond1 = 0
    cCond2 = 0
    cDateCol = 0
    cSum = 0

    For c = 1 To lastCol
        hName = UCase(Trim(CStr(wsSrc.Cells(headerRow, c).Value)))

        If hName = "KIND" Then
            cKind = c
        ElseIf hName = "ENTITY" Then
            cEntity = c
        ElseIf hName = "ENDING BALANCE" Then
            cBal = c
        ElseIf hName = "COND 1" Then
            cCond1 = c
        ElseIf hName = "ENTITY CHECK 2" Then
            cCond2 = c
        ElseIf hName = "DATE" Or hName = "MONTH" Then
            cDateCol = c
        ElseIf hName = "FX_ADJUSTMENT" Or hName = "FX ADJUSTMENT" Then
            cSum = c
        End If
    Next c

    If cSum = 0 Then cSum = cBal

    If cKind = 0 Or cEntity = 0 Or cSum = 0 Then
        MsgBox "Thieu cot: Kind / Entity / FX_adjustment (hoac Ending Balance).", vbCritical
        Exit Sub
    End If

    '--- 2. Thu thap Date ---
    Dim dictD As Object
    Set dictD = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim dKey As String
    Dim kVal As String

    If cDateCol > 0 Then
        For r = headerRow + 1 To lastRow
            kVal = UCase(Trim(CStr(wsSrc.Cells(r, cKind).Value)))
            If kVal = "AR" Or kVal = "AP" Then
                dKey = CStr(wsSrc.Cells(r, cDateCol).Value)
                If Len(dKey) > 0 Then
                    If Not dictD.Exists(dKey) Then dictD.Add dKey, dKey
                End If
            End If
        Next r
    End If

    Dim dates() As Variant

    If dictD.Count > 0 Then
        dates = dictD.keys

        Dim a As Long
        Dim b As Long
        Dim tmp As Variant
        Dim isBigger As Boolean

        For a = LBound(dates) To UBound(dates) - 1
            For b = a + 1 To UBound(dates)
                If IsDate(dates(a)) And IsDate(dates(b)) Then
                    isBigger = (CDate(dates(a)) > CDate(dates(b)))
                Else
                    isBigger = (CStr(dates(a)) > CStr(dates(b)))
                End If
                If isBigger Then
                    tmp = dates(a)
                    dates(a) = dates(b)
                    dates(b) = tmp
                End If
            Next b
        Next a
    Else
        Dim userInput As String
        Dim defaultVal As String
        defaultVal = "01/07/2026"

        userInput = InputBox( _
            "Khong tim thay cot Date/Month trong du lieu." & vbCrLf & vbCrLf & _
            "Vui long nhap gia tri thang cho bang AR/AP." & vbCrLf & _
            "Vi du: 01/07/2026 hoac 2026-07", _
            "Nhap thang (Date)", defaultVal)

        If Len(Trim(userInput)) = 0 Then
            MsgBox "Da huy.", vbInformation
            Exit Sub
        End If

        ReDim dates(0 To 0)
        dates(0) = Trim(userInput)
    End If

    '--- 3. Entity list ---
    Dim entities() As Variant
    entities = Array("DO", "MP3", "EPF", "ADF", "CAN", "DOAU", "ABC")

    '--- 4. Output sheet ---
    Dim wsOut As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("AR_AP_Elimination")
    If Not wsOut Is Nothing Then wsOut.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsSrc)
    wsOut.Name = "AR_AP_Elimination"

    '--- 5. Build source ranges ---
    Dim srcName As String
    srcName = "'" & Replace(wsSrc.Name, "'", "''") & "'"

    Dim rngKind As String
    Dim rngEntity As String
    Dim rngSum As String
    Dim rngCond1 As String
    Dim rngCond2 As String
    Dim rngDate As String

    rngKind = srcName & "!" & ColAbs(cKind) & "$" & (headerRow + 1) & ":" & ColAbs(cKind) & "$" & lastRow
    rngEntity = srcName & "!" & ColAbs(cEntity) & "$" & (headerRow + 1) & ":" & ColAbs(cEntity) & "$" & lastRow
    rngSum = srcName & "!" & ColAbs(cSum) & "$" & (headerRow + 1) & ":" & ColAbs(cSum) & "$" & lastRow

    If cCond1 > 0 Then
        rngCond1 = srcName & "!" & ColAbs(cCond1) & "$" & (headerRow + 1) & ":" & ColAbs(cCond1) & "$" & lastRow
    End If

    If cCond2 > 0 Then
        rngCond2 = srcName & "!" & ColAbs(cCond2) & "$" & (headerRow + 1) & ":" & ColAbs(cCond2) & "$" & lastRow
    End If

    If cDateCol > 0 Then
        rngDate = srcName & "!" & ColAbs(cDateCol) & "$" & (headerRow + 1) & ":" & ColAbs(cDateCol) & "$" & lastRow
    Else
        rngDate = ""
    End If

    '--- 6. Write 2 blocks ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    WriteSideBlock wsOut, 1, "Total AR", "AR", rngKind, rngEntity, rngSum, _
        rngCond1, rngCond2, rngDate, dates, entities
    WriteSideBlock wsOut, 11, "Total AP", "AP", rngKind, rngEntity, rngSum, _
        rngCond1, rngCond2, rngDate, dates, entities

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '--- 7. Format ---
    wsOut.Columns("A:R").ColumnWidth = 18
    wsOut.Columns("I:I").ColumnWidth = 3
    wsOut.Range("B1:H500").NumberFormat = "#,##0"
    wsOut.Range("L1:R500").NumberFormat = "#,##0"
    wsOut.Activate
    wsOut.Range("A1").Select

    MsgBox "Hoan tat! Xem sheet '" & wsOut.Name & "'.", vbInformation
End Sub


'============================================================
'  WRITE ONE BLOCK (AR or AP)
'============================================================
Private Sub WriteSideBlock(wsOut As Worksheet, startCol As Long, blockTitle As String, kindVal As String, _
    rngKind As String, rngEntity As String, rngSum As String, _
    rngCond1 As String, rngCond2 As String, rngDate As String, _
    dates() As Variant, entities() As Variant)

    Dim r As Long
    Dim c As Long
    Dim s As Long
    Dim i As Long
    Dim j As Long

    r = 1
    c = startCol

    '--- Block title ---
    wsOut.Cells(r, c).Value = blockTitle
    wsOut.Cells(r, c).Font.Bold = True
    wsOut.Cells(r, c).Font.Size = 14
    wsOut.Cells(r, c).Interior.Color = RGB(220, 230, 241)

    For j = 1 To UBound(entities) + 1
        wsOut.Cells(r, c + j).Interior.Color = RGB(220, 230, 241)
    Next j
    r = r + 1

    Dim sections As Variant
    sections = Array("Before elimination", "Elimination value", "After elimination")

    Dim beforeDataStart As Long
    Dim elimDataStart As Long
    beforeDataStart = 0
    elimDataStart = 0

    For s = 0 To UBound(sections)
        '--- Section title ---
        wsOut.Cells(r, c).Value = sections(s)
        wsOut.Cells(r, c).Font.Bold = True
        wsOut.Cells(r, c).Font.Italic = True
        r = r + 1

        '--- Header row ---
        wsOut.Cells(r, c).Value = "Date"
        wsOut.Cells(r, c).Font.Bold = True
        wsOut.Cells(r, c).HorizontalAlignment = xlCenter
        wsOut.Cells(r, c).Interior.Color = RGB(242, 242, 242)

        For j = 0 To UBound(entities)
            wsOut.Cells(r, c + 1 + j).Value = entities(j)
            wsOut.Cells(r, c + 1 + j).Font.Bold = True
            wsOut.Cells(r, c + 1 + j).HorizontalAlignment = xlCenter
            wsOut.Cells(r, c + 1 + j).Interior.Color = RGB(242, 242, 242)
        Next j

        Dim hdrRow As Long
        Dim dataStart As Long
        hdrRow = r
        dataStart = r + 1
        r = r + 1

        '--- Data rows ---
        For i = 0 To UBound(dates)
            wsOut.Cells(r, c).Value = dates(i)

            For j = 0 To UBound(entities)
                Dim entRef As String
                entRef = wsOut.Cells(hdrRow, c + 1 + j).Address(False, False)

                Dim baseStr As String
                baseStr = rngSum & "," & rngKind & ",""" & kindVal & """," & rngEntity & "," & entRef

                If Len(rngDate) > 0 Then
                    baseStr = baseStr & "," & rngDate & "," & wsOut.Cells(r, c).Address(True, False)
                End If

                Dim f As String
                Dim colLetter As String

                If s = 0 Then
                    '--- Before ---
                    f = "=SUMIFS(" & baseStr & ")"

                ElseIf s = 1 Then
                    '--- Elimination ---
                    Dim extra As String
                    extra = ""

                    If Len(rngCond2) > 0 Then
                        extra = extra & "," & rngCond2 & ",""TRUE"""
                    End If

                    If Len(rngCond1) > 0 Then
                        extra = extra & "," & rngCond1 & ",""TRUE"""
                    End If

                    f = "=SUMIFS(" & baseStr & extra & ")"

                Else
                    '--- After = Before(row) - Elimination(row) ---
                    ' FIX: dung .Address mac dinh -> $B$14 -> Split("$B$14","$")(1) = "B"
                    colLetter = Split(wsOut.Cells(r, c + 1 + j).Address, "$")(1)
                    f = "=" & colLetter & (beforeDataStart + i) & "-" & colLetter & (elimDataStart + i)
                End If

                wsOut.Cells(r, c + 1 + j).Formula = f
            Next j
            r = r + 1
        Next i

        '--- Total row ---
        wsOut.Cells(r, c).Value = "Total"
        wsOut.Cells(r, c).Font.Bold = True
        wsOut.Cells(r, c).Interior.Color = RGB(255, 242, 204)

        For j = 0 To UBound(entities)
            Dim colLetter2 As String
            ' FIX tuong tu
            colLetter2 = Split(wsOut.Cells(r, c + 1 + j).Address, "$")(1)
            wsOut.Cells(r, c + 1 + j).Formula = "=SUM(" & colLetter2 & dataStart & ":" & colLetter2 & (r - 1) & ")"
            wsOut.Cells(r, c + 1 + j).Font.Bold = True
            wsOut.Cells(r, c + 1 + j).Interior.Color = RGB(255, 242, 204)
        Next j

        If s = 0 Then beforeDataStart = dataStart
        If s = 1 Then elimDataStart = dataStart

        r = r + 2
    Next s
End Sub


'============================================================
'  Helper: get absolute column address (e.g. $M)
'============================================================
Private Function ColAbs(ByVal colIdx As Long) As String
    Dim parts() As String
    parts = Split(Cells(1, colIdx).Address(True, True), "$")
    ColAbs = "$" & parts(1)
End Function

