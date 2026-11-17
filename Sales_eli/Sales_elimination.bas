Option Explicit

'============================================================
'  XUAT BANG REVENUE & COGS THEO ENTITY / MONTH
'  - 3 block con: Before elimination / Elimination / After elimination
'  - Entity TU DONG gom tu cot Entity trong data (khong hard-code)
'  - DOAU = tong cac cot Entity (cot cuoi cung)
'============================================================

Sub GenerateRevenueCOGSReport()
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    ' Chan truong hop dang dung o sheet output
    If wsSrc.Name = "Revenue_COGS_Summary" Then
        MsgBox "Hay chon sheet du lieu nguon truoc khi chay macro.", vbExclamation
        Exit Sub
    End If

    '=== 1. Header ===
    Dim headerRow As Long
    headerRow = 1

    Dim lastRow As Long
    Dim lastCol As Long
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column

    '=== 2. Map cac cot can dung ===
    Dim cEntity As Long
    Dim cMonth As Long
    Dim cNetSales As Long
    Dim cTotalCost As Long
    Dim cCond As Long
    Dim c As Long
    Dim hName As String

    cEntity = 0
    cMonth = 0
    cNetSales = 0
    cTotalCost = 0
    cCond = 0

    For c = 1 To lastCol
        hName = UCase(Trim(CStr(wsSrc.Cells(headerRow, c).Value)))

        If hName = "ENTITY" Then
            cEntity = c
        ElseIf hName = "MONTH" Then
            cMonth = c
        ElseIf hName = "NET SALES" Then
            cNetSales = c
        ElseIf hName = "TOTAL COST" Then
            cTotalCost = c
        ElseIf hName = "COND" Then
            cCond = c
        End If
    Next c

    If cEntity = 0 Or cMonth = 0 Or cNetSales = 0 Or cTotalCost = 0 Or cCond = 0 Then
        MsgBox "Thieu cot bat buoc (Entity / Month / Net Sales / Total Cost / Cond).", vbCritical
        Exit Sub
    End If

    '--- lastRow = dong cuoi lon nhat trong cac cot can dung ---
    Dim colsUsed As Variant
    Dim k As Long
    Dim tmpLast As Long

    colsUsed = Array(cEntity, cMonth, cNetSales, cTotalCost, cCond)
    lastRow = headerRow

    For k = 0 To UBound(colsUsed)
        tmpLast = wsSrc.Cells(wsSrc.Rows.Count, colsUsed(k)).End(xlUp).Row
        If tmpLast > lastRow Then lastRow = tmpLast
    Next k

    If lastRow <= headerRow Then
        MsgBox "Khong co dong du lieu.", vbCritical
        Exit Sub
    End If

    '=== 3. Gom danh sach thang (unique + sort tang dan) ===
    Dim dictM As Object
    Set dictM = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim mKey As String

    For r = headerRow + 1 To lastRow
        mKey = CStr(wsSrc.Cells(r, cMonth).Value)
        If Len(mKey) > 0 Then
            If Not dictM.Exists(mKey) Then dictM.Add mKey, mKey
        End If
    Next r

    If dictM.Count = 0 Then
        MsgBox "Cot Month khong co du lieu.", vbCritical
        Exit Sub
    End If

    Dim months() As Variant
    months = dictM.keys

    Dim a As Long
    Dim b As Long
    Dim t As Variant
    Dim isBigger As Boolean

    For a = LBound(months) To UBound(months) - 1
        For b = a + 1 To UBound(months)
            If IsDate(months(a)) And IsDate(months(b)) Then
                isBigger = (CDate(months(a)) > CDate(months(b)))
            Else
                isBigger = (CStr(months(a)) > CStr(months(b)))
            End If
            If isBigger Then
                t = months(a)
                months(a) = months(b)
                months(b) = t
            End If
        Next b
    Next a

    '=== 4. TU DONG GOM ENTITY TU COT Entity (bo DOAU) ===
    Dim dictE As Object
    Set dictE = CreateObject("Scripting.Dictionary")

    Dim eKey As String
    Dim eRaw As String
    Dim badSpaceCount As Long
    badSpaceCount = 0

    For r = headerRow + 1 To lastRow
        eRaw = CStr(wsSrc.Cells(r, cEntity).Value)
        eKey = UCase(Trim(eRaw))

        ' Dem so o co khoang trang thua (SUMIFS se khong khop)
        If Len(eKey) > 0 Then
            If UCase(eRaw) <> eKey Then badSpaceCount = badSpaceCount + 1
        End If

        If Len(eKey) > 0 And eKey <> "DOAU" Then
            If Not dictE.Exists(eKey) Then dictE.Add eKey, eKey
        End If
    Next r

    If dictE.Count = 0 Then
        MsgBox "Cot Entity khong co du lieu hop le.", vbCritical
        Exit Sub
    End If

    If badSpaceCount > 0 Then
        If MsgBox("Co " & badSpaceCount & " o trong cot Entity bi thua khoang trang dau/cuoi." & vbCrLf & _
                  "SUMIFS se KHONG khop cac dong nay (so lieu bi thieu)." & vbCrLf & vbCrLf & _
                  "Nen dung Trim cot Entity truoc roi chay lai." & vbCrLf & _
                  "Ban co muon tiep tuc chay khong?", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    Dim entities() As Variant
    entities = dictE.keys

    ' Sort A -> Z
    Dim ae As Long
    Dim be As Long
    Dim te As Variant

    For ae = LBound(entities) To UBound(entities) - 1
        For be = ae + 1 To UBound(entities)
            If CStr(entities(ae)) > CStr(entities(be)) Then
                te = entities(ae)
                entities(ae) = entities(be)
                entities(be) = te
            End If
        Next be
    Next ae

    '=== 5. Tao sheet output ===
    Dim wb As Workbook
    Set wb = wsSrc.Parent

    Dim wsOut As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOut = wb.Worksheets("Revenue_COGS_Summary")
    If Not wsOut Is Nothing Then wsOut.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = wb.Worksheets.Add(After:=wsSrc)
    wsOut.Name = "Revenue_COGS_Summary"

    '=== 6. Range nguon ===
    Dim srcName As String
    srcName = "'" & Replace(wsSrc.Name, "'", "''") & "'"

    Dim rMonth As String
    Dim rEntity As String
    Dim rCond As String
    Dim rNet As String
    Dim rCost As String

    rMonth = srcName & "!" & ColAbs(cMonth) & "$" & (headerRow + 1) & ":" & ColAbs(cMonth) & "$" & lastRow
    rEntity = srcName & "!" & ColAbs(cEntity) & "$" & (headerRow + 1) & ":" & ColAbs(cEntity) & "$" & lastRow
    rCond = srcName & "!" & ColAbs(cCond) & "$" & (headerRow + 1) & ":" & ColAbs(cCond) & "$" & lastRow
    rNet = srcName & "!" & ColAbs(cNetSales) & "$" & (headerRow + 1) & ":" & ColAbs(cNetSales) & "$" & lastRow
    rCost = srcName & "!" & ColAbs(cTotalCost) & "$" & (headerRow + 1) & ":" & ColAbs(cTotalCost) & "$" & lastRow

    '=== 7. Xuat 2 block ===
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim outRow As Long
    outRow = 1

    outRow = WriteBlock(wsOut, outRow, "Revenue", rNet, rMonth, rEntity, rCond, months, entities)
    outRow = outRow + 2
    outRow = WriteBlock(wsOut, outRow, "COGS", rCost, rMonth, rEntity, rCond, months, entities)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    On Error GoTo 0

    '=== 8. Format ===
    wsOut.Columns("A:Z").ColumnWidth = 16
    wsOut.Range("B1:Z" & outRow).NumberFormat = "#,##0"
    wsOut.Activate
    wsOut.Range("A1").Select

    MsgBox "Hoan tat! Xem sheet '" & wsOut.Name & "'.", vbInformation
    Exit Sub

EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbCritical
End Sub


'============================================================
'  XUAT 1 BLOCK (Revenue hoac COGS)
'============================================================
Private Function WriteBlock(wsOut As Worksheet, startRow As Long, blockTitle As String, _
    rAmt As String, rMonth As String, rEntity As String, rCond As String, _
    months() As Variant, entities() As Variant) As Long

    Dim r As Long
    Dim s As Long
    Dim i As Long
    Dim j As Long
    Dim doauCol As Long
    Dim hdrRow As Long
    Dim f As String
    Dim entRef As String

    r = startRow

    wsOut.Cells(r, 1).Value = blockTitle
    wsOut.Cells(r, 1).Font.Bold = True
    wsOut.Cells(r, 1).Font.Size = 14
    r = r + 1

    Dim sections As Variant
    sections = Array("Before elimination", "Elimination value", "After elimination")

    doauCol = UBound(entities) + 2

    For s = 0 To UBound(sections)
        '--- Section title ---
        wsOut.Cells(r, 1).Value = sections(s)
        wsOut.Cells(r, 1).Font.Bold = True
        wsOut.Cells(r, 1).Font.Italic = True
        r = r + 1

        '--- Header dong ---
        wsOut.Cells(r, 1).Value = "Month"
        wsOut.Cells(r, 1).Font.Bold = True
        wsOut.Cells(r, 1).HorizontalAlignment = xlCenter

        For j = 0 To UBound(entities)
            wsOut.Cells(r, j + 2).Value = entities(j)
            wsOut.Cells(r, j + 2).Font.Bold = True
            wsOut.Cells(r, j + 2).HorizontalAlignment = xlCenter
        Next j

        '--- DOAU cot cuoi ---
        wsOut.Cells(r, doauCol).Value = "DOAU"
        wsOut.Cells(r, doauCol).Font.Bold = True
        wsOut.Cells(r, doauCol).HorizontalAlignment = xlCenter

        hdrRow = r
        r = r + 1

        '--- Tung thang ---
        For i = 0 To UBound(months)
            wsOut.Cells(r, 1).Value = months(i)

            For j = 0 To UBound(entities)
                entRef = wsOut.Cells(hdrRow, j + 2).Address(True, False)

                If s = 0 Then
                    '--- Before elimination ---
                    f = "=SUMIFS(" & rAmt & "," & rMonth & ",$A" & r & _
                        "," & rEntity & "," & entRef & ")"

                ElseIf s = 1 Then
                    '--- Elimination: Cond = TRUE ---
                    f = "=SUMIFS(" & rAmt & "," & rMonth & ",$A" & r & _
                        "," & rEntity & "," & entRef & _
                        "," & rCond & ",""TRUE"")"

                Else
                    '--- After: Cond <> TRUE ---
                    f = "=SUMIFS(" & rAmt & "," & rMonth & ",$A" & r & _
                        "," & rEntity & "," & entRef & _
                        "," & rCond & ",""<>TRUE"")"
                End If

                wsOut.Cells(r, j + 2).Formula = f
            Next j

            '--- DOAU = SUM cac cot entity (B -> truoc DOAU) ---
            wsOut.Cells(r, doauCol).Formula = _
                "=SUM(" & wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, doauCol - 1)).Address(False, False) & ")"

            r = r + 1
        Next i

        r = r + 1   ' dong trong giua cac section
    Next s

    WriteBlock = r
End Function


'============================================================
'  Helper: lay dia chi cot tuyet doi (vd: $M)
'============================================================
Private Function ColAbs(ByVal colIdx As Long) As String
    Dim parts() As String
    parts = Split(Cells(1, colIdx).Address(True, True), "$")
    ColAbs = "$" & parts(1)
End Function

