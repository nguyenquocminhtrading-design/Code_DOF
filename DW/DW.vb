Option Explicit

'================================================================
' Macro nay TU SINH sheet "003-DW" hoan toan tu sheet "DW".
' KHONG can sheet "004-DW" ton tai nua (thay the hoan toan cho 004-DW).
' Nhap thang dang "2026.07" (dung nhu header dong 4/43 trong DW).
'================================================================

Sub Create003DW()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsDW As Worksheet
    Dim inputMonth As String
    Dim targetCol As Long
    Dim colLetter As String
    Dim found As Range
    Dim i As Long
    Dim items As Variant

    Set wb = ThisWorkbook

    inputMonth = InputBox("Enter the month (e.g. 2026.07):", "Month Selection", "2026.07")
    If inputMonth = "" Then Exit Sub

    On Error Resume Next
    Set wsDW = wb.Sheets("DW")
    On Error GoTo 0
    If wsDW Is Nothing Then
        MsgBox "Khong tim thay sheet 'DW'.", vbCritical
        Exit Sub
    End If

    ' --- Tim cot thang trong dong 4 cua DW ---
    Set found = wsDW.Rows(4).Find(What:=inputMonth, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "Khong tim thay thang '" & inputMonth & "' o dong 4 cua sheet 'DW'.", vbCritical
        Exit Sub
    End If
    targetCol = found.Column
    colLetter = Left(wsDW.Cells(1, targetCol).Address(False, False), Len(wsDW.Cells(1, targetCol).Address(False, False)) - 1)

    ' --- Tao lai sheet 003-DW tu dau ---
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("003-DW").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "003-DW"

    Application.ScreenUpdating = False

    '================================================================
    ' BALANCE SHEET
    '================================================================
    wsNew.Range("A2").Value = "BALANCE SHEET"
    wsNew.Range("B2").Value = inputMonth

    wsNew.Range("A4").Value = "Non-current assets"
    items = Array( _
        Array(5, "Equipment held for operating lease"), _
        Array(6, "Accu. Depreciation"), _
        Array(7, "Investment in joint venture") _
    )
    For i = LBound(items) To UBound(items)
        WriteBS wsNew, colLetter, CLng(items(i)(0)), CStr(items(i)(1))
    Next i

    wsNew.Range("A10").Value = "Current assets"
    items = Array( _
        Array(11, "Prepayment"), _
        Array(12, "Amount due from related parties"), _
        Array(13, "Accounts receivable"), _
        Array(14, "Inventory"), _
        Array(15, "Cash at bank") _
    )
    For i = LBound(items) To UBound(items)
        WriteBS wsNew, colLetter, CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B16").Formula = "=SUM(B4:B15)"

    wsNew.Range("A18").Value = "Current liabilities"
    items = Array( _
        Array(19, "Accounts payable"), _
        Array(20, "Accruals"), _
        Array(21, "Other payable") _
    )
    For i = LBound(items) To UBound(items)
        WriteBS wsNew, colLetter, CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B22").Formula = "=SUM(B17:B21)"
    WriteBS wsNew, colLetter, 23, "Amount due to a director"

    ' Net current assets & NET ASSETS: DW da tinh san 2 dong nay, chi can khop ten
    WriteBS wsNew, colLetter, 26, "Net current assets "
    wsNew.Range("A28").Value = "NET ASSETS"
    WriteBS wsNew, colLetter, 28, "NET ASSETS"

    wsNew.Range("A31").Value = "Equity"
    WriteBS wsNew, colLetter, 32, "Share capital"
    WriteBS wsNew, colLetter, 33, "Retained profits"
    wsNew.Range("B34").Formula = "=SUM(B32:B33)"
    wsNew.Range("A34").Value = ""
    wsNew.Range("A35").Value = "check"
    wsNew.Range("B35").Formula = "=B34-B28"

    '================================================================
    ' PROFIT AND LOSS
    '================================================================
    wsNew.Range("A36").Value = "PROFIT AND LOSS"
    wsNew.Range("B36").Value = inputMonth

    wsNew.Range("A38").Value = "Revenue"
    items = Array( _
        Array(39, "  Sale of goods"), _
        Array(40, "  Commission income"), _
        Array(41, "  Rental income"), _
        Array(42, "  Gain on disposal of fixed asset"), _
        Array(43, "Other income"), _
        Array(44, "  Discount received"), _
        Array(45, "  Financial income"), _
        Array(46, "  Gains on foreign currency exchange"), _
        Array(47, "Others income") _
    )
    For i = LBound(items) To UBound(items)
        WritePL wsNew, colLetter, CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B48").Formula = "=SUM(B39:B47)"

    WritePL wsNew, colLetter, 50, "Less: Cost of good sold"

    wsNew.Range("A51").Value = "Less:  Administrative expenses"
    items = Array( _
        Array(52, "Accounting fee"), _
        Array(53, "Auditors' remuneration"), _
        Array(54, "Bank charges"), _
        Array(55, "Business registration fee"), _
        Array(56, "Sales commission"), _
        Array(57, "Depreciation"), _
        Array(58, "Director's fee"), _
        Array(59, "Incorporation fee"), _
        Array(60, "Legal and professional fee"), _
        Array(61, "Postage and courier"), _
        Array(62, "Sundry"), _
        Array(63, "Other expense -  Losses on foreign currency exchange") _
    )
    For i = LBound(items) To UBound(items)
        WritePL wsNew, colLetter, CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B64").Formula = "=SUM(B52:B63)"

    wsNew.Range("A66").Value = "Profit before taxation"
    wsNew.Range("B66").Formula = "=B48-B64"

    wsNew.Range("A68").Value = "Retained profits"
    wsNew.Range("B68").Formula = "=DW!" & colLetter & "73"

    wsNew.Range("A70").Value = "Dividend paid to DDFH"

    '================================================================
    ' Dinh dang
    '================================================================
    With wsNew
        .Rows("2:4").Font.Bold = True
        .Range("A10,A18,A28,A31,A38,A51").Font.Bold = True
        .Columns("A:B").AutoFit
        .Activate
    End With

    Application.ScreenUpdating = True

    MsgBox "Da tao sheet 003-DW cho thang " & inputMonth & " thanh cong!" & vbCrLf & _
           "Sheet nay KHONG con phu thuoc vao '004-DW' - chi doc truc tiep tu 'DW'.", vbInformation
End Sub

'================================================================
' Helper: ghi 1 dong Balance Sheet, khop ten trong vung DW!A45:A73
'================================================================
Private Sub WriteBS(wsNew As Worksheet, ByVal colLetter As String, ByVal r As Long, ByVal label As String)
    wsNew.Cells(r, "A").Value = label
    wsNew.Cells(r, "B").Formula = _
        "=SUMIFS(DW!" & colLetter & "$45:" & colLetter & "$73,DW!$A$45:$A$73,""" & label & """)"
End Sub

'================================================================
' Helper: ghi 1 dong Profit & Loss, khop ten trong vung DW!A6:A36
'================================================================
Private Sub WritePL(wsNew As Worksheet, ByVal colLetter As String, ByVal r As Long, ByVal label As String)
    wsNew.Cells(r, "A").Value = label
    wsNew.Cells(r, "B").Formula = _
        "=SUMIFS(DW!" & colLetter & "$6:" & colLetter & "$36,DW!$A$6:$A$36,""" & label & """)"
End Sub
