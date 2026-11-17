Option Explicit

Sub Create006DDFH()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsIS As Worksheet
    Dim i As Long, r As Long
    Dim accountName As String, code As Variant
    Dim bsItems As Variant, plItems As Variant
    Dim idx As Long
    Dim lastRowBS As Long, lastRowIS As Long
    Dim inputMonth As String
    Dim targetCol As Long
    Dim colNumberBS As Long, colNumberIS As Long
    Dim colLetterBS As String, colLetterIS As String
    Dim shOld As Worksheet
    Dim shLast As Worksheet
    Dim anySheet As Worksheet

    ' ==== LƯU STATE + T?T EVENTS/SCREENUPDATE ====
    Dim prevEvents As Boolean, prevScreen As Boolean, prevAlerts As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail

    Set wb = ThisWorkbook

    ' --- Nh?p tháng ---
    inputMonth = InputBox("Enter the month (e.g. 2026.07):", "Month Selection", "2026.07")
    If inputMonth = "" Then GoTo CleanExit
    If Not IsDate(Replace(inputMonth, ".", "/")) Then
        MsgBox "Invalid date format. Please use YYYY.MM", vbCritical
        GoTo CleanExit
    End If

    ' --- Ki?m tra sheet t?n t?i ---
    On Error Resume Next
    Set wsBS = wb.Sheets("B Sheet")
    Set wsIS = wb.Sheets("IS")
    On Error GoTo CleanFail

    If wsBS Is Nothing Then
        MsgBox "Sheet 'B Sheet' not found.", vbCritical
        GoTo CleanExit
    End If
    If wsIS Is Nothing Then
        MsgBox "Sheet 'IS' not found.", vbCritical
        GoTo CleanExit
    End If

    ' ==== UNGROUP SHEETS (b? multi-selection) ====
    ' C?n thi?t v? Sheets.Add/Delete s? l?i 1004 n?u nhi?u sheet đang đư?c ch?n
    For Each anySheet In wb.Worksheets
        If anySheet.Visible = xlSheetVisible Then
            anySheet.Activate
            Exit For
        End If
    Next anySheet

    ' --- T?o c?t Name Cleaned (c?t G) trong B Sheet ---
    If wsBS.Cells(6, "G").Value <> "Name Cleaned" Then
        wsBS.Columns("G:G").Insert Shift:=xlToRight
        wsBS.Cells(6, "G").Value = "Name Cleaned"
        wsBS.Cells(6, "G").Font.Bold = True
        lastRowBS = wsBS.Cells(wsBS.Rows.Count, "A").End(xlUp).Row
        If lastRowBS < 7 Then lastRowBS = 7
        wsBS.Range("G7:G" & lastRowBS).Formula = "=TRIM(PROPER(CLEAN(A7)))"
        wsBS.Range("G7").AutoFill Destination:=wsBS.Range("G7:G" & lastRowBS)
    End If

    ' --- T?o c?t Code for Number trong B Sheet ---
    colNumberBS = wsBS.Cells(6, wsBS.Columns.Count).End(xlToLeft).Column + 1
    If colNumberBS < 8 Then colNumberBS = 8
    If wsBS.Cells(6, colNumberBS).Value <> "Code for Number" Then
        wsBS.Cells(6, colNumberBS).Value = "Code for Number"
        wsBS.Cells(6, colNumberBS).Font.Bold = True
        lastRowBS = wsBS.Cells(wsBS.Rows.Count, "A").End(xlUp).Row
        If lastRowBS < 7 Then lastRowBS = 7
        wsBS.Range(wsBS.Cells(7, colNumberBS), wsBS.Cells(lastRowBS, colNumberBS)).Formula = "=IF(B7=0,FALSE,ISNUMBER(B7))"
        wsBS.Cells(7, colNumberBS).AutoFill Destination:=wsBS.Range(wsBS.Cells(7, colNumberBS), wsBS.Cells(lastRowBS, colNumberBS))
    End If
    colLetterBS = Split(wsBS.Cells(1, colNumberBS).Address, "$")(1)

    ' --- Tương t? cho sheet IS ---
    If wsIS.Cells(6, "G").Value <> "Name Cleaned" Then
        wsIS.Columns("G:G").Insert Shift:=xlToRight
        wsIS.Cells(6, "G").Value = "Name Cleaned"
        wsIS.Cells(6, "G").Font.Bold = True
        lastRowIS = wsIS.Cells(wsIS.Rows.Count, "A").End(xlUp).Row
        If lastRowIS < 8 Then lastRowIS = 8
        wsIS.Range("G8:G" & lastRowIS).Formula = "=TRIM(PROPER(CLEAN(A8)))"
        wsIS.Range("G8").AutoFill Destination:=wsIS.Range("G8:G" & lastRowIS)
    End If

    colNumberIS = wsIS.Cells(6, wsIS.Columns.Count).End(xlToLeft).Column + 1
    If colNumberIS < 8 Then colNumberIS = 8
    If wsIS.Cells(6, colNumberIS).Value <> "Code for Number" Then
        wsIS.Cells(6, colNumberIS).Value = "Code for Number"
        wsIS.Cells(6, colNumberIS).Font.Bold = True
        lastRowIS = wsIS.Cells(wsIS.Rows.Count, "A").End(xlUp).Row
        If lastRowIS < 8 Then lastRowIS = 8
        wsIS.Range(wsIS.Cells(8, colNumberIS), wsIS.Cells(lastRowIS, colNumberIS)).Formula = "=IF(C8=0,FALSE,ISNUMBER(C8))"
        wsIS.Cells(8, colNumberIS).AutoFill Destination:=wsIS.Range(wsIS.Cells(8, colNumberIS), wsIS.Cells(lastRowIS, colNumberIS))
    End If
    colLetterIS = Split(wsIS.Cells(1, colNumberIS).Address, "$")(1)

    ' --- Xoá sheet c? (n?u t?n t?i) ---
    On Error Resume Next
    Set shOld = wb.Sheets("006-DDFH")
    On Error GoTo CleanFail
    If Not shOld Is Nothing Then
        shOld.Delete
        Set shOld = Nothing
    End If

    ' --- T?o sheet m?i ---
    Set shLast = wb.Sheets(wb.Sheets.Count)
    Set wsNew = wb.Sheets.Add(After:=shLast)
    wsNew.Name = "006-DDFH"

    ' --- Đ?t header tháng + các tiêu đ? chính ---
    targetCol = 3   ' c?t C
    With wsNew
        .Cells(3, targetCol).Value = inputMonth
        .Range("B3").Value = "Month"
        .Range("A1").Value = "BALANCE SHEET"
        .Range("A27").Value = "Total Assets"
        .Range("A50").Value = "PROFIT AND LOSS"
        .Range("A108").Value = "Check"
        .Range("A111").Value = "EBITDA"
        .Range("U7").Value = "Code for Consol"
        .Range("V7").Value = "Name Cleaned"
        .Range("U7:V7").Font.Bold = True
    End With

    ' ================================================================
    ' 1. BALANCE SHEET — 33 items
    ' ================================================================
    ReDim bsItems(1 To 33)
    idx = 1
    bsItems(idx) = Array(5, "ASSETS", ""): idx = idx + 1
    bsItems(idx) = Array(6, "CURRENT ASSETS", ""): idx = idx + 1
    bsItems(idx) = Array(7, "Cash", 111): idx = idx + 1
    bsItems(idx) = Array(8, "Accounts receivable", 131): idx = idx + 1
    bsItems(idx) = Array(9, "Due to related parties", 133): idx = idx + 1
    bsItems(idx) = Array(10, "Due from government agencies", 163): idx = idx + 1
    bsItems(idx) = Array(11, "Due from shareholder", 135): idx = idx + 1
    bsItems(idx) = Array(12, "Inventory", 141): idx = idx + 1
    bsItems(idx) = Array(13, "Investment in securities", 121): idx = idx + 1
    bsItems(idx) = Array(14, "Income taxes recoverable", 163): idx = idx + 1
    bsItems(idx) = Array(15, "Prepaid expenses", 161): idx = idx + 1
    bsItems(idx) = Array(16, "Future income taxes", 163): idx = idx + 1
    ' Rows 17, 18, 19: tr?ng
    bsItems(idx) = Array(20, "SECURITY DEPOSITS", 123): idx = idx + 1
    bsItems(idx) = Array(21, "DUE FROM RELATED PARTIES", 133): idx = idx + 1
    bsItems(idx) = Array(22, "Investment in subsidiaries", 261): idx = idx + 1
    bsItems(idx) = Array(23, "PLANT & EQUIPMENT", 222): idx = idx + 1
    bsItems(idx) = Array(24, "INTANGIBLE ASSETS", 228): idx = idx + 1
    ' Rows 25, 26: tr?ng
    ' Row 27: Total Assets (formula riêng)
    bsItems(idx) = Array(28, "LIABILITIES", ""): idx = idx + 1
    bsItems(idx) = Array(29, "CURRENT LIABILITIES", ""): idx = idx + 1
    bsItems(idx) = Array(30, "Bank indebtedness", 321): idx = idx + 1
    bsItems(idx) = Array(31, "Account payable and accrued liabilities", 311): idx = idx + 1
    bsItems(idx) = Array(32, "Due to related parties", 317): idx = idx + 1
    bsItems(idx) = Array(33, "Due to government agencies", 314): idx = idx + 1
    bsItems(idx) = Array(34, "Income taxes payable", 314): idx = idx + 1
    bsItems(idx) = Array(35, "Demand loans", 321): idx = idx + 1
    bsItems(idx) = Array(36, "Future income taxes", 163): idx = idx + 1
    bsItems(idx) = Array(37, "Due to shareholder", 320): idx = idx + 1
    ' Rows 38, 39: tr?ng
    bsItems(idx) = Array(40, "FUTURE INCOME TAXES", ""): idx = idx + 1
    ' Row 41: tr?ng
    bsItems(idx) = Array(42, "SHAREHOLDER'S EQUITY", ""): idx = idx + 1
    bsItems(idx) = Array(43, "SHARE CAPITAL", "411a"): idx = idx + 1
    bsItems(idx) = Array(44, "CONTRIBUTED SURPLUS", 412): idx = idx + 1
    bsItems(idx) = Array(45, "Dividends", "420a"): idx = idx + 1
    bsItems(idx) = Array(46, "REAINED EARNINGS", "420a"): idx = idx + 1
    ' Rows 47, 48, 49: tr?ng

    ' ================================================================
    ' 2. PROFIT & LOSS — 45 items
    ' ================================================================
    ReDim plItems(1 To 45)
    idx = 1
    plItems(idx) = Array(53, "SALES", ""): idx = idx + 1
    plItems(idx) = Array(54, "SALES", 1): idx = idx + 1
    plItems(idx) = Array(55, "Cashdiscount allowed", ""): idx = idx + 1
    ' Row 56: tr?ng
    plItems(idx) = Array(57, "COST OF SALES", 11): idx = idx + 1
    ' Row 58: tr?ng
    plItems(idx) = Array(59, "GROSS PROFIT", ""): idx = idx + 1
    plItems(idx) = Array(60, "Volume rebates", 1): idx = idx + 1
    ' Row 61: tr?ng
    plItems(idx) = Array(62, "GROSS PROFIT after Rebates", ""): idx = idx + 1
    ' Row 63: tr?ng
    plItems(idx) = Array(64, "EXPENSES", ""): idx = idx + 1
    plItems(idx) = Array(65, "Accounting, legal and professional", 26): idx = idx + 1
    plItems(idx) = Array(66, "Advertising and promotion", 25): idx = idx + 1
    plItems(idx) = Array(67, "  Amortization", 26): idx = idx + 1
    plItems(idx) = Array(68, "Bad debts (recovery)", 26): idx = idx + 1
    plItems(idx) = Array(69, "Consulting fees", 26): idx = idx + 1
    plItems(idx) = Array(70, "Computer services", 26): idx = idx + 1
    plItems(idx) = Array(71, "Entertainment expenses", 26): idx = idx + 1
    plItems(idx) = Array(72, "Equipment rental", 26): idx = idx + 1
    plItems(idx) = Array(73, "Insurance", 26): idx = idx + 1
    plItems(idx) = Array(74, "  Interest and bank charges  Interest and bank charges", 26): idx = idx + 1
    plItems(idx) = Array(75, "Interest on long tearm debt", 23): idx = idx + 1
    plItems(idx) = Array(76, "Management salary", 26): idx = idx + 1
    plItems(idx) = Array(77, "Brokerage fee", 26): idx = idx + 1
    plItems(idx) = Array(78, "Office and miscellaneous", 26): idx = idx + 1
    plItems(idx) = Array(79, "Postage and freight", 25): idx = idx + 1
    plItems(idx) = Array(80, "Premises rent & property tax", 26): idx = idx + 1
    plItems(idx) = Array(81, "Repairs and maintenance", 11): idx = idx + 1
    plItems(idx) = Array(82, "Telephone and utilities", 26): idx = idx + 1
    plItems(idx) = Array(83, "Travel expenses", 26): idx = idx + 1
    plItems(idx) = Array(84, "Truck and automobile upkeep", 25): idx = idx + 1
    plItems(idx) = Array(85, "Wages and benefits", 26): idx = idx + 1
    ' Row 86: tr?ng
    plItems(idx) = Array(87, "OPERATING PROFIT (LOSS)", ""): idx = idx + 1
    ' Row 88: tr?ng
    plItems(idx) = Array(89, "OTHER", ""): idx = idx + 1
    plItems(idx) = Array(90, "Foreign exchange (loss)", 23): idx = idx + 1
    plItems(idx) = Array(91, "Foreign exchange gain", 22): idx = idx + 1
    plItems(idx) = Array(92, "Gain on disposal of equipment", 31): idx = idx + 1
    plItems(idx) = Array(93, "Dividend Income", 22): idx = idx + 1
    plItems(idx) = Array(94, "Commission", 1): idx = idx + 1
    plItems(idx) = Array(95, "Interest Income", 22): idx = idx + 1
    plItems(idx) = Array(96, "Insurance proceeds", 31): idx = idx + 1
    plItems(idx) = Array(97, "Employee Loyalty", 32): idx = idx + 1
    plItems(idx) = Array(98, "Other Income", 31): idx = idx + 1
    ' Row 99: tr?ng
    plItems(idx) = Array(100, "EARNINGS BEFORE INCOME TAXES", ""): idx = idx + 1
    ' Row 101: tr?ng
    plItems(idx) = Array(102, "Income taxes - prior year", 51): idx = idx + 1
    plItems(idx) = Array(103, "Income taxes - current", 51): idx = idx + 1
    plItems(idx) = Array(104, "Future income taxes", 52): idx = idx + 1
    ' Row 105: tr?ng
    plItems(idx) = Array(106, "NET EARNINGS FOR THE PERIOD", ""): idx = idx + 1
    ' Rows 107-110: tr?ng (108 = Check, không set value)
    ' Row 111: EBITDA (formula riêng)

    ' ================================================================
    ' 3. Đi?n Balance Sheet (BS)
    ' ================================================================
    For i = LBound(bsItems) To UBound(bsItems)
        If Not IsArray(bsItems(i)) Then GoTo SkipBS
        r = bsItems(i)(0)
        accountName = bsItems(i)(1)
        code = bsItems(i)(2)
        If r > 0 Then
            wsNew.Cells(r, "B").Value = accountName
            If code <> "" Then wsNew.Cells(r, "U").Value = code
            If accountName <> "" Then
                wsNew.Cells(r, "V").Formula = "=TRIM(PROPER(CLEAN(B" & r & ")))"
                wsNew.Cells(r, targetCol).Formula = "=SUMIFS('B Sheet'!$B$7:$B$43,'B Sheet'!$G$7:$G$43,V" & r & ",'B Sheet'!$" & colLetterBS & "$7:$" & colLetterBS & "$43,TRUE)"
            End If
        End If
SkipBS:
    Next i

    ' D?ng 27: Total Assets
    wsNew.Cells(27, targetCol).Formula = "=SUM(C7:C24)"

    ' ================================================================
    ' 4. Đi?n Profit & Loss (PL)
    ' ================================================================
    For i = LBound(plItems) To UBound(plItems)
        If Not IsArray(plItems(i)) Then GoTo SkipPL
        r = plItems(i)(0)
        accountName = plItems(i)(1)
        code = plItems(i)(2)
        If r > 0 Then
            wsNew.Cells(r, "B").Value = accountName
            If code <> "" Then wsNew.Cells(r, "U").Value = code
            If accountName <> "" Then
                wsNew.Cells(r, "V").Formula = "=TRIM(PROPER(CLEAN(B" & r & ")))"
                wsNew.Cells(r, targetCol).Formula = "=SUMIFS(IS!$C$8:$C$59,IS!$G$8:$G$59,V" & r & ",IS!$" & colLetterIS & "$8:$" & colLetterIS & "$59,TRUE)"
            End If
        End If
SkipPL:
    Next i

    ' ================================================================
    ' 5. Các công th?c đ?c bi?t
    ' ================================================================
    With wsNew
        ' Operating Profit (Loss) — d?ng 87
        .Cells(87, targetCol).Formula = "=C62-SUM(C65:C85)"

        ' EBT — d?ng 100
        .Cells(100, targetCol).Formula = "=C87+SUM(C90:C98)"

        ' Net Earnings — d?ng 106
        .Cells(106, targetCol).Formula = "=C100"

        ' Check — d?ng 108
        .Cells(108, targetCol).Value = ""

        ' EBITDA — d?ng 111
        .Cells(111, targetCol).Formula = "=C106+SUM(C102:C104)-C95-C93+C67"
    End With

    ' ================================================================
    ' 6. Đ?nh d?ng
    ' ================================================================
    With wsNew
        .Rows("1:3").Font.Bold = True
        .Columns("A:V").AutoFit
    End With

    MsgBox "Sheet 006-DDFH created successfully for month " & inputMonth & "!" & vbCrLf & _
           "Columns 'Code for Number' have been auto-created in source sheets.", vbInformation

CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    Exit Sub

CleanFail:
    MsgBox "L?i t?i macro Create006DDFH:" & vbCrLf & vbCrLf & _
           "Mô t?: " & Err.Description & vbCrLf & _
           "M? l?i: " & Err.Number, vbCritical, "Runtime Error"
    Resume CleanExit
End Sub



