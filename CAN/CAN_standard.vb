Option Explicit

'================================================================
' Macro nay TU SINH sheet "003-CAN" hoan toan tu B Sheet / IS / TB.
' KHONG can sheet "004- CAN" ton tai nua (thay the hoan toan cho 004).
'
' Truoc khi sinh bao cao, tu dong refresh 2 cot helper:
'   - B Sheet : G = Name Cleaned | H = Code for Number
'   - IS      : W = Name Cleaned | X = Code for Number
'================================================================

Sub Create003CAN()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsIS As Worksheet, wsTB As Worksheet
    Dim inputMonth As String
    Dim i As Long
    Dim items As Variant

    Set wb = ThisWorkbook

    inputMonth = InputBox("Enter the month (e.g. 2026.07):", "Month Selection", "2026.07")
    If inputMonth = "" Then Exit Sub

    On Error Resume Next
    Set wsBS = wb.Sheets("B Sheet")
    Set wsIS = wb.Sheets("IS")
    Set wsTB = wb.Sheets("TB")
    On Error GoTo 0

    If wsBS Is Nothing Or wsIS Is Nothing Then
        MsgBox "Thieu sheet 'B Sheet' hoac 'IS'.", vbCritical
        Exit Sub
    End If
    If wsTB Is Nothing Then
        MsgBox "Khong tim thay sheet 'TB' - phan EBITDA (Interest add-back) se de trong.", vbExclamation
    End If

    ' >>> TAO/REFRESH 2 COT HELPER TRUOC KHI SINH BAO CAO <<<
    CreateHelperColumns
    ' Sau loi goi nay, B Sheet!G/H va IS!W/X da san sang.

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("003-CAN").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "003-CAN"

    Application.ScreenUpdating = False

    '================================================================
    ' TIEU DE
    '================================================================
    wsNew.Range("A2").Value = "BALANCE SHEET"
    wsNew.Range("C4").Value = inputMonth
    wsNew.Range("A6").Value = "ASSETS"
    wsNew.Range("B7").Value = "CURRENT ASSETS"

    '================================================================
    ' 1. BALANCE SHEET
    '================================================================
    items = Array( _
        Array(8, "Cash"), _
        Array(9, "Accounts receivable"), _
        Array(10, "Due to related parties"), _
        Array(11, "Due from government agencies"), _
        Array(12, "Due from shareholder"), _
        Array(13, "Inventory"), _
        Array(14, "Investment in securities"), _
        Array(15, "Income taxes recoverable"), _
        Array(16, "Prepaid expenses") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "B Sheet", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("C19").Formula = "=SUM(C8:C18)"

    items = Array( _
        Array(21, "SECURITY DEPOSITS"), _
        Array(22, "DUE FROM RELATED PARTIES"), _
        Array(23, "DUE FROM RELATED PARTIES (Loan)"), _
        Array(24, "INVESTMENTS IN SUBSIDIARIES"), _
        Array(25, "Equipment & leasehold improvements"), _
        Array(26, "Intangible assets (net)"), _
        Array(27, "FIXED ASSETS"), _
        Array(28, "ACCUM. DEPRECIATION") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "B Sheet", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("C30").Formula = "=SUM(C19:C29)"

    wsNew.Range("A32").Value = "LIABILITIES"
    wsNew.Range("B33").Value = "CURRENT LIABILITIES"
    WriteLine wsNew, "B Sheet", 33, "CURRENT LIABILITIES"

    items = Array( _
        Array(34, "Bank indebtedness"), _
        Array(35, "Account payable and accrued liabilities"), _
        Array(36, "Accruals"), _
        Array(37, "Due to related parties"), _
        Array(38, "Due to government agencies"), _
        Array(39, "Income taxes payable"), _
        Array(40, "Demand loans"), _
        Array(41, "Long-term loans"), _
        Array(43, "Due to shareholder") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "B Sheet", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("C45").Formula = "=SUM(C34:C43)"

    WriteLine wsNew, "B Sheet", 46, "FUTURE INCOME TAXES"
    wsNew.Range("C47").Formula = "=SUM(C45:C46)"

    wsNew.Range("A48").Value = "SHAREHOLDER'S EQUITY"
    items = Array( _
        Array(49, "SHARE CAPITAL"), _
        Array(50, "CONTRIBUTED SURPLUS"), _
        Array(51, "REAINED EARNINGS") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "B Sheet", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("C52").Formula = "=SUM(C49:C51)"
    wsNew.Range("C54").Formula = "=SUM(C52,C47)"
    wsNew.Range("A55").Value = "Check"
    wsNew.Range("C55").Formula = "=C54-C30"

    '================================================================
    ' 2. PROFIT AND LOSS
    '================================================================
    wsNew.Range("A58").Value = "PROFIT AND LOSS"
    wsNew.Range("C59").Value = inputMonth

    WriteLine wsNew, "IS", 61, "  SALES"
    WriteLine wsNew, "IS", 62, "  Less: Volume rebates"
    WriteLine wsNew, "IS", 63, "  Less:  Cash discount"
    WriteLine wsNew, "IS", 64, "  Less:  Advertising & promotion"
    wsNew.Range("C65").Formula = "=C61-C62-C63-C64"
    wsNew.Range("B65").Value = "  Net Sales"

    WriteLine wsNew, "IS", 67, "COST OF SALES"
    wsNew.Range("C69").Formula = "=C65-C67"
    wsNew.Range("B69").Value = "GROSS PROFIT"

    wsNew.Range("B71").Value = "EXPENSES"
    wsNew.Range("B72").Value = "SALES EXPENSES"
    items = Array( _
        Array(73, "  Listing Fees"), _
        Array(74, "  Advertising and promotion"), _
        Array(75, "  Auto Gas"), _
        Array(76, "  Auto Insurance"), _
        Array(77, "  Auto - R&M"), _
        Array(78, "  Brokerage fee"), _
        Array(79, "  Entertainment"), _
        Array(80, "  Travel expenses"), _
        Array(81, "  Salaries & wages"), _
        Array(82, "  Employees benefits") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "IS", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B83").Value = "  Total Sales Expenses"
    wsNew.Range("C83").Formula = "=SUM(C73:C82)"

    wsNew.Range("B85").Value = "OPERATING EXPENSES"
    items = Array( _
        Array(86, "  Amotization"), _
        Array(87, "  Equipment leases"), _
        Array(88, "  Freight-out"), _
        Array(89, "  Garbage removal"), _
        Array(90, "  Insurance "), _
        Array(91, "  Operating supplies"), _
        Array(92, "  Laundry"), _
        Array(93, "  Real estate rental & Property tax"), _
        Array(94, "  Repairs & Maint - Vehicle"), _
        Array(95, "  Repair & Maint - Bldg upkeep"), _
        Array(96, "  Repair & Maint - Landscaping"), _
        Array(97, "  Repair & Maint - Mach & Equip"), _
        Array(98, "  Truck gas & upkeep"), _
        Array(99, "  Utilities"), _
        Array(100, "  Salaries and wages"), _
        Array(101, "  Employees benefits ") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "IS", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B102").Value = "  Total Operating Expenses"
    wsNew.Range("C102").Formula = "=SUM(C86:C101)"

    wsNew.Range("B104").Value = "  GENERAL & ADMINISTRATION EXPENSES"
    items = Array( _
        Array(105, "  Accounting, legal, and professional"), _
        Array(106, "  Bad debts (recovery)"), _
        Array(107, "  Consulting fees"), _
        Array(108, "  Computer services"), _
        Array(109, "  Interest and bank charges  "), _
        Array(110, "  Insurances"), _
        Array(111, "  Office expense"), _
        Array(112, "  Postage "), _
        Array(113, "  Repairs and maintenance"), _
        Array(114, "  Telephone "), _
        Array(116, "  Vehicle Lease"), _
        Array(117, "  Wages and benefits") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "IS", CLng(items(i)(0)), CStr(items(i)(1))
    Next i
    wsNew.Range("B118").Value = "  Total G&A Expenses"
    wsNew.Range("C118").Formula = "=SUM(C105:C117)"

    wsNew.Range("B120").Value = "  Total  Expenses"
    wsNew.Range("C120").Formula = "=SUM(C118,C102,C83)"

    wsNew.Range("B122").Value = "OPERATING PROFIT/LOSS"
    wsNew.Range("C122").Formula = "=C65-C67-C83-C102-C118"

    wsNew.Range("B124").Value = "OTHER"
    items = Array( _
        Array(125, "Foreign exchange gain/(loss)"), _
        Array(126, "Foreign exchange loss"), _
        Array(127, "Gain on disposal of equipment"), _
        Array(128, "Employee Loyalty & EHT"), _
        Array(129, "Dividend Income"), _
        Array(130, "Commission Income"), _
        Array(131, "Interest Income"), _
        Array(132, "Other Income") _
    )
    For i = LBound(items) To UBound(items)
        WriteLine wsNew, "IS", CLng(items(i)(0)), CStr(items(i)(1))
    Next i

    wsNew.Range("B134").Value = "EARNINGS BEFORE INCOME TAXES"
    wsNew.Range("C134").Formula = "=C122+C125-C126+C127-C128+C129+C130+C131+C132"

    WriteLine wsNew, "IS", 135, "  Income taxes - current"
    WriteLine wsNew, "IS", 136, "  Income taxes - prior year"

    wsNew.Range("B138").Value = "NET EARNINGS FOR THE YEAR"
    wsNew.Range("C138").Formula = "=C134-C135-C136"

    '================================================================
    ' 3. EBITDA - Interest add-back lay truc tiep tu TB (Net Change)
    '================================================================
    If Not wsTB Is Nothing Then
        wsNew.Range("B145").Value = "  Interest - Loan"
        wsNew.Range("C145").Formula = "=SUMIFS(TB!$C:$C,TB!$A:$A,""800210"")"
        wsNew.Range("B146").Value = "  Interest - Other"
        wsNew.Range("C146").Formula = "=SUMIFS(TB!$C:$C,TB!$A:$A,""800220"")"
        wsNew.Range("B147").Value = "  Interest - Overdraft"
        wsNew.Range("C147").Formula = "=SUMIFS(TB!$C:$C,TB!$A:$A,""800230"")"
    End If

    wsNew.Range("A142").Value = "EBITDA"
    wsNew.Range("C142").Formula = "=C138+C86-C131+SUM(C145:C147)-C127+C128"

    '================================================================
    ' 4. Dinh dang
    '================================================================
    With wsNew
        .Rows("2:4").Font.Bold = True
        .Range("B7,A32,B33,B71,B72,B85,B104,A48,B124,A142").Font.Bold = True
        .Columns("A:C").AutoFit
        .Activate
    End With

    Application.ScreenUpdating = True

    MsgBox "Da tao sheet 003-CAN cho thang " & inputMonth & " thanh cong!" & vbCrLf & _
           "Sheet nay KHONG con phu thuoc vao '004- CAN' - chi doc truc tiep tu 'B Sheet', 'IS'" & vbCrLf & _
           "va 'TB' (cho phan EBITDA).", vbInformation
End Sub


'================================================================
' Sub doc lap: tao/refresh 2 cot helper tren B Sheet va IS
'   - B Sheet : G = Name Cleaned | H = Code for Number  (tu dong 7)
'   - IS      : W = Name Cleaned | X = Code for Number  (tu dong 8)
' Duoc goi tu Create003CAN, nhung cung chay rieng duoc.
'================================================================
Sub CreateHelperColumns()
    Dim wb As Workbook
    Dim wsBS As Worksheet, wsIS As Worksheet
    Dim lastRowBS As Long, lastRowIS As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean

    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsBS = wb.Sheets("B Sheet")
    Set wsIS = wb.Sheets("IS")
    On Error GoTo 0

    If wsBS Is Nothing And wsIS Is Nothing Then
        MsgBox "Khong tim thay sheet 'B Sheet' va 'IS'.", vbCritical
        Exit Sub
    End If

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    ' ---------- B Sheet : G, H ----------
    If Not wsBS Is Nothing Then
        lastRowBS = wsBS.Cells(wsBS.Rows.Count, "B").End(xlUp).Row
        If lastRowBS < 7 Then lastRowBS = 7

        wsBS.Range("G6:H" & lastRowBS).ClearContents
        wsBS.Range("G6").Value = "Name Cleaned"
        wsBS.Range("H6").Value = "Code for Number"

        If lastRowBS >= 7 Then
            wsBS.Range("G7:G" & lastRowBS).Formula = "=TRIM(PROPER(CLEAN(B7)))"
            wsBS.Range("H7:H" & lastRowBS).Formula = "=IF(C7=0,FALSE,ISNUMBER(C7))"
        End If
    End If

    ' ---------- IS : W, X ----------
    If Not wsIS Is Nothing Then
        lastRowIS = wsIS.Cells(wsIS.Rows.Count, "B").End(xlUp).Row
        If lastRowIS < 8 Then lastRowIS = 8

        wsIS.Range("W7:X" & lastRowIS).ClearContents
        wsIS.Range("W7").Value = "Name Cleaned"
        wsIS.Range("X7").Value = "Code for Number"

        If lastRowIS >= 8 Then
            wsIS.Range("W8:W" & lastRowIS).Formula = "=TRIM(PROPER(CLEAN(B8)))"
            wsIS.Range("X8:X" & lastRowIS).Formula = "=IF(C8=0,FALSE,ISNUMBER(C8))"
        End If
    End If

    Application.Calculate

Cleanup:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbCritical
End Sub


'================================================================
' Helper: ghi 1 dong line-item vao 003-CAN
'   - B Sheet : SUMIFS tren C7:C44, B7:B44 (Name), H7:H44 (CodeForNumber)
'   - IS      : SUMIFS tren C8:C85, B8:B85 (Name), X8:X85 (CodeForNumber)
'================================================================
Private Sub WriteLine(wsNew As Worksheet, sourceSheet As String, ByVal r As Long, ByVal label As String)
    wsNew.Cells(r, "B").Value = label
    If sourceSheet = "B Sheet" Then
        wsNew.Cells(r, "C").Formula = _
            "=SUMIFS('B Sheet'!$C$7:$C$44,'B Sheet'!$B$7:$B$44,""" & label & """,'B Sheet'!$H$7:$H$44,""TRUE"")"
    ElseIf sourceSheet = "IS" Then
        wsNew.Cells(r, "C").Formula = _
            "=SUMIFS(IS!$C$8:$C$85,IS!$B$8:$B$85,""" & label & """,IS!$X$8:$X$85,""TRUE"")"
    End If
End Sub


