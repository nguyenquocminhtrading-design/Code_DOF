Option Explicit

'================================================================
' 003-CAN — Logic chuyen tu 006-DDFH:
'   - Auto tao helper columns (B Sheet: G/H | IS: W/X)
'   - Cot U = Code for Consol | V = Name Cleaned tren output
'   - SUMIFS criteria = V{r} (thay vi $B{r})
'   - Range dong (dynamic theo lastRow)
'================================================================
Sub Create003CAN()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsIS As Worksheet, wsTB As Worksheet
    Dim i As Long, r As Long
    Dim inputMonth As String
    Dim items As Variant
    Dim lastRowBS As Long, lastRowIS As Long
    Dim nameBS As String, rangeBS As String
    Dim nameIS As String, rangeIS As String
    Dim colLetterBS As String

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
        MsgBox "Khong tim thay 'TB' - phan EBITDA se de trong.", vbExclamation
    End If

    ' ================================================================
    ' >>> TANG 1: AUTO TAO HELPER COLUMNS
    '     B Sheet : G = Name Cleaned | H = Code for Number
    '     IS      : W = Name Cleaned | X = Code for Number
    ' ================================================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    ' --- B Sheet : G = Name Cleaned, H = Code for Number ---
    lastRowBS = wsBS.Cells(wsBS.Rows.Count, "B").End(xlUp).Row
    If lastRowBS < 7 Then lastRowBS = 7

    wsBS.Range("G6").Value = "Name Cleaned"
    wsBS.Range("H6").Value = "Code for Number"
    wsBS.Range("G6:H6").Font.Bold = True

    If lastRowBS >= 7 Then
        wsBS.Range("G7:G" & lastRowBS).Formula = "=TRIM(PROPER(CLEAN(B7)))"
        wsBS.Range("H7:H" & lastRowBS).Formula = "=IF(C7=0,FALSE,ISNUMBER(C7))"
    End If
    colLetterBS = "H"

    ' --- IS : W = Name Cleaned, X = Code for Number ---
    lastRowIS = wsIS.Cells(wsIS.Rows.Count, "B").End(xlUp).Row
    If lastRowIS < 8 Then lastRowIS = 8

    wsIS.Range("W7").Value = "Name Cleaned"
    wsIS.Range("X7").Value = "Code for Number"
    wsIS.Range("W7:X7").Font.Bold = True

    If lastRowIS >= 8 Then
        wsIS.Range("W8:W" & lastRowIS).Formula = "=TRIM(PROPER(CLEAN(B8)))"
        wsIS.Range("X8:X" & lastRowIS).Formula = "=IF(C8=0,FALSE,ISNUMBER(C8))"
    End If

    ' --- Chuan bi chuoi range dong (1 lan, dung cho toan bo SUMIFS) ---
    nameBS = "'B Sheet'!$B$7:$B$" & lastRowBS
    rangeBS = "'B Sheet'!$" & colLetterBS & "$7:$" & colLetterBS & "$" & lastRowBS
    nameIS = "IS!$W$8:$W$" & lastRowIS
    rangeIS = "IS!$X$8:$X$" & lastRowIS    ' co dinh X cho IS

    Application.Calculate

    ' ================================================================
    ' >>> TANG 2: TAO SHEET OUTPUT
    ' ================================================================
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("003-CAN").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "003-CAN"

    ' Tieu de
    wsNew.Range("A2").Value = "BALANCE SHEET"
    wsNew.Range("C4").Value = inputMonth
    wsNew.Range("A6").Value = "ASSETS"
    wsNew.Range("B7").Value = "CURRENT ASSETS"

    ' Header U, V (giong 006-DDFH)
    wsNew.Range("U6").Value = "Code for Consol"
    wsNew.Range("V6").Value = "Name Cleaned"
    wsNew.Range("U6:V6").Font.Bold = True

    ' ================================================================
    ' >>> TANG 3: GHI TUNG DONG LINE-ITEM (SUMIFS dung V{r})
    ' ================================================================

    ' --- 3.1 BALANCE SHEET ---
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
        WriteLineBS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameBS, rangeBS
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
        WriteLineBS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameBS, rangeBS
    Next i
    wsNew.Range("C30").Formula = "=SUM(C19:C29)"

    wsNew.Range("A32").Value = "LIABILITIES"
    wsNew.Range("B33").Value = "CURRENT LIABILITIES"
    WriteLineBS wsNew, 33, "CURRENT LIABILITIES", nameBS, rangeBS

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
        WriteLineBS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameBS, rangeBS
    Next i
    wsNew.Range("C45").Formula = "=SUM(C34:C43)"

    WriteLineBS wsNew, 46, "FUTURE INCOME TAXES", nameBS, rangeBS
    wsNew.Range("C47").Formula = "=SUM(C45:C46)"

    wsNew.Range("A48").Value = "SHAREHOLDER'S EQUITY"
    items = Array( _
        Array(49, "SHARE CAPITAL"), _
        Array(50, "CONTRIBUTED SURPLUS"), _
        Array(51, "REAINED EARNINGS") _
    )
    For i = LBound(items) To UBound(items)
        WriteLineBS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameBS, rangeBS
    Next i
    wsNew.Range("C52").Formula = "=SUM(C49:C51)"
    wsNew.Range("C54").Formula = "=SUM(C52,C47)"
    wsNew.Range("A55").Value = "Check"
    wsNew.Range("C55").Formula = "=C54-C30"

    ' --- 3.2 PROFIT AND LOSS ---
    wsNew.Range("A58").Value = "PROFIT AND LOSS"
    wsNew.Range("C59").Value = inputMonth

    WriteLineIS wsNew, 61, "  SALES", nameIS, rangeIS
    WriteLineIS wsNew, 62, "  Less: Volume rebates", nameIS, rangeIS
    WriteLineIS wsNew, 63, "  Less:  Cash discount", nameIS, rangeIS
    WriteLineIS wsNew, 64, "  Less:  Advertising & promotion", nameIS, rangeIS
    wsNew.Range("C65").Formula = "=C61-C62-C63-C64"
    wsNew.Range("B65").Value = "  Net Sales"

    WriteLineIS wsNew, 67, "COST OF SALES", nameIS, rangeIS
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
        WriteLineIS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameIS, rangeIS
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
        WriteLineIS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameIS, rangeIS
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
        WriteLineIS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameIS, rangeIS
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
        WriteLineIS wsNew, CLng(items(i)(0)), CStr(items(i)(1)), nameIS, rangeIS
    Next i

    wsNew.Range("B134").Value = "EARNINGS BEFORE INCOME TAXES"
    wsNew.Range("C134").Formula = "=C122+C125-C126+C127-C128+C129+C130+C131+C132"

    WriteLineIS wsNew, 135, "  Income taxes - current", nameIS, rangeIS
    WriteLineIS wsNew, 136, "  Income taxes - prior year", nameIS, rangeIS

    wsNew.Range("B138").Value = "NET EARNINGS FOR THE YEAR"
    wsNew.Range("C138").Formula = "=C134-C135-C136"

    ' --- 3.3 EBITDA ---
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

    ' ================================================================
    ' >>> TANG 4: DINH DANG + KET THUC
    ' ================================================================
    With wsNew
        .Rows("2:4").Font.Bold = True
        .Range("B7,A32,B33,B71,B72,B85,B104,A48,B124,A142").Font.Bold = True
        .Columns("A:C").AutoFit
        .Activate
    End With

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Loi: " & Err.Description, vbCritical
        Exit Sub
    End If

    MsgBox "Da tao sheet 003-CAN cho thang " & inputMonth & "!" & vbCrLf & _
           "Helper columns: B Sheet!G/H | IS!W/X" & vbCrLf & _
           "Dynamic range: BS toi dong " & lastRowBS & " | IS toi dong " & lastRowIS, _
           vbInformation
End Sub


'================================================================
' Helper: ghi 1 dong tu B Sheet
'   - B = label
'   - V = TRIM(PROPER(CLEAN(B{r})))   <-- criteria cho SUMIFS
'   - C = SUMIFS('B Sheet'!C, 'B Sheet'!B=V{r}, CodeForNumber=TRUE)
'================================================================
Private Sub WriteLineBS(wsNew As Worksheet, ByVal r As Long, ByVal label As String, _
                        ByVal nameBS As String, ByVal rangeBS As String)
    Dim lastR As Long

    wsNew.Cells(r, "B").Value = label
    If Len(Trim(label)) > 0 Then
        wsNew.Cells(r, "V").Formula = "=TRIM(PROPER(CLEAN(B" & r & ")))"

        ' Lay dong cuoi tu nameBS (dang "'B Sheet'!$B$7:$B$123")
        lastR = CLng(Mid(nameBS, InStrRev(nameBS, "$") + 1))

        With wsNew.Cells(r, "C")
            .NumberFormat = "General"
            .Formula = "=SUMIFS('B Sheet'!$C$7:$C$" & lastR & _
                       "," & nameBS & ",V" & r & _
                       "," & rangeBS & ",TRUE)"
        End With
    End If
End Sub


'================================================================
' Helper: ghi 1 dong tu IS
'   - B = label
'   - V = TRIM(PROPER(CLEAN(B{r})))
'   - C = SUMIFS(IS!C, IS!B=V{r}, IS!X=TRUE)   <-- co dinh cot X
'================================================================
Private Sub WriteLineIS(wsNew As Worksheet, ByVal r As Long, ByVal label As String, _
                        ByVal nameIS As String, ByVal rangeIS As String)
    Dim lastR As Long

    wsNew.Cells(r, "B").Value = label
    If Len(Trim(label)) > 0 Then
        wsNew.Cells(r, "V").Formula = "=TRIM(PROPER(CLEAN(B" & r & ")))"

        ' Lay dong cuoi tu nameIS (dang "IS!$B$8:$B$123")
        lastR = CLng(Mid(nameIS, InStrRev(nameIS, "$") + 1))

        With wsNew.Cells(r, "C")
            .NumberFormat = "General"
            .Formula = "=SUMIFS(IS!$C$8:$C$" & lastR & _
                       "," & nameIS & ",V" & r & _
                       "," & rangeIS & ",TRUE)"
        End With
    End If
End Sub

