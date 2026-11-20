Option Explicit

' =====================================================================
' DOAU_v2.bas - Creates a 'Summary' sheet matching 100% row-to-row with
' '007-DOAU data' (Rows 1 to 234).
'
' Key features:
' 1. Exact Row-to-Row Alignment:
'    Row r in Summary = Row r in 007-DOAU data.
'    Allows 1-click Copy-Paste Values of D4:D234 directly into 007-DOAU data.
' 2. Dynamic SUMIFS Criteria:
'    Uses column B cell reference (e.g. B7, B8...) instead of hardcoded strings.
' 3. Standardized Subtotal Formulas:
'    Matches the exact ROUND/SUM accounting formulas of 007-DOAU data.
' =====================================================================
Sub CreateSummaryDOAU()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsPL As Worksheet
    Dim i As Long, r As Long
    Dim lastRowBS As Long, lastRowPL As Long
    Dim inputMonth As String
    Dim colAfterBS As String, colAfterPL As String
    Dim helperLevelBS As Long, helperNameBS As Long
    Dim helperLevelPL As Long, helperNamePL As Long
    Dim colNameBS As String, colNamePL As String
    Dim rowsMap() As Variant
    Dim lvl As Long, accountName As String, rowType As String, formulaSpec As String

    Set wb = ThisWorkbook

    inputMonth = InputBox("Enter the month this file covers (e.g. Sep 2026):", "Month Selection", "Sep 2026")
    If inputMonth = "" Then Exit Sub

    On Error Resume Next
    Set wsBS = wb.Sheets("BS")
    Set wsPL = wb.Sheets("PL")
    On Error GoTo 0
    If wsBS Is Nothing Then MsgBox "Sheet 'BS' not found.", vbCritical: Exit Sub
    If wsPL Is Nothing Then MsgBox "Sheet 'PL' not found.", vbCritical: Exit Sub

    colAfterBS = FindColumnByHeader(wsBS, 1, "After")
    colAfterPL = FindColumnByHeader(wsPL, 1, "After")
    If colAfterBS = "" Then MsgBox "Could not find the 'After' column in BS row 1.", vbCritical: Exit Sub
    If colAfterPL = "" Then MsgBox "Could not find the 'After' column in PL row 1.", vbCritical: Exit Sub

    lastRowBS = wsBS.Cells(wsBS.Rows.Count, "A").End(xlUp).Row
    lastRowPL = wsPL.Cells(wsPL.Rows.Count, "A").End(xlUp).Row
    Dim tmpLast As Long
    For i = 2 To 7
        tmpLast = wsBS.Cells(wsBS.Rows.Count, i).End(xlUp).Row
        If tmpLast > lastRowBS Then lastRowBS = tmpLast
        tmpLast = wsPL.Cells(wsPL.Rows.Count, i).End(xlUp).Row
        If tmpLast > lastRowPL Then lastRowPL = tmpLast
    Next i

    helperLevelBS = AddHelperColumns(wsBS, lastRowBS)
    helperLevelPL = AddHelperColumns(wsPL, lastRowPL)
    helperNameBS = helperLevelBS + 1
    helperNamePL = helperLevelPL + 1

    colNameBS = Split(wsBS.Cells(1, helperNameBS).Address, "$")(1)
    colNamePL = Split(wsPL.Cells(1, helperNamePL).Address, "$")(1)

    ' Recreate the Summary sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "Summary"

    ' Initialize row map (1 to 234)
    InitRowsMap rowsMap

    Application.ScreenUpdating = False

    For r = 1 To 234
        lvl = rowsMap(r)(0)
        accountName = rowsMap(r)(1)
        rowType = rowsMap(r)(2)
        formulaSpec = rowsMap(r)(3)

        Select Case rowType
            Case "TITLE"
                wsNew.Cells(r, "A").Value = accountName
                wsNew.Cells(r, "A").Font.Bold = True
            Case "HEADER_MONTH"
                wsNew.Cells(r, "D").Value = inputMonth
                wsNew.Cells(r, "D").Font.Bold = True
                wsNew.Cells(r, "D").HorizontalAlignment = xlRight
            Case "HEADER"
                wsNew.Cells(r, "B").Value = accountName
                wsNew.Cells(r, "B").IndentLevel = lvl - 1
                If lvl <= 2 Then wsNew.Cells(r, "B").Font.Bold = True
            Case "BS_SUMIFS"
                wsNew.Cells(r, "B").Value = accountName
                wsNew.Cells(r, "B").IndentLevel = lvl - 1
                wsNew.Cells(r, "D").Formula = _
                    "=SUMIFS(BS!$" & colAfterBS & "$3:$" & colAfterBS & "$" & lastRowBS & "," & _
                    "BS!$" & colNameBS & "$3:$" & colNameBS & "$" & lastRowBS & ",B" & r & ")"
            Case "PL_SUMIFS"
                wsNew.Cells(r, "B").Value = accountName
                wsNew.Cells(r, "B").IndentLevel = lvl - 1
                wsNew.Cells(r, "D").Formula = _
                    "=SUMIFS(PL!$" & colAfterPL & "$3:$" & colAfterPL & "$" & lastRowPL & "," & _
                    "PL!$" & colNamePL & "$3:$" & colNamePL & "$" & lastRowPL & ",B" & r & ")"
            Case "FORMULA"
                wsNew.Cells(r, "B").Value = accountName
                wsNew.Cells(r, "B").IndentLevel = lvl - 1
                If lvl <= 2 Then wsNew.Cells(r, "B").Font.Bold = True
                wsNew.Cells(r, "D").Formula = formulaSpec
            Case "BLANK"
                ' Leave blank to keep exact row alignment
        End Select
    Next r

    ' Formatting
    With wsNew
        .Columns("A:B").AutoFit
        .Columns("D").ColumnWidth = 18
        .Range("D4:D234").NumberFormat = "#,##0.00;(#,##0.00)"
    End With

    Application.ScreenUpdating = True

    MsgBox "Sheet 'Summary' created successfully matching 007-DOAU data rows 1:1!" & vbCrLf & _
           "You can now copy Summary!D4:D234 directly into 007-DOAU data.", vbInformation
End Sub

' =====================================================================
' Helper: Initialize row specifications (1 to 234)
' Format: Array(Level, Name, RowType, FormulaSpec)
' =====================================================================
Private Sub InitRowsMap(ByRef m() As Variant)
    ReDim m(1 To 234)
    InitRowsMapPart1 m
    InitRowsMapPart2 m
    InitRowsMapPart3 m
End Sub

Private Sub InitRowsMapPart1(ByRef m() As Variant)
    m(1) = Array(1, "FINANCIAL STATEMENT SUMMARY", "TITLE", "")
    m(2) = Array(1, "BALANCE SHEET", "TITLE", "")
    m(3) = Array(0, "Month", "HEADER_MONTH", "")
    m(4) = Array(1, "ASSETS", "HEADER", "")
    m(5) = Array(2, "Current Assets", "HEADER", "")
    m(6) = Array(3, "Checking/Savings", "HEADER", "")
    m(7) = Array(4, "East West Bank -8078001859", "BS_SUMIFS", "")
    m(8) = Array(4, "EWB Payroll 2725", "BS_SUMIFS", "")
    m(9) = Array(4, "Fresno First Bank - 3743", "BS_SUMIFS", "")
    m(10) = Array(4, "Petty Cash", "BS_SUMIFS", "")
    m(11) = Array(3, "Total Checking/Savings", "FORMULA", "=ROUND(SUM(D6:D10),5)")
    m(12) = Array(3, "Accounts Receivable", "HEADER", "")
    m(13) = Array(4, "Accounts Receivable", "BS_SUMIFS", "")
    m(14) = Array(4, "Allowance for Doubtful Accounts", "BS_SUMIFS", "")
    m(15) = Array(3, "Total Accounts Receivable", "FORMULA", "=ROUND(SUM(D12:D14),5)")
    m(16) = Array(3, "Other Current Assets", "HEADER", "")
    m(17) = Array(4, "Inventory Asset", "HEADER", "")
    m(18) = Array(5, "Inventory Clearing", "BS_SUMIFS", "")
    m(19) = Array(5, "Inventory Asset", "BS_SUMIFS", "")
    m(20) = Array(4, "Total Inventory Asset", "FORMULA", "=ROUND(SUM(D17:D19),5)")
    m(21) = Array(4, "Undeposited Funds", "BS_SUMIFS", "")
    m(22) = Array(4, "Pre-paid", "BS_SUMIFS", "")
    m(23) = Array(3, "Total Other Current Assets", "FORMULA", "=ROUND(D16+SUM(D20:D22),5)")
    m(24) = Array(2, "Total Current Assets", "FORMULA", "=ROUND(D5+D11+D15+D23,5)")
    m(25) = Array(2, "Fixed Assets", "HEADER", "")
    m(26) = Array(3, "Accum. Depreciation - Equipment", "BS_SUMIFS", "")
    m(27) = Array(3, "Accum. Depreciation - Furniture", "BS_SUMIFS", "")
    m(28) = Array(3, "Accum. Depreciations", "BS_SUMIFS", "")
    m(29) = Array(3, "Auto", "BS_SUMIFS", "")
    m(30) = Array(3, "Automation", "BS_SUMIFS", "")
    m(31) = Array(3, "Building Improvements - 12368", "BS_SUMIFS", "")
    m(32) = Array(3, "Building Improvements - 2628", "BS_SUMIFS", "")
    m(33) = Array(3, "Computer", "BS_SUMIFS", "")
    m(34) = Array(3, "Construction in Progress", "BS_SUMIFS", "")
    m(35) = Array(3, "Furniture and Equipment", "BS_SUMIFS", "")
    m(36) = Array(3, "Office Equipment", "BS_SUMIFS", "")
    m(37) = Array(3, "Production Equipment", "BS_SUMIFS", "")
    m(38) = Array(2, "Total Fixed Assets", "FORMULA", "=ROUND(SUM(D25:D37),5)")
    m(39) = Array(2, "Other Assets", "HEADER", "")
    m(40) = Array(3, "Deposit", "BS_SUMIFS", "")
    m(41) = Array(3, "Employee Advance", "BS_SUMIFS", "")
    m(42) = Array(2, "Total Other Assets", "FORMULA", "=ROUND(SUM(D39:D41),5)")
    m(43) = Array(1, "TOTAL ASSETS", "FORMULA", "=ROUND(D4+D24+D38+D42,5)")
    m(44) = Array(1, "LIABILITIES & EQUITY", "HEADER", "")
    m(45) = Array(2, "Liabilities", "HEADER", "")
    m(46) = Array(3, "Current Liabilities", "HEADER", "")
    m(47) = Array(4, "Accounts Payable", "BS_SUMIFS", "")
    m(48) = Array(4, "Credit Cards", "HEADER", "")
    m(49) = Array(5, "Home Depot", "BS_SUMIFS", "")
    m(50) = Array(5, "Macy's", "BS_SUMIFS", "")
    m(51) = Array(5, "VISA - 7025", "BS_SUMIFS", "")
    m(52) = Array(4, "Total Credit Cards", "FORMULA", "=ROUND(SUM(D48:D51),5)")
    m(53) = Array(4, "Other Current Liabilities", "HEADER", "")
    m(54) = Array(5, "Accrued Expenses", "BS_SUMIFS", "")
    m(55) = Array(5, "Installment Loan", "HEADER", "")
    m(56) = Array(6, "Electric Forklift 7FBEU20", "BS_SUMIFS", "")
    m(57) = Array(6, "Forklift Battery", "BS_SUMIFS", "")
    m(58) = Array(6, "Toyota", "BS_SUMIFS", "")
    m(59) = Array(5, "Total Installment Loan", "FORMULA", "=ROUND(SUM(D56:D58),5)")
    m(60) = Array(5, "Payroll Liabilities", "HEADER", "")
    m(61) = Array(6, "Payroll Liabilities - Other", "BS_SUMIFS", "")
    m(62) = Array(6, "Payroll Accrual", "BS_SUMIFS", "")
    m(63) = Array(5, "Total Payroll Liabilities", "FORMULA", "=ROUND(SUM(D60:D62),5)")
    m(64) = Array(4, "Total Other Current Liabilities", "FORMULA", "=ROUND(SUM(D53:D54)+D59+D63,5)")
    m(65) = Array(3, "Total Current Liabilities", "FORMULA", "=ROUND(D47+D52+D64,5)")
    m(66) = Array(3, "Long Term Liabilities", "HEADER", "")
    m(67) = Array(4, "Advance from Danway Holdings Lt", "BS_SUMIFS", "")
    m(68) = Array(4, "East West Loan # 332001332", "BS_SUMIFS", "")
    m(69) = Array(4, "East West Loan # 355001220", "BS_SUMIFS", "")
    m(70) = Array(4, "Loan from Shareholders", "BS_SUMIFS", "")
    m(71) = Array(4, "Loan Payable - Alex Do", "BS_SUMIFS", "")
    m(72) = Array(4, "Loan Payable - DDFH", "BS_SUMIFS", "")
    m(73) = Array(4, "Loan Payable - Danways", "BS_SUMIFS", "")
    m(74) = Array(4, "Loan Payable - DDF Hongkong", "BS_SUMIFS", "")
    m(75) = Array(4, "Loan Payable - DDP, LLC", "BS_SUMIFS", "")
    m(76) = Array(4, "Loan Payable - Horseshoe", "BS_SUMIFS", "")
    m(77) = Array(3, "Total Long Term Liabilities", "FORMULA", "=ROUND(SUM(D66:D76),5)")
    m(78) = Array(2, "Total Liabilities", "FORMULA", "=ROUND(+D65+D77,5)")
    m(79) = Array(2, "Equity", "HEADER", "")
    m(80) = Array(3, "Capital Stock", "BS_SUMIFS", "")
    m(81) = Array(3, "Members Equity", "BS_SUMIFS", "")
    m(82) = Array(3, "Retained Earnings", "BS_SUMIFS", "")
    m(83) = Array(3, "Net Income", "BS_SUMIFS", "")
    m(84) = Array(2, "Total Equity", "FORMULA", "=ROUND(SUM(D79:D83),5)")
    m(85) = Array(1, "TOTAL LIABILITIES & EQUITY", "FORMULA", "=ROUND(D46+D78+D84,5)")
    m(86) = Array(1, "Check", "FORMULA", "=D85-D43")
    m(87) = Array(0, "", "BLANK", "")
    m(88) = Array(0, "", "BLANK", "")
End Sub

Private Sub InitRowsMapPart2(ByRef m() As Variant)
    m(89) = Array(1, "PROFIT AND LOSS", "TITLE", "")
    m(90) = Array(1, "Income", "HEADER", "")
    m(91) = Array(2, "REVENUE", "HEADER", "")
    m(92) = Array(3, "Custom Processing", "HEADER", "")
    m(93) = Array(4, "Mixing fee", "PL_SUMIFS", "")
    m(94) = Array(4, "Packing Fees", "PL_SUMIFS", "")
    m(95) = Array(4, "Custom Processing - Other", "PL_SUMIFS", "")
    m(96) = Array(4, "Roasting Fees", "PL_SUMIFS", "")
    m(97) = Array(3, "Total Custom Processing", "FORMULA", "=ROUND(SUM(D92:D96),5)")
    m(98) = Array(3, "Discounts Allowed", "PL_SUMIFS", "")
    m(99) = Array(3, "Other Revenue", "PL_SUMIFS", "")
    m(100) = Array(3, "Returns and Allowances", "PL_SUMIFS", "")
    m(101) = Array(3, "Sales", "PL_SUMIFS", "")
    m(102) = Array(3, "Store Sales", "PL_SUMIFS", "")
    m(103) = Array(2, "Total REVENUE", "FORMULA", "=ROUND(D91+SUM(D97:D102),5)")
    m(104) = Array(1, "Total Income", "FORMULA", "=ROUND(D90+D103,5)")
    m(105) = Array(1, "Cost of Goods Sold", "HEADER", "")
    m(106) = Array(2, "Cost of Goods Sold", "HEADER", "")
    m(107) = Array(3, "Discounts on Purchases", "PL_SUMIFS", "")
    m(108) = Array(3, "Inventory Write Off", "PL_SUMIFS", "")
    m(109) = Array(3, "Physical Inventory Variance", "PL_SUMIFS", "")
    m(110) = Array(3, "Cost of Goods Sold - Other", "PL_SUMIFS", "")
    m(111) = Array(2, "Total Cost of Goods Sold", "FORMULA", "=ROUND(SUM(D106:D110),5)")
    m(112) = Array(2, "Direct Labor", "HEADER", "")
    m(113) = Array(3, "Applied DL", "PL_SUMIFS", "")
    m(114) = Array(3, "Bonus", "PL_SUMIFS", "")
    m(115) = Array(3, "Payroll Taxes", "PL_SUMIFS", "")
    m(116) = Array(3, "Temporary Labor", "PL_SUMIFS", "")
    m(117) = Array(3, "Direct Labor", "PL_SUMIFS", "")
    m(118) = Array(2, "Total Direct Labor", "FORMULA", "=ROUND(SUM(D112:D117),5)")
    m(119) = Array(2, "FACTORY OVERHEAD", "HEADER", "")
    m(120) = Array(3, "Applied OH", "PL_SUMIFS", "")
    m(121) = Array(3, "Bonus", "PL_SUMIFS", "")
    m(122) = Array(3, "Building Repairs", "PL_SUMIFS", "")
    m(123) = Array(3, "Calibration & Certification", "PL_SUMIFS", "")
    m(124) = Array(3, "Cleaning Services", "PL_SUMIFS", "")
    m(125) = Array(3, "Consulting", "PL_SUMIFS", "")
    m(126) = Array(3, "Depreciation Expense", "PL_SUMIFS", "")
    m(127) = Array(3, "Equipment Rental", "PL_SUMIFS", "")
    m(128) = Array(3, "Equipment Repairs", "PL_SUMIFS", "")
    m(129) = Array(3, "Fire Protection", "PL_SUMIFS", "")
    m(130) = Array(3, "Fumigation", "PL_SUMIFS", "")
    m(131) = Array(3, "Janitorial Supplies", "PL_SUMIFS", "")
    m(132) = Array(3, "Lab Fees", "PL_SUMIFS", "")
    m(133) = Array(3, "Lease Payment", "PL_SUMIFS", "")
    m(134) = Array(3, "Licenses & Permits", "PL_SUMIFS", "")
    m(135) = Array(3, "Pallets", "PL_SUMIFS", "")
    m(136) = Array(3, "Payroll Expense", "HEADER", "")
    m(137) = Array(4, "Payroll Taxes Production", "PL_SUMIFS", "")
    m(138) = Array(4, "Payroll Expense - Other", "PL_SUMIFS", "")
    m(139) = Array(4, "Salaries", "PL_SUMIFS", "")
    m(140) = Array(3, "Total Payroll Expense", "FORMULA", "=ROUND(SUM(D136:D139),5)")
    m(141) = Array(3, "Pest Control", "PL_SUMIFS", "")
    m(142) = Array(3, "Plate Charge", "PL_SUMIFS", "")
    m(143) = Array(3, "Production Supplies", "PL_SUMIFS", "")
    m(144) = Array(3, "Rent Expense", "PL_SUMIFS", "")
    m(145) = Array(3, "Repairs and Maintenance", "PL_SUMIFS", "")
    m(146) = Array(3, "Sales Tax/Use Tax", "PL_SUMIFS", "")
    m(147) = Array(3, "Tools & Instruments", "PL_SUMIFS", "")
    m(148) = Array(3, "Training and Seminars", "PL_SUMIFS", "")
    m(149) = Array(3, "Uniforms", "PL_SUMIFS", "")
    m(150) = Array(3, "Utilities", "PL_SUMIFS", "")
    m(151) = Array(3, "Validation Expenses", "PL_SUMIFS", "")
    m(152) = Array(3, "Workers Comp", "PL_SUMIFS", "")
    m(153) = Array(3, "FACTORY OVERHEAD - Other", "PL_SUMIFS", "")
    m(154) = Array(3, "Freight and Shipping Costs", "PL_SUMIFS", "")
    m(155) = Array(2, "Total FACTORY OVERHEAD", "FORMULA", "=ROUND(SUM(D119:D135)+SUM(D140:D154),5)")
    m(156) = Array(1, "Total COGS", "FORMULA", "=ROUND(D105+D111+D118+D155,5)")
    m(157) = Array(1, "Gross Profit", "FORMULA", "=ROUND(D104-D156,5)")
    m(158) = Array(1, "Expense", "HEADER", "")
    m(159) = Array(2, "ADMINISTRATION EXPENSES", "HEADER", "")
    m(160) = Array(3, "Accounting Fees", "PL_SUMIFS", "")
    m(161) = Array(3, "Automobile Expense", "HEADER", "")
    m(162) = Array(4, "Auto Repairs", "PL_SUMIFS", "")
    m(163) = Array(4, "Gasoline, parking & mileage", "PL_SUMIFS", "")
    m(164) = Array(4, "Registration", "PL_SUMIFS", "")
    m(165) = Array(4, "Automobile Expense - Other", "PL_SUMIFS", "")
End Sub

Private Sub InitRowsMapPart3(ByRef m() As Variant)
    m(166) = Array(3, "Total Automobile Expense", "FORMULA", "=ROUND(SUM(D161:D165),5)")
    m(167) = Array(3, "Bad Debts Expense", "PL_SUMIFS", "")
    m(168) = Array(3, "Bank Service Charges", "PL_SUMIFS", "")
    m(169) = Array(3, "Bank Document Fees", "PL_SUMIFS", "")
    m(170) = Array(3, "Business Tax", "PL_SUMIFS", "")
    m(171) = Array(3, "Computer and Internet Expenses", "PL_SUMIFS", "")
    m(172) = Array(3, "Credit Card Service Charge", "PL_SUMIFS", "")
    m(173) = Array(3, "Donation", "PL_SUMIFS", "")
    m(174) = Array(3, "Dues and Subscriptions", "PL_SUMIFS", "")
    m(175) = Array(3, "Insurance Expense", "HEADER", "")
    m(176) = Array(4, "General Liability", "PL_SUMIFS", "")
    m(177) = Array(4, "Property", "PL_SUMIFS", "")
    m(178) = Array(4, "Insurance Expense - Other", "PL_SUMIFS", "")
    m(179) = Array(4, "Vehicle Insurance", "PL_SUMIFS", "")
    m(180) = Array(3, "Total Insurance Expense", "FORMULA", "=ROUND(SUM(D175:D179),5)")
    m(181) = Array(3, "Legal", "PL_SUMIFS", "")
    m(182) = Array(3, "Licenses & Permits", "PL_SUMIFS", "")
    m(183) = Array(3, "Loan Processing Fees", "PL_SUMIFS", "")
    m(184) = Array(3, "Meals and Entertainment", "PL_SUMIFS", "")
    m(185) = Array(3, "Office Expenses", "PL_SUMIFS", "")
    m(186) = Array(3, "Payroll Expenses", "HEADER", "")
    m(187) = Array(4, "Bonus", "PL_SUMIFS", "")
    m(188) = Array(4, "Payroll Processing Fees", "PL_SUMIFS", "")
    m(189) = Array(4, "Payroll Taxes - Admin", "PL_SUMIFS", "")
    m(190) = Array(4, "Salaries - Admin", "PL_SUMIFS", "")
    m(191) = Array(4, "Medical Expenses", "PL_SUMIFS", "")
    m(192) = Array(4, "Employee Benefits", "PL_SUMIFS", "")
    m(193) = Array(4, "Payroll Expenses - Other", "PL_SUMIFS", "")
    m(194) = Array(3, "Total Payroll Expenses", "FORMULA", "=ROUND(SUM(D186:D193),5)")
    m(195) = Array(3, "Penalty & fine", "PL_SUMIFS", "")
    m(196) = Array(3, "Property Taxes Paid", "PL_SUMIFS", "")
    m(197) = Array(3, "Printing & Reproduction", "PL_SUMIFS", "")
    m(198) = Array(3, "Security and Alarm System", "PL_SUMIFS", "")
    m(199) = Array(3, "Suspense", "PL_SUMIFS", "")
    m(200) = Array(3, "Storage Expense", "PL_SUMIFS", "")
    m(201) = Array(3, "State Tax", "PL_SUMIFS", "")
    m(202) = Array(3, "ADMINISTRATION EXPENSES - Other", "PL_SUMIFS", "")
    m(203) = Array(3, "Telephone Expense", "PL_SUMIFS", "")
    m(204) = Array(2, "Total ADMINISTRATION EXPENSES", "FORMULA", "=ROUND(SUM(D159:D160)+SUM(D166:D174)+SUM(D180:D185)+SUM(D194:D203),5)")
    m(205) = Array(2, "SALES EXPENSES", "HEADER", "")
    m(206) = Array(3, "Commissions", "PL_SUMIFS", "")
    m(207) = Array(3, "Domestic Shipping & Handling", "PL_SUMIFS", "")
    m(208) = Array(3, "EDI Fees", "PL_SUMIFS", "")
    m(209) = Array(3, "Export Shipping & Handling", "PL_SUMIFS", "")
    m(210) = Array(3, "Postage and Delivery Expense", "PL_SUMIFS", "")
    m(211) = Array(3, "Service Fee", "PL_SUMIFS", "")
    m(212) = Array(3, "Shipping Supplies", "PL_SUMIFS", "")
    m(213) = Array(3, "Trade Show/Promotion", "PL_SUMIFS", "")
    m(214) = Array(3, "SALES EXPENSES - Other", "PL_SUMIFS", "")
    m(215) = Array(3, "Travel Expense", "PL_SUMIFS", "")
    m(216) = Array(2, "Total SALES EXPENSES", "FORMULA", "=ROUND(SUM(D205:D215),5)")
    m(217) = Array(2, "Taxes", "PL_SUMIFS", "")
    m(218) = Array(1, "Total Expense", "FORMULA", "=ROUND(D158+D204+SUM(D216:D217),5)")
    m(219) = Array(1, "Net Ordinary Income", "FORMULA", "=ROUND(D157-D218,5)")
    m(220) = Array(1, "Other Income/Expense", "HEADER", "")
    m(221) = Array(2, "Other Income", "HEADER", "")
    m(222) = Array(3, "Foreign Currency Gain/Loss", "PL_SUMIFS", "")
    m(223) = Array(3, "Other Income", "PL_SUMIFS", "")
    m(224) = Array(2, "Total Other Income", "FORMULA", "=ROUND(SUM(D221:D223),5)")
    m(225) = Array(2, "Other Expense", "HEADER", "")
    m(226) = Array(3, "Interest Expense", "PL_SUMIFS", "")
    m(227) = Array(3, "Loss(Gain) on disposal of Asset", "PL_SUMIFS", "")
    m(228) = Array(3, "Other Expense (adjustment)", "PL_SUMIFS", "")
    m(229) = Array(2, "Total Other Expense", "FORMULA", "=ROUND(SUM(D225:D228),5)")
    m(230) = Array(1, "Net Other Income", "FORMULA", "=ROUND(D224-D229,5)")
    m(231) = Array(1, "Net operating Income", "FORMULA", "=ROUND(D219+D230,5)")
    m(232) = Array(0, "", "BLANK", "")
    m(233) = Array(1, "Commission/Service Fee - Vietnam", "PL_SUMIFS", "")
    m(234) = Array(1, "Adjusted Net Income", "FORMULA", "=SUM(D231:D233)")
End Sub

' =====================================================================
' Helper: Find column letter by header text
' =====================================================================
Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerText As String) As String
    Dim c As Range
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 20 Then lastCol = 20
    For Each c In ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        If InStr(1, CStr(c.Value), headerText, vbTextCompare) > 0 Then
            FindColumnByHeader = Split(c.Address, "$")(1)
            Exit Function
        End If
    Next c
    FindColumnByHeader = ""
End Function

' =====================================================================
' Helper: Add helper columns (Level, Name Cleaned) to BS and PL
' =====================================================================
Function AddHelperColumns(ws As Worksheet, lastRow As Long) As Long
    Dim colLevel As Long, colName As Long
    Dim r As Long
    Dim q As String
    q = Chr(34) & Chr(34)

    If ws.Cells(1, "P").Value = "Level" And ws.Cells(1, "Q").Value = "Name Cleaned" Then
        AddHelperColumns = ws.Range("P1").Column
        Exit Function
    End If

    colLevel = 16
    colName = 17

    ws.Cells(1, colLevel).Value = "Level"
    ws.Cells(1, colName).Value = "Name Cleaned"
    ws.Range(ws.Cells(1, colLevel), ws.Cells(1, colName)).Font.Bold = True

    Dim formulaLevel As String, formulaName As String
    For r = 3 To lastRow
        formulaLevel = "=IF(LEN(A" & r & ")>0,1,IF(LEN(B" & r & ")>0,2,IF(LEN(C" & r & ")>0,3," & _
                       "IF(LEN(D" & r & ")>0,4,IF(LEN(E" & r & ")>0,5,IF(LEN(F" & r & ")>0,6," & _
                       "IF(LEN(G" & r & ")>0,7," & q & ")))))))"
        ws.Cells(r, colLevel).Formula = formulaLevel

        formulaName = "=TRIM(PROPER(CLEAN(IF(LEN(A" & r & ")>0,A" & r & ",IF(LEN(B" & r & ")>0,B" & r & _
                      ",IF(LEN(C" & r & ")>0,C" & r & ",IF(LEN(D" & r & ")>0,D" & r & _
                      ",IF(LEN(E" & r & ")>0,E" & r & ",IF(LEN(F" & r & ")>0,F" & r & ",G" & r & ")))))))))"
        ws.Cells(r, colName).Formula = formulaName
    Next r

    AddHelperColumns = colLevel
End Function
