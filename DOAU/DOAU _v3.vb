Option Explicit

' =====================================================================
' DOAU_v3.vb - Creates a 'Summary' sheet matching 100% row-to-row with
' 'DOAU' sheet (Rows 1 to 237, ending at Adjusted Net Income).
'
' Key features:
' 1. Exact Row-to-Row Alignment  : Row r in Summary = Row r in DOAU sheet.
' 2. Dynamic SUMIFS Criteria    : Uses B-column cell reference (B7, B8...).
' 3. Standardized Subtotal       : ROUND/SUM formulas match DOAU sheet.
'
' Source: FS DOAU Sep 26 - Adjusted.xlsm  »  Sheet "DOAU"
' Rows included: 1–237  (stop at Adjusted Net Income)
' Rows EXCLUDED: 238 Check, 239–247 (Per ch? Khánh, EBITDA, etc.)
' =====================================================================

Sub CreateSummaryDOAU()

    Dim wb          As Workbook
    Dim wsNew       As Worksheet, wsBS As Worksheet, wsPL As Worksheet
    Dim i           As Long, r As Long, tmpLast As Long
    Dim lastRowBS   As Long, lastRowPL As Long
    Dim inputMonth  As String
    Dim colAfterBS  As String, colAfterPL As String
    Dim helperLevelBS As Long, helperNameBS As Long
    Dim helperLevelPL As Long, helperNamePL As Long
    Dim colNameBS   As String, colNamePL As String
    Dim rowsMap()   As Variant
    Dim lvl         As Long
    Dim accountName As String, rowType As String, formulaSpec As String

    On Error GoTo ErrorHandler

    Set wb = ThisWorkbook

    inputMonth = InputBox("Enter the month this file covers (e.g. Sep 2026):", _
                          "Month Selection", "Sep 2026")
    If inputMonth = "" Then Exit Sub

    ' ---- Locate BS / PL ----
    On Error Resume Next
    Set wsBS = wb.Sheets("BS")
    Set wsPL = wb.Sheets("PL")
    On Error GoTo ErrorHandler

    If wsBS Is Nothing Then
        MsgBox "Sheet 'BS' not found.", vbCritical
        Exit Sub
    End If
    If wsPL Is Nothing Then
        MsgBox "Sheet 'PL' not found.", vbCritical
        Exit Sub
    End If

    ' ---- Find 'After' columns ----
    colAfterBS = FindColumnByHeader(wsBS, 1, "After")
    colAfterPL = FindColumnByHeader(wsPL, 1, "After")

    If colAfterBS = "" Then
        MsgBox "Could not find the 'After' column in BS row 1.", vbCritical
        Exit Sub
    End If
    If colAfterPL = "" Then
        MsgBox "Could not find the 'After' column in PL row 1.", vbCritical
        Exit Sub
    End If

    ' ---- Last used row (scan columns A..G) ----
    lastRowBS = wsBS.Cells(wsBS.Rows.Count, "A").End(xlUp).Row
    lastRowPL = wsPL.Cells(wsPL.Rows.Count, "A").End(xlUp).Row
    For i = 2 To 7
        tmpLast = wsBS.Cells(wsBS.Rows.Count, i).End(xlUp).Row
        If tmpLast > lastRowBS Then lastRowBS = tmpLast
        tmpLast = wsPL.Cells(wsPL.Rows.Count, i).End(xlUp).Row
        If tmpLast > lastRowPL Then lastRowPL = tmpLast
    Next i

    ' ---- Helper columns ----
    helperLevelBS = AddHelperColumns(wsBS, lastRowBS)
    helperLevelPL = AddHelperColumns(wsPL, lastRowPL)
    helperNameBS = helperLevelBS + 1
    helperNamePL = helperLevelPL + 1

    colNameBS = Split(wsBS.Cells(1, helperNameBS).Address, "$")(1)
    colNamePL = Split(wsPL.Cells(1, helperNamePL).Address, "$")(1)

    ' ---- Recreate Summary sheet ----
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Summary").Delete
    Err.Clear
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "Summary"

    ' ---- Row map 1..237 (matching DOAU sheet row-for-row) ----
    InitRowsMap rowsMap

    Application.ScreenUpdating = False

    For r = 1 To 237
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

    ' ---- Formatting ----
    With wsNew
        .Columns("A:B").AutoFit
        .Columns("D").ColumnWidth = 18
        .Range("D4:D237").NumberFormat = "#,##0.00;(#,##0.00)"
    End With

    Application.ScreenUpdating = True

    MsgBox "Sheet 'Summary' created successfully matching DOAU sheet rows 1:1!" & vbCrLf & _
           "You can now copy Summary!D4:D237 directly into DOAU sheet.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


' =====================================================================
' Helper: Initialize row specifications (1 to 237)
' Array(Level, Name, RowType, FormulaSpec)
' Source: DOAU sheet in FS DOAU Sep 26 - Adjusted.xlsm
' =====================================================================
Private Sub InitRowsMap(ByRef m() As Variant)
    ReDim m(1 To 237)
    InitRowsMapPart1 m
    InitRowsMapPart2 m
    InitRowsMapPart3 m
End Sub

' =====================================================================
' Part 1: Rows 1–89  (BALANCE SHEET — from FX title to Check)
' =====================================================================
Private Sub InitRowsMapPart1(ByRef m() As Variant)
    ' --- Row 1-3: Headers ---
    m(1) = Array(1, "FX", "TITLE", "")
    m(2) = Array(1, "BALANCE SHEET", "TITLE", "")
    m(3) = Array(0, "Month", "HEADER_MONTH", "")

    ' --- ASSETS section ---
    m(4) = Array(1, "ASSETS", "HEADER", "")
    m(5) = Array(2, "Current Assets", "HEADER", "")
    m(6) = Array(3, "Checking/Savings", "HEADER", "")
    m(7) = Array(4, "East West Bank -8078001859", "BS_SUMIFS", "")
    m(8) = Array(4, "EWB Payroll 2725", "BS_SUMIFS", "")
    m(9) = Array(4, "Fresno First Bank - 3743", "BS_SUMIFS", "")
    m(10) = Array(4, "Petty Cash", "BS_SUMIFS", "")
    m(11) = Array(3, "Total Checking/Savings", "FORMULA", "=ROUND(SUM(D6:D10),5)")
    m(12) = Array(3, "Accounts Receivable", "HEADER", "")
    m(13) = Array(4, "Accounts Receivable - Other", "BS_SUMIFS", "")
    m(14) = Array(4, "Allowance for doubtful debts", "BS_SUMIFS", "")
    m(15) = Array(3, "Total Accounts Receivable", "FORMULA", "=ROUND(SUM(D12:D14),5)")
    m(16) = Array(3, "Other Current Assets", "HEADER", "")
    m(17) = Array(4, "Inventory Asset", "HEADER", "")
    m(18) = Array(5, "Inventory Clearing", "BS_SUMIFS", "")
    m(19) = Array(5, "Work-in-Process", "BS_SUMIFS", "")
    m(20) = Array(5, "Inventory Asset - Other", "BS_SUMIFS", "")
    m(21) = Array(4, "Total Inventory Asset", "FORMULA", "=ROUND(SUM(D17:D20),5)")
    m(22) = Array(4, "Undeposited Funds", "BS_SUMIFS", "")
    m(23) = Array(4, "Pre-paid", "BS_SUMIFS", "")
    m(24) = Array(3, "Total Other Current Assets", "FORMULA", "=ROUND(D16+SUM(D21:D23),5)")
    m(25) = Array(2, "Total Current Assets", "FORMULA", "=ROUND(D5+D11+D15+D24,5)")
    m(26) = Array(2, "Fixed Assets", "HEADER", "")
    m(27) = Array(3, "Accum. Depreciation - Equipment", "BS_SUMIFS", "")
    m(28) = Array(3, "Accum. Depreciation - Furniture", "BS_SUMIFS", "")
    m(29) = Array(3, "Accum. Depreciations", "BS_SUMIFS", "")
    m(30) = Array(3, "Auto", "BS_SUMIFS", "")
    m(31) = Array(3, "Automation", "BS_SUMIFS", "")
    m(32) = Array(3, "Building Improvements - 12368", "BS_SUMIFS", "")
    m(33) = Array(3, "Building Improvements - 2628", "BS_SUMIFS", "")
    m(34) = Array(3, "Computer", "BS_SUMIFS", "")
    m(35) = Array(3, "Construction in Progress", "BS_SUMIFS", "")
    m(36) = Array(3, "Furniture and Equipment", "BS_SUMIFS", "")
    m(37) = Array(3, "Office Equipment", "BS_SUMIFS", "")
    m(38) = Array(3, "Production Equipment", "BS_SUMIFS", "")
    m(39) = Array(2, "Total Fixed Assets", "FORMULA", "=ROUND(SUM(D26:D38),5)")
    m(40) = Array(2, "Other Assets", "HEADER", "")
    m(41) = Array(3, "Deposit", "BS_SUMIFS", "")
    m(42) = Array(3, "Employee Advance", "BS_SUMIFS", "")
    m(43) = Array(2, "Total Other Assets", "FORMULA", "=ROUND(SUM(D40:D42),5)")
    m(44) = Array(1, "TOTAL ASSETS", "FORMULA", "=ROUND(D4+D25+D39+D43,5)")

    ' --- LIABILITIES & EQUITY section ---
    m(45) = Array(1, "LIABILITIES & EQUITY", "HEADER", "")
    m(46) = Array(2, "Liabilities", "HEADER", "")
    m(47) = Array(3, "Current Liabilities", "HEADER", "")
    m(48) = Array(4, "Accounts Payable", "BS_SUMIFS", "")
    m(49) = Array(4, "Credit Cards", "HEADER", "")
    m(50) = Array(5, "Home Depot", "BS_SUMIFS", "")
    m(51) = Array(5, "Macy's", "BS_SUMIFS", "")
    m(52) = Array(5, "Visa - 6188", "BS_SUMIFS", "")
    m(53) = Array(5, "VISA - 3240", "BS_SUMIFS", "")
    m(54) = Array(4, "Total Credit Cards", "FORMULA", "=ROUND(SUM(D49:D53),5)")
    m(55) = Array(4, "Other Current Liabilities", "HEADER", "")
    m(56) = Array(5, "Accrued Expenses", "BS_SUMIFS", "")
    m(57) = Array(5, "Installment Loan", "HEADER", "")
    m(58) = Array(6, "Electric Forklift", "BS_SUMIFS", "")
    m(59) = Array(6, "Loan East West Bank -8078001859", "BS_SUMIFS", "")
    m(60) = Array(6, "Toyota", "BS_SUMIFS", "")
    m(61) = Array(5, "Total Installment Loan", "FORMULA", "=ROUND(SUM(D58:D60),5)")
    m(62) = Array(5, "Payroll Liabilities", "HEADER", "")
    m(63) = Array(6, "Payroll Liabilities - Other", "BS_SUMIFS", "")
    m(64) = Array(6, "Payroll Accrual", "BS_SUMIFS", "")
    m(65) = Array(5, "Total Payroll Liabilities", "FORMULA", "=ROUND(SUM(D62:D64),5)")
    m(66) = Array(4, "Total Other Current Liabilities", "FORMULA", "=D56+D61+D65")
    m(67) = Array(3, "Total Current Liabilities", "FORMULA", "=D48+D54+D66")
    m(68) = Array(3, "Long Term Liabilities", "HEADER", "")
    m(69) = Array(4, "Advance from Danway Holdings Lt", "BS_SUMIFS", "")
    m(70) = Array(4, "East West Loan # 332001332", "BS_SUMIFS", "")
    m(71) = Array(4, "East West Loan # 355001220", "BS_SUMIFS", "")
    m(72) = Array(4, "Loan from Shareholders", "BS_SUMIFS", "")
    m(73) = Array(4, "Loan Payable - Alex Do", "BS_SUMIFS", "")
    m(74) = Array(4, "Loan Payable - Dan-D-Foods- Ltd", "BS_SUMIFS", "")
    m(75) = Array(4, "Loan Payable - Danway", "BS_SUMIFS", "")
    m(76) = Array(4, "Loan Payable - DDF Hongkong", "BS_SUMIFS", "")
    m(77) = Array(4, "Loan Payable - DDP, LLC", "BS_SUMIFS", "")
    m(78) = Array(4, "Loan Payable - Horseshoe", "BS_SUMIFS", "")
    m(79) = Array(4, "Loan Payable- Dan D Pistachio", "BS_SUMIFS", "")
    m(80) = Array(3, "Total Long Term Liabilities", "FORMULA", "=SUM(D69:D79)")
    m(81) = Array(2, "Total Liabilities", "FORMULA", "=ROUND(+D67+D80,5)")
    m(82) = Array(2, "Equity", "HEADER", "")
    m(83) = Array(3, "Capital Stock", "BS_SUMIFS", "")
    m(84) = Array(3, "Members Equity", "BS_SUMIFS", "")
    m(85) = Array(3, "Retained Earnings", "BS_SUMIFS", "")
    m(86) = Array(3, "Net Income", "BS_SUMIFS", "")
    m(87) = Array(2, "Total Equity", "FORMULA", "=ROUND(SUM(D82:D86),5)")
    m(88) = Array(1, "TOTAL LIABILITIES & EQUITY", "FORMULA", "=ROUND(D47+D81+D87,5)")
    m(89) = Array(1, "Check", "FORMULA", "=D88-D44")
    m(90) = Array(0, "", "BLANK", "")
    m(91) = Array(0, "", "BLANK", "")
End Sub

' =====================================================================
' Part 2: Rows 92–165  (PROFIT AND LOSS — Income through FACTORY OVERHEAD)
' =====================================================================
Private Sub InitRowsMapPart2(ByRef m() As Variant)
    ' --- PROFIT AND LOSS ---
    m(92) = Array(1, "PROFIT AND LOSS", "TITLE", "")
    m(93) = Array(1, "Income", "HEADER", "")
    m(94) = Array(2, "REVENUE", "HEADER", "")
    m(95) = Array(3, "Custom Processing", "HEADER", "")
    m(96) = Array(4, "Mixing fee", "PL_SUMIFS", "")
    m(97) = Array(4, "Packing Fees", "PL_SUMIFS", "")
    m(98) = Array(4, "Roasting Fees", "PL_SUMIFS", "")
    m(99) = Array(4, "Custom Processing - Other", "PL_SUMIFS", "")
    m(100) = Array(3, "Total Custom Processing", "FORMULA", "=ROUND(SUM(D95:D99),5)")
    m(101) = Array(3, "Discounts Allowed", "PL_SUMIFS", "")
    m(102) = Array(3, "Other Revenue", "PL_SUMIFS", "")
    m(103) = Array(3, "Returns and Allowances", "PL_SUMIFS", "")
    m(104) = Array(3, "Sales", "PL_SUMIFS", "")
    m(105) = Array(3, "Store Sales", "PL_SUMIFS", "")
    m(106) = Array(2, "Total REVENUE", "FORMULA", "=ROUND(D94+SUM(D100:D105),5)")
    m(107) = Array(1, "Total Income", "FORMULA", "=ROUND(D93+D106,5)")
    m(108) = Array(1, "Cost Of Goods Sold", "HEADER", "")
    m(109) = Array(2, "Cost of Goods Sold", "HEADER", "")
    m(110) = Array(3, "Discounts on Purchases", "PL_SUMIFS", "")
    m(111) = Array(3, "Inventory Write Off", "PL_SUMIFS", "")
    m(112) = Array(3, "Physical Inventory Variance", "PL_SUMIFS", "")
    m(113) = Array(3, "Cost Of Goods Sold", "PL_SUMIFS", "")
    m(114) = Array(2, "Total Cost of Goods Sold", "FORMULA", "=ROUND(SUM(D109:D113),5)")
    m(115) = Array(2, "Direct Labor", "HEADER", "")
    m(116) = Array(3, "Applied DL", "PL_SUMIFS", "")
    m(117) = Array(3, "Bonus", "PL_SUMIFS", "")
    m(118) = Array(3, "Payroll Taxes", "PL_SUMIFS", "")
    m(119) = Array(3, "Temporary Labor", "PL_SUMIFS", "")
    m(120) = Array(3, "Direct Labor - Other", "PL_SUMIFS", "")
    m(121) = Array(2, "Total Direct Labor", "FORMULA", "=ROUND(SUM(D115:D120),5)")
    m(122) = Array(2, "FACTORY OVERHEAD", "HEADER", "")
    m(123) = Array(3, "Applied OH", "PL_SUMIFS", "")
    m(124) = Array(3, "Bonus", "PL_SUMIFS", "")
    m(125) = Array(3, "Building Repairs", "PL_SUMIFS", "")
    m(126) = Array(3, "Calibration & Certification", "PL_SUMIFS", "")
    m(127) = Array(3, "Cleaning Services", "PL_SUMIFS", "")
    m(128) = Array(3, "Consulting", "PL_SUMIFS", "")
    m(129) = Array(3, "Depreciation Expense", "PL_SUMIFS", "")
    m(130) = Array(3, "Equipment Rental", "PL_SUMIFS", "")
    m(131) = Array(3, "Equipment Repairs", "PL_SUMIFS", "")
    m(132) = Array(3, "Fire Protection", "PL_SUMIFS", "")
    m(133) = Array(3, "Fumigation", "PL_SUMIFS", "")
    m(134) = Array(3, "Janitorial Supplies", "PL_SUMIFS", "")
    m(135) = Array(3, "Lab Fees", "PL_SUMIFS", "")
    m(136) = Array(3, "Lease Payment", "PL_SUMIFS", "")
    m(137) = Array(3, "Licenses & Permits_", "PL_SUMIFS", "")
    m(138) = Array(3, "Pallets", "PL_SUMIFS", "")
    m(139) = Array(3, "Payroll Expense", "HEADER", "")
    m(140) = Array(4, "Payroll Taxes Production", "PL_SUMIFS", "")
    m(141) = Array(4, "Payroll Expense - Other", "PL_SUMIFS", "")
    m(142) = Array(4, "Salaries", "PL_SUMIFS", "")
    m(143) = Array(3, "Total Payroll Expense", "FORMULA", "=ROUND(SUM(D139:D142),5)")
    m(144) = Array(3, "Pest Control", "PL_SUMIFS", "")
    m(145) = Array(3, "Plate Charge", "PL_SUMIFS", "")
    m(146) = Array(3, "Production Supplies", "PL_SUMIFS", "")
    m(147) = Array(3, "Rent Expense", "PL_SUMIFS", "")
    m(148) = Array(3, "Repairs and Maintenance", "PL_SUMIFS", "")
    m(149) = Array(3, "Sales Tax/Use Tax", "PL_SUMIFS", "")
    m(150) = Array(3, "Tools & Instruments", "PL_SUMIFS", "")
    m(151) = Array(3, "Training and Seminars", "PL_SUMIFS", "")
    m(152) = Array(3, "Uniforms", "PL_SUMIFS", "")
    m(153) = Array(3, "Utilities", "PL_SUMIFS", "")
    m(154) = Array(3, "Validation Expenses", "PL_SUMIFS", "")
    m(155) = Array(3, "Workers Comp", "PL_SUMIFS", "")
    m(156) = Array(3, "FACTORY OVERHEAD - Other", "PL_SUMIFS", "")
    m(157) = Array(3, "Freight and Shipping Costs", "PL_SUMIFS", "")
    m(158) = Array(2, "Total FACTORY OVERHEAD", "FORMULA", "=ROUND(SUM(D122:D138)+SUM(D143:D157),5)")
    m(159) = Array(1, "Total COGS", "FORMULA", "=ROUND(D108+D114+D121+D158,5)")
    m(160) = Array(1, "Gross Profit", "FORMULA", "=ROUND(D107-D159,5)")
    m(161) = Array(1, "Expense", "HEADER", "")
    m(162) = Array(2, "ADMINISTRATION EXPENSES", "HEADER", "")
    m(163) = Array(3, "Accounting Fees", "PL_SUMIFS", "")
    m(164) = Array(3, "Automobile Expense", "HEADER", "")
    m(165) = Array(4, "Auto Repairs", "PL_SUMIFS", "")
End Sub

' =====================================================================
' Part 3: Rows 166–237  (Admin Expenses continued through Adjusted Net Income)
' =====================================================================
Private Sub InitRowsMapPart3(ByRef m() As Variant)
    m(166) = Array(4, "Gasoline, parking & mileage", "PL_SUMIFS", "")
    m(167) = Array(4, "Registration", "PL_SUMIFS", "")
    m(168) = Array(4, "Automobile Expense - Other", "PL_SUMIFS", "")
    m(169) = Array(3, "Total Automobile Expense", "FORMULA", "=ROUND(SUM(D164:D168),5)")
    m(170) = Array(3, "Bad Debts Expense", "PL_SUMIFS", "")
    m(171) = Array(3, "Bank Service Charges", "PL_SUMIFS", "")
    m(172) = Array(3, "Bank Document Fees", "PL_SUMIFS", "")
    m(173) = Array(3, "Business Tax", "PL_SUMIFS", "")
    m(174) = Array(3, "Computer and Internet Expenses", "PL_SUMIFS", "")
    m(175) = Array(3, "Credit Card Service Charge", "PL_SUMIFS", "")
    m(176) = Array(3, "Donation", "PL_SUMIFS", "")
    m(177) = Array(3, "Dues and Subscriptions", "PL_SUMIFS", "")
    m(178) = Array(3, "Insurance Expense", "HEADER", "")
    m(179) = Array(4, "General Liability", "PL_SUMIFS", "")
    m(180) = Array(4, "Property", "PL_SUMIFS", "")
    m(181) = Array(4, "Insurance Expense - Other", "PL_SUMIFS", "")
    m(182) = Array(4, "Vehicle Insurance", "PL_SUMIFS", "")
    m(183) = Array(3, "Total Insurance Expense", "FORMULA", "=ROUND(SUM(D178:D182),5)")
    m(184) = Array(3, "Legal", "PL_SUMIFS", "")
    m(185) = Array(3, "Licenses & Permits", "PL_SUMIFS", "")
    m(186) = Array(3, "Loan Processing Fees", "PL_SUMIFS", "")
    m(187) = Array(3, "Meals and Entertainment", "PL_SUMIFS", "")
    m(188) = Array(3, "Office Expenses", "PL_SUMIFS", "")
    m(189) = Array(3, "Payroll Expenses", "HEADER", "")
    m(190) = Array(4, "Bonus", "PL_SUMIFS", "")
    m(191) = Array(4, "Payroll Processing Fees", "PL_SUMIFS", "")
    m(192) = Array(4, "Payroll Taxes - Admin", "PL_SUMIFS", "")
    m(193) = Array(4, "Salaries - Admin", "PL_SUMIFS", "")
    m(194) = Array(4, "Medical Expenses", "PL_SUMIFS", "")
    m(195) = Array(4, "Employee Benefits", "PL_SUMIFS", "")
    m(196) = Array(4, "Payroll Expenses - Other", "PL_SUMIFS", "")
    m(197) = Array(3, "Total Payroll Expenses", "FORMULA", "=ROUND(SUM(D189:D196),5)")
    m(198) = Array(3, "Penalty & fine", "PL_SUMIFS", "")
    m(199) = Array(3, "Property Taxes Paid", "PL_SUMIFS", "")
    m(200) = Array(3, "Printing & Reproduction", "PL_SUMIFS", "")
    m(201) = Array(3, "Security and Alarm System", "PL_SUMIFS", "")
    m(202) = Array(3, "Suspense", "PL_SUMIFS", "")
    m(203) = Array(3, "Storage Expense", "PL_SUMIFS", "")
    m(204) = Array(3, "State Tax", "PL_SUMIFS", "")
    m(205) = Array(3, "ADMINISTRATION EXPENSES - Other", "PL_SUMIFS", "")
    m(206) = Array(3, "Telephone Expense", "PL_SUMIFS", "")
    m(207) = Array(2, "Total ADMINISTRATION EXPENSES", "FORMULA", "=ROUND(SUM(D162:D163)+SUM(D169:D177)+SUM(D183:D188)+SUM(D197:D206),5)")
    m(208) = Array(2, "SALES EXPENSES", "HEADER", "")
    m(209) = Array(3, "Commissions", "PL_SUMIFS", "")
    m(210) = Array(3, "Domestic Shipping & Handling", "PL_SUMIFS", "")
    m(211) = Array(3, "EDI Fees", "PL_SUMIFS", "")
    m(212) = Array(3, "Export Shipping & Handling", "PL_SUMIFS", "")
    m(213) = Array(3, "Postage and Delivery Expense", "PL_SUMIFS", "")
    m(214) = Array(3, "Service Fee", "PL_SUMIFS", "")
    m(215) = Array(3, "Shipping Supplies", "PL_SUMIFS", "")
    m(216) = Array(3, "Trade Show/Promotion", "PL_SUMIFS", "")
    m(217) = Array(3, "SALES EXPENSES - Other", "PL_SUMIFS", "")
    m(218) = Array(3, "Travel Expense", "PL_SUMIFS", "")
    m(219) = Array(2, "Total SALES EXPENSES", "FORMULA", "=ROUND(SUM(D208:D218),5)")
    m(220) = Array(2, "Taxes", "PL_SUMIFS", "")
    m(221) = Array(1, "Total Expense", "FORMULA", "=ROUND(D161+D207+SUM(D219:D220),5)")
    m(222) = Array(1, "Net Ordinary Income", "FORMULA", "=ROUND(D160-D221,5)")
    m(223) = Array(1, "Other Income/Expense", "HEADER", "")
    m(224) = Array(2, "Other Income", "HEADER", "")
    m(225) = Array(3, "Foreign Currency Gain/Loss", "PL_SUMIFS", "")
    m(226) = Array(3, "Other Income", "PL_SUMIFS", "")
    m(227) = Array(2, "Total Other Income", "FORMULA", "=ROUND(SUM(D224:D226),5)")
    m(228) = Array(2, "Other Expense", "HEADER", "")
    m(229) = Array(3, "Interest Expense", "PL_SUMIFS", "")
    m(230) = Array(3, "Loss(Gain) on disposal of Asset", "PL_SUMIFS", "")
    m(231) = Array(3, "Other Expense (adjustment)", "PL_SUMIFS", "")
    m(232) = Array(2, "Total Other Expense", "FORMULA", "=ROUND(SUM(D228:D231),5)")
    m(233) = Array(1, "Net Other Income", "FORMULA", "=ROUND(D227-D232,5)")
    m(234) = Array(1, "Net operating Income", "FORMULA", "=ROUND(D222+D233,5)")
    m(235) = Array(0, "", "BLANK", "")
    m(236) = Array(1, "Commission/Service Fee - Vietnam", "PL_SUMIFS", "")
    m(237) = Array(1, "Adjusted Net Income", "FORMULA", "=SUM(D234:D236)")
End Sub


' =====================================================================
' Helper: Find column letter by header text (row 1)
' =====================================================================
Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerText As String) As String
    Dim c        As Range
    Dim lastCol  As Long

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
' Helper: Add (or refresh) helper columns P=Level, Q=Name Cleaned.
' FIX: luôn refresh formula cho toàn b? lastRow — tránh bug khi tháng sau
'      d? li?u BS/PL tăng thêm d?ng.
' =====================================================================
Function AddHelperColumns(ws As Worksheet, lastRow As Long) As Long
    Dim colLevel As Long, colName As Long
    Dim r        As Long
    Dim q        As String

    q = Chr(34) & Chr(34)          ' = "" (empty string for FALSE branch)

    colLevel = ws.Range("P1").Column      ' 16
    colName = ws.Range("Q1").Column       ' 17

    ' Headers (idempotent)
    ws.Cells(1, colLevel).Value = "Level"
    ws.Cells(1, colName).Value = "Name Cleaned"
    ws.Range(ws.Cells(1, colLevel), ws.Cells(1, colName)).Font.Bold = True

    ' Always refresh formulas 3..lastRow
    For r = 3 To lastRow
        ws.Cells(r, colLevel).Formula = _
            "=IF(LEN(A" & r & ")>0,1,IF(LEN(B" & r & ")>0,2,IF(LEN(C" & r & ")>0,3," & _
            "IF(LEN(D" & r & ")>0,4,IF(LEN(E" & r & ")>0,5,IF(LEN(F" & r & ")>0,6," & _
            "IF(LEN(G" & r & ")>0,7," & q & ")))))))"

        ws.Cells(r, colName).Formula = _
            "=TRIM(PROPER(CLEAN(IF(LEN(A" & r & ")>0,A" & r & _
            ",IF(LEN(B" & r & ")>0,B" & r & ",IF(LEN(C" & r & ")>0,C" & r & _
            ",IF(LEN(D" & r & ")>0,D" & r & ",IF(LEN(E" & r & ")>0,E" & r & _
            ",IF(LEN(F" & r & ")>0,F" & r & ",G" & r & ")))))))))"
    Next r

    AddHelperColumns = colLevel
End Function