Sub CalculateRatios_SelectedRange()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '-----------------------------------------------------------
    ' 1. Validate selected range & read data
    '-----------------------------------------------------------
    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "No range selected. Please select a data range.", vbExclamation
        GoTo CleanExit
    End If
    Set rng = rng.Areas(1)

    Dim ws As Worksheet
    Set ws = rng.Parent

    Dim firstRow As Long, firstCol As Long
    Dim lastRow As Long, lastCol As Long
    firstRow = rng.Row
    firstCol = rng.Column
    lastRow = rng.Row + rng.Rows.count - 1
    lastCol = rng.Column + rng.Columns.count - 1

    If rng.Rows.count < 2 Or rng.Columns.count < 3 Then
        MsgBox "Range must include header row, label column, and at least 2 data columns.", vbExclamation
        GoTo CleanExit
    End If

    Dim dataArray As Variant
    dataArray = rng.Value

    '-----------------------------------------------------------
    ' 2. Determine first data column (skip VN/EN text, code, note, STT)
    ' Supports Format 1 (1 label col + FY cols) and Format 2 (VN, EN, Code, Note + Date cols)
    '-----------------------------------------------------------
    Dim dataStartCol As Long
    Dim found As Boolean
    Dim col As Long

    dataStartCol = 0
    found = False

    Dim quyStr As String, namStr As String
    quyStr = "qu" & ChrW(&HFD)          ' "quy" with accent (Quy = Quarter)
    namStr = "n" & ChrW(&H103) & "m"    ' "nam" with accent (Nam = Year)

    ' Priority 1: Check header for dates, quarters, years
    For col = 2 To UBound(dataArray, 2)
        Dim headerVal As String
        headerVal = Trim(CStr(dataArray(1, col)))
        If Len(headerVal) > 0 Then
            Dim hLower As String
            hLower = LCase(headerVal)
            Dim isDescCol As Boolean
            isDescCol = (InStr(1, hLower, "ma so", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "thuyet minh", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "chi tieu", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "line item", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "item code", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "stt", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "note", vbTextCompare) > 0)

            If Not isDescCol Then
                If InStr(1, headerVal, "/", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "-", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, quyStr, vbTextCompare) > 0 Or _
                   InStr(1, headerVal, namStr, vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "quarter", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "year", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "FY", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Jun", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Dec", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Mar", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Sep", vbTextCompare) > 0 Or _
                   (IsNumeric(headerVal) And Len(headerVal) = 4 And Val(headerVal) >= 1990 And Val(headerVal) <= 2100) Then
                    dataStartCol = col
                    found = True
                    Exit For
                End If
            End If
        End If
    Next col

    ' Priority 2: Check by inspecting data contents (finding first column with majority numeric values)
    If Not found Then
        For col = 2 To UBound(dataArray, 2)
            Dim numCount As Long, totalCheck As Long, checkR As Long
            numCount = 0: totalCheck = 0
            For checkR = 2 To Application.Min(UBound(dataArray, 1), 15)
                Dim vTest As Variant
                vTest = dataArray(checkR, col)
                If Len(Trim(CStr(vTest))) > 0 Then
                    totalCheck = totalCheck + 1
                    If IsNumeric(vTest) Then
                        If col <= 4 And (vTest = CLng(Val(vTest))) And Val(vTest) >= 1 And Val(vTest) <= 99 Then
                            ' likely code col
                        Else
                            numCount = numCount + 1
                        End If
                    End If
                End If
            Next checkR
            If totalCheck > 0 And (numCount / totalCheck) >= 0.5 Then
                dataStartCol = col
                found = True
                Exit For
            End If
        Next col
    End If

    ' Priority 3: Fallback verification prompt if confidence is not 100% or multi-column layout detected
    If Not found Or dataStartCol > 2 Then
        Dim promptDefault As Long
        If found Then promptDefault = dataStartCol Else promptDefault = 2
        Dim colPromptMsg As String
        colPromptMsg = "Confirm Data Start Column:" & vbNewLine & vbNewLine & _
                       "Detected column " & promptDefault & " ('" & dataArray(1, promptDefault) & "') as first data column." & vbNewLine & _
                       "Columns 1 to " & (promptDefault - 1) & " will be treated as item descriptions/codes." & vbNewLine & vbNewLine & _
                       "Enter column index (e.g. 2, 3, 5):"
        Dim userColInput As String
        userColInput = InputBox(colPromptMsg, "Data Start Column Confirmation", CStr(promptDefault))
        If userColInput = "" Then GoTo CleanExit
        If IsNumeric(userColInput) Then
            dataStartCol = CLng(userColInput)
            If dataStartCol >= 2 And dataStartCol <= UBound(dataArray, 2) Then
                found = True
            Else
                MsgBox "Column index out of range.", vbExclamation
                GoTo CleanExit
            End If
        End If
    End If

    If Not found Or dataStartCol > UBound(dataArray, 2) Then
        MsgBox "Could not determine data columns.", vbExclamation
        GoTo CleanExit
    End If

    '-----------------------------------------------------------
    ' 2.5. Ask user for Statement Mode (BS, IS, or Combined)
    '-----------------------------------------------------------
    Dim modePrompt As String
    modePrompt = "Select Financial Statement Mode to calculate ratios:" & vbNewLine & vbNewLine & _
                 "1 = Balance Sheet only (Bang Can Doi Ke Toan)" & vbNewLine & _
                 "2 = Income Statement only (Ket Qua Kinh Doanh)" & vbNewLine & _
                 "3 = Combined / Both (Toan bo bao cao tong hop)" & vbNewLine & vbNewLine & _
                 "Enter 1, 2, or 3 (or click Cancel to abort):"

    Dim runMode As String
    runMode = InputBox(modePrompt, "Statement Selection", "3")
    If StrPtr(runMode) = 0 Then
        MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
        GoTo CleanExit
    End If
    runMode = Trim(runMode)
    If runMode <> "1" And runMode <> "2" And runMode <> "3" Then
        MsgBox "Invalid selection. Please enter 1, 2, or 3.", vbExclamation, "Invalid Choice"
        GoTo CleanExit
    End If

    Dim includeIS As Boolean, includeBS As Boolean
    includeBS = (runMode = "1" Or runMode = "3")
    includeIS = (runMode = "2" Or runMode = "3")

    '-----------------------------------------------------------
    ' 3. Define required line items based on chosen mode
    '-----------------------------------------------------------
    Dim lineItems As Object
    Set lineItems = CreateObject("Scripting.Dictionary")

    Dim aA As String, aG As String, aH As String, aT As String, aD As String
    Dim awA As String
    Dim acA As String, acG As String, acD As String
    Dim eCirc As String, eCD As String
    Dim iA As String
    Dim oA As String, oG As String
    Dim ocA As String, ocD As String
    Dim ohD As String
    Dim uA As String, uhD As String, uH As String  ' BUG14 FIX: added uH declaration
    Dim yA As String
    Dim dd As String

    aA = ChrW(&HE1): aG = ChrW(&HE0): aH = ChrW(&H1EA3): aT = ChrW(&HE3): aD = ChrW(&H1EA1)
    awA = ChrW(&H1EAF)
    acA = ChrW(&H1EA5): acG = ChrW(&H1EA7): acD = ChrW(&H1EAD)
    eCirc = ChrW(&HEA): eCD = ChrW(&H1EC7)
    iA = ChrW(&HED)
    oA = ChrW(&HF3): oG = ChrW(&HF2)
    ocA = ChrW(&H1ED1): ocD = ChrW(&H1ED9)
    ohD = ChrW(&H1EE3)
    uA = ChrW(&HFA): uhD = ChrW(&H1EF1)
    uH = ChrW(&H1EED)   ' BUG14 FIX: "ủ" (u with hook above) used in TotalEquity VN keyword
    yA = ChrW(&HFD)
    dd = ChrW(&H111)

    ' ===== INCOME STATEMENT ITEMS =====
    If includeIS Then
        ' NetSales: treat "Revenue" (gross), "Gross Operating Revenue", and "Net Sales" as equivalent bases.
        ' Priority: Gross Operating Revenue -> Revenue -> Net Sales -> Net Revenue -> specific variants (VN/EN)
        ' Exclusions: guard against Financial Income, Other Income, Deferred Revenue being matched.
        lineItems.Add "NetSales", Array( _
            Array("gross operating revenue", "revenue", "net sales", "net turnover", _
                  "net revenue", "net revenue from sales", _
                  "doanh thu b" & aA & "n h" & aG & "ng", _
                  "doanh thu thu" & acG & "n", "tong doanh thu", "doanh thu"), _
            Array("deferred", "ch" & uA & " th" & iA & "ch", _
                  "financial", "financial income", "other income", "other incomes", "thu nhap khac", "tai chinh", _
                  "chua thuc hien", "unearned", "giam tru", "deductions", "revenue deductions"))

        lineItems.Add "COGS", Array( _
            Array("cost of goods sold", "cogs", "cost of sales", "gi" & aA & " v" & ocA & "n h" & aG & "ng b" & aA & "n", "gia von"), _
            Array("deferred"))

        lineItems.Add "GrossProfit", Array( _
            Array("gross profit", "l" & ohD & "i nhu" & acD & "n g" & ocD & "p", "loi nhuan gop"), _
            Array("deferred", "margin"))

        lineItems.Add "Depreciation", Array( _
            Array("accumulated depreciation and amortization", "depreciation and amortization", _
                  "accumulated depreciation", "depreciation", "amortization", "kh" & acA & "u hao"), _
            Array("deferred"))

        lineItems.Add "SellingExp", Array( _
            Array("selling expenses", "selling expense", "cost of sales", "chi ph" & iA & " b" & aA & "n h" & aG & "ng"), _
            Array("deferred", "cost of goods sold"))

        lineItems.Add "GAExp", Array( _
            Array("general and admin expenses", "general and admin", "general and administrative", _
                  "business management expenses", "management expenses", "administrative expenses", _
                  "chi ph" & iA & " qu" & aH & "n l" & yA & " doanh nghi" & eCD & "p"), _
            Array("deferred"))

        lineItems.Add "FinancialExp", Array( _
            Array("financial expenses", "financial expense", "financing costs", "finance costs", _
                  "chi ph" & iA & " t" & aG & "i ch" & iA & "nh"), _
            Array("borrowing costs", "interest expenses", "interest expense", "l" & aT & "i vay", "in which", "trong " & dd & oA))

        lineItems.Add "InterestExp", Array( _
            Array("of which: interest expenses", "of which: interest expense", "interest expenses", "interest expense", _
                  "borrowing costs", "chi ph" & iA & " l" & aT & "i vay"), _
            Array("deferred"))

        lineItems.Add "OperatingProfit", Array( _
            Array("net profit from operating activities", "net profit from business activities", "profit from business activities", _
                  "l" & ohD & "i nhu" & acD & "n thu" & acG & "n t" & uhD & " ho" & aD & "t " & dd & ocD & "ng kinh doanh", "loi nhuan thuan tu hdkd"), _
            Array("deferred", "financial"))

        lineItems.Add "ProfitBeforeTax", Array( _
            Array("pretax profit", "profit before tax", "accounting profit before tax", _
                  "t" & ohD & "ng l" & ohD & "i nhu" & acD & "n tr" & uhD & "" & ocA & "c thu" & eCirc, "loi nhuan truoc thue"), _
            Array("deferred"))

        ' NetProfit target: Support both parent company and general net profit
        lineItems.Add "NetProfit", Array( _
            Array("net profit after tax of parent company", "net profit after tax", "profit after corporate income tax", _
                  "profit after tax", "net profit", "l" & ohD & "i nhu" & acD & "n sau thu" & eCirc, "loi nhuan sau thue"), _
            Array("deferred", "before tax", "operating", "margin", "business activities", "minority", "benefits of minority"))
    End If

    ' ===== BALANCE SHEET ITEMS =====
    If includeBS Then
        lineItems.Add "Cash", Array( _
            Array("cash and cash equivalents", "cash equivalents", "cash", _
                  "ti" & eCirc & "n v" & aG & " c" & aA & "c kho" & aH & "n t" & uhD & "" & ohD & "ng " & dd & uhD & "" & ohD & "ng ti" & eCirc & "n"), _
            Array("provision", "d" & uhD & " ph" & oG & "ng"))

        lineItems.Add "ShortTermInvest", Array( _
            Array("short-term investments", "short term investments", "held to maturity investment", _
                  "" & dd & acG & "u t" & uA & " t" & aG & "i ch" & iA & "nh ng" & awA & "n h" & aD & "n"), _
            Array("provision", "provision for short term", "d" & uhD & " ph" & oG & "ng", "loss"))

        lineItems.Add "Receivables", Array( _
            Array("short term account receivables", "account receivable", "short-term trade receivables", _
                  "short-term receivables", "c" & aA & "c kho" & aH & "n ph" & aH & "i thu ng" & awA & "n h" & aD & "n"), _
            Array("provision", "provision for doubtful", "d" & uhD & " ph" & oG & "ng", "doubtful", _
                  "other receivables", "other receivable", "kh" & aA & "c", "long-term", "long term", "d" & aG & "i h" & aD & "n", "loan receivables"))

        lineItems.Add "Inventory", Array( _
            Array("inventories", "inventory", "h" & aG & "ng t" & ocD & "n kho", "hang ton kho"), _
            Array("provision", "d" & uhD & " ph" & oG & "ng"))

        ' CurrentAssets: Priority Current Assets; Exclude "other current assets"
        lineItems.Add "CurrentAssets", Array( _
            Array("current assets", "short-term assets", "short term assets", "t" & aG & "i s" & aH & "n ng" & awA & "n h" & aD & "n"), _
            Array("other current assets", "other current", "cash", "financial investment", "receivables", "inventory", "biological", "kh" & aA & "c"))

        lineItems.Add "FixedAssets", Array( _
            Array("tangible fixed assets", "fixed assets", "t" & aG & "i s" & aH & "n c" & ocA & " " & dd & "nh"), _
            Array("incomplete", "d" & ohD & " dang", "depreciation", "kh" & acA & "u hao", "financial leases", "intangible"))

        lineItems.Add "TotalAssets", Array( _
            Array("total assets", "t" & ohD & "ng c" & ocD & "ng t" & aG & "i s" & aH & "n", "tong tai san"), _
            Array("current assets", "short-term", "long-term", "non-current"))

        lineItems.Add "Payables", Array( _
            Array("payment for suppliers", "payable to short-term sellers", "trade payables", "short-term sellers", _
                  "ph" & aH & "i tr" & aH & " ng" & uhD & "" & ohD & "i b" & aA & "n ng" & awA & "n h" & aD & "n"), _
            Array("provision", "d" & uhD & " ph" & oG & "ng", "prepaid", "prepaid from customers", "kh" & aA & "ch h" & aG & "ng tr" & aH & " tr" & uhD & "" & ocA & "c", "tax", "employees"))

        ' CurrentLiabilities: Matches "short term liabilities" or "current liabilities".
        ' Exclude loans, payables, provisions so it matches the section total row.
        lineItems.Add "CurrentLiabilities", Array( _
            Array("short term liabilities", "current liabilities", "n" & ohD & " ng" & awA & "n h" & aD & "n"), _
            Array("short term loan", "loan and liabilities", "financial loans", "payment for suppliers", _
                  "payable to short-term", "dividends", "provision", "bonus", "unrealized", "tax", "employees", "expenses payable"))

        lineItems.Add "ShortTermDebt", Array( _
            Array("short term loan and liabilities", "short-term financial loans", "short term loans", "short-term loans", _
                  "vay v" & aG & " n" & ohD & " thu" & eCirc & " t" & aG & "i ch" & iA & "nh ng" & awA & "n h" & aD & "n"), _
            Array("long-term", "long term", "d" & aG & "i h" & aD & "n"))

        lineItems.Add "LongTermDebt", Array( _
            Array("loans and long term liabilities", "long-term financial loans", "long term loans", "long-term loans", _
                  "vay v" & aG & " n" & ohD & " thu" & eCirc & " t" & aG & "i ch" & iA & "nh d" & aG & "i h" & aD & "n"), _
            Array("short-term", "short term", "ng" & awA & "n h" & aD & "n"))

        ' TotalLiabilities: Matches "total liabilities" or "liabilities" (VAS 300).
        ' Strictly excludes "short term", "long term", "current" so "Liabilities" doesn't match sub-rows.
        lineItems.Add "TotalLiabilities", Array( _
            Array("total liabilities", "liabilities", "c - liabilities", "n" & ohD & " ph" & aH & "i tr" & aH & ""), _
            Array("short term", "short-term", "long term", "long-term", "current", "loan", "payable", "provision"))

        ' TotalEquity: Prioritize "owners equity" and "capital and reserves" over "total equity"
        ' because in many Vietnamese financial statements, the bottom row "Total Liabilities and Equity"
        ' is mislabeled in English as "Total Equity" (matching Total Assets).
        lineItems.Add "TotalEquity", Array( _
            Array("owners equity", "capital and reserves", "total equity", "d - equity", "equity", "v" & ocA & "n ch" & uH & " s" & ohD & " h" & uhD & "u"), _
            Array("paid-in", "common shares", "share premium", "treasury", "exchange", _
                  "undistributed", "non controlling", "minority", "contributed capital", "surplus", "voting", "preferred stock"))
    End If

    '-----------------------------------------------------------
    ' 4. Find each line item in the data
    ' In Combined Mode (mode 3), partition search into Balance Sheet section
    ' (before Revenue/Net Sales) and Income Statement section (from Revenue/Net Sales downwards).
    ' This prevents cross-contamination (e.g. BS Accumulated Depreciation matching IS Depreciation,
    ' or IS interest matching BS items).
    '-----------------------------------------------------------
    Dim dataStartRow As Long
    dataStartRow = 2
    Dim itemRows As Object
    Set itemRows = CreateObject("Scripting.Dictionary")

    Dim maxLabelCol As Long
    If dataStartCol > 1 Then
        maxLabelCol = dataStartCol - 1
    Else
        maxLabelCol = 1
    End If
    If maxLabelCol > 2 Then maxLabelCol = 2 ' Focus primary matching on Col 1 (VN) and Col 2 (EN)

    Dim totalRows As Long
    totalRows = UBound(dataArray, 1)

    ' Detect where Income Statement begins in Combined Mode
    Dim isSectionStartRow As Long
    isSectionStartRow = 0

    If runMode = "3" Then
        ' Search for Revenue / Net Sales row to demarcate the start of Income Statement
        Dim revKw As Variant
        revKw = Array("gross operating revenue", "revenue", "net sales", "doanh thu ban hang", "doanh thu thuan", "doanh thu")
        Dim revExcl As Variant
        revExcl = Array("deferred", "unrealized", "financial", "tai chinh", "other income", "thu nhap khac", "deductions", "giam tru")
        
        isSectionStartRow = FindRowByKeywords(dataArray, dataStartRow, totalRows, revKw, revExcl, maxLabelCol)
    End If

    Dim bsRowStart As Long, bsRowEnd As Long
    Dim isRowStart As Long, isRowEnd As Long

    If runMode = "3" And isSectionStartRow > dataStartRow Then
        bsRowStart = dataStartRow
        bsRowEnd = isSectionStartRow - 1
        isRowStart = isSectionStartRow
        isRowEnd = totalRows
    Else
        bsRowStart = dataStartRow
        bsRowEnd = totalRows
        isRowStart = dataStartRow
        isRowEnd = totalRows
    End If

    Dim key As Variant
    For Each key In lineItems.Keys
        Dim keywords As Variant
        Dim exclusions As Variant
        keywords = lineItems(key)(0)
        exclusions = lineItems(key)(1)

        Dim sStart As Long, sEnd As Long
        ' Partition based on whether the line item belongs to Income Statement or Balance Sheet
        Select Case key
            Case "NetSales", "COGS", "GrossProfit", "Depreciation", "SellingExp", "GAExp", _
                 "FinancialExp", "InterestExp", "OperatingProfit", "ProfitBeforeTax", "NetProfit"
                sStart = isRowStart
                sEnd = isRowEnd
            Case Else
                sStart = bsRowStart
                sEnd = bsRowEnd
        End Select

        Dim foundRow As Long
        foundRow = FindRowByKeywords(dataArray, sStart, sEnd, keywords, exclusions, maxLabelCol)
        If foundRow > 0 Then
            itemRows.Add key, foundRow
        Else
            itemRows.Add key, 0
        End If
    Next key

    '-----------------------------------------------------------
    ' 5. Show confirmation to user (with Cancel option)
    '-----------------------------------------------------------
    Dim modeTitle As String
    Select Case runMode
        Case "1": modeTitle = "Balance Sheet Mode"
        Case "2": modeTitle = "Income Statement Mode"
        Case "3": modeTitle = "Combined Mode (Both)"
    End Select

    Dim confirmMsg As String
    confirmMsg = "[" & modeTitle & "] Found line items:" & vbNewLine & vbNewLine
    For Each key In lineItems.Keys
        If itemRows(key) > 0 Then
            confirmMsg = confirmMsg & key & ": " & dataArray(itemRows(key), 1) & " (row " & itemRows(key) & ")" & vbNewLine
        Else
            confirmMsg = confirmMsg & key & ": NOT FOUND" & vbNewLine
        End If
    Next key
    confirmMsg = confirmMsg & vbNewLine & "Click 'Yes' to proceed, 'No' to edit manually, or 'Cancel' to abort."

    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(confirmMsg, vbYesNoCancel + vbQuestion, "Confirm Line Items - " & modeTitle)

    If userChoice = vbCancel Then
        MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
        GoTo CleanExit
    ElseIf userChoice = vbNo Then
        For Each key In lineItems.Keys
            Dim userInput As String
            If itemRows(key) = 0 Then
                userInput = InputBox("Enter row number for '" & key & "' (or leave blank to skip, Cancel to abort macro):", "Manual Entry", "")
                If StrPtr(userInput) = 0 Then
                    MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
                    GoTo CleanExit
                ElseIf userInput <> "" And IsNumeric(userInput) Then
                    Dim rowNum As Long
                    rowNum = CLng(userInput)
                    If rowNum >= dataStartRow And rowNum <= UBound(dataArray, 1) Then
                        itemRows(key) = rowNum
                    Else
                        MsgBox "Row number out of range. Skipping.", vbExclamation
                    End If
                End If
            Else
                Dim keepChoice As VbMsgBoxResult
                keepChoice = MsgBox("Keep '" & key & "' at row " & itemRows(key) & " (" & dataArray(itemRows(key), 1) & ")?" & vbNewLine & _
                                    "Yes: Keep | No: Change row | Cancel: Abort macro", vbYesNoCancel + vbQuestion, "Confirm Row")
                If keepChoice = vbCancel Then
                    MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
                    GoTo CleanExit
                ElseIf keepChoice = vbNo Then
                    userInput = InputBox("Enter new row number for '" & key & "':", "Change Row", itemRows(key))
                    If StrPtr(userInput) = 0 Then
                        MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
                        GoTo CleanExit
                    ElseIf userInput <> "" And IsNumeric(userInput) Then
                        rowNum = CLng(userInput)
                        If rowNum >= dataStartRow And rowNum <= UBound(dataArray, 1) Then
                            itemRows(key) = rowNum
                        Else
                            MsgBox "Row number out of range. Keeping old value.", vbExclamation
                        End If
                    End If
                End If
            End If
        Next key
    End If

    '-----------------------------------------------------------
    ' 6. Prepare helper function to get Total Debt (if BS included)
    '-----------------------------------------------------------
    Dim debtRow As Long
    Dim useTotalLiabilities As Boolean
    debtRow = 0
    useTotalLiabilities = False

    If includeBS Then
        If itemRows("ShortTermDebt") > 0 And itemRows("LongTermDebt") > 0 Then
            debtRow = -1   ' special flag: compute from two rows
            useTotalLiabilities = False
        ElseIf itemRows("TotalLiabilities") > 0 Then
            debtRow = itemRows("TotalLiabilities")
            useTotalLiabilities = True
        Else
            debtRow = 0
            useTotalLiabilities = False
        End If
    End If

    '-----------------------------------------------------------
    ' 7. Define all ratios to compute according to Mode
    '-----------------------------------------------------------
    Dim ratios As Object
    Set ratios = CreateObject("Scripting.Dictionary")

    ' --- 7A. Income Statement Mode (runMode = "2") ---
    If runMode = "2" Then
        ratios.Add "Gross profit margin (% sales)", "GrossProfit/NetSales"
        ratios.Add "Cost of goods sold (% sales)", "COGS/NetSales"
        ratios.Add "Operating profit margin (%)", "OperatingProfit/NetSales"
        ratios.Add "Selling expense/sales (%)", "SellingExp/NetSales"
        ratios.Add "Administrative expense/sales (%)", "GAExp/NetSales"
        ratios.Add "Financial expense/sales (%)", "FinancialExp/NetSales"
        ratios.Add "Depreciation rate (% sales)", "Depreciation/NetSales"
        ratios.Add "Net Profit margin (%)", "NetProfit/NetSales"

    ' --- 7B. Balance Sheet Mode (runMode = "1") ---
    ElseIf runMode = "1" Then
        ' Liquidity
        ratios.Add "Current ratio", "CurrentAssets/CurrentLiabilities"
        ratios.Add "Acid test ratio", "(CurrentAssets-Inventory)/CurrentLiabilities"
        ratios.Add "Cash ratio", "Cash/CurrentLiabilities"
        ratios.Add "Quick ratio", "(Cash+ShortTermInvest+Receivables)/CurrentLiabilities"
        ' Solvency & Capital Structure
        ratios.Add "Total debt/equity ratio", "TotalDebt/TotalEquity"
        ratios.Add "Total equity/debt ratio", "TotalEquity/TotalDebt"
        ratios.Add "Equity/(Debt+Equity)", "TotalEquity/(TotalDebt+TotalEquity)"
        ratios.Add "Debt/(Debt+Equity)", "TotalDebt/(TotalDebt+TotalEquity)"
        ratios.Add "Debt to assets ratio (%)", "TotalDebt/TotalAssets"

    ' --- 7C. Combined Mode (runMode = "3") ---
    Else
        ' Profitability
        ratios.Add "Cost of goods sold (% sales)", "COGS/NetSales"
        ratios.Add "Gross profit margin (% sales)", "GrossProfit/NetSales"
        ratios.Add "Operating profit margin (%)", "OperatingProfit/NetSales"
        ratios.Add "Depreciation rate (% sales)", "Depreciation/NetSales"
        ratios.Add "Selling expense/sales (%)", "SellingExp/NetSales"
        ratios.Add "Administrative expense/sales (%)", "GAExp/NetSales"
        ratios.Add "Financial expense/sales (%)", "FinancialExp/NetSales"
        ratios.Add "Net Profit margin (%)", "NetProfit/NetSales"
        ratios.Add "Return on total assets - ROA (%)", "NetProfit/TotalAssets"
        ratios.Add "Return on equity - ROE (%)", "NetProfit/TotalEquity"
        ratios.Add "Return on capital employed - ROCE (%)", "EBIT/CapitalEmployed"

        ' Operation & Activity (Days)
        ratios.Add "Inventory days", "Inventory/(COGS/365)"
        ratios.Add "Receivable days", "Receivables/(NetSales/365)"
        ratios.Add "Payables days", "Payables/(COGS/365)"

        ' Liquidity
        ratios.Add "Current ratio", "CurrentAssets/CurrentLiabilities"
        ratios.Add "Acid test ratio", "(CurrentAssets-Inventory)/CurrentLiabilities"
        ratios.Add "Cash ratio", "Cash/CurrentLiabilities"
        ratios.Add "Quick ratio", "(Cash+ShortTermInvest+Receivables)/CurrentLiabilities"

        ' Solvency & Coverage
        ratios.Add "Total debt/equity ratio", "TotalDebt/TotalEquity"
        ratios.Add "Total equity/debt ratio", "TotalEquity/TotalDebt"
        ratios.Add "Equity/(Debt+Equity)", "TotalEquity/(TotalDebt+TotalEquity)"
        ratios.Add "Debt/(Debt+Equity)", "TotalDebt/(TotalDebt+TotalEquity)"
        ratios.Add "Debt to assets ratio (%)", "TotalDebt/TotalAssets"
        ratios.Add "Interest coverage ratio", "EBIT/InterestExp"
    End If

    '-----------------------------------------------------------
    ' 8. Build ratio FORMULAS for each data column
    '-----------------------------------------------------------
    Dim baseRow As Long, baseCol As Long
    baseRow = rng.Row
    baseCol = rng.Column

    Dim numDataCols As Long
    numDataCols = UBound(dataArray, 2) - dataStartCol + 1

    ' Check if IS section has more columns than the BS header row
    If runMode = "3" And isSectionStartRow > 0 Then
        Dim maxISCols As Long, rTest As Long, cTest As Long
        maxISCols = numDataCols
        For rTest = isSectionStartRow To totalRows
            For cTest = UBound(dataArray, 2) To dataStartCol Step -1
                If Len(Trim(CStr(dataArray(rTest, cTest)))) > 0 Then
                    If (cTest - dataStartCol + 1) > maxISCols Then
                        maxISCols = cTest - dataStartCol + 1
                    End If
                    Exit For
                End If
            Next cTest
        Next rTest
        
        If maxISCols > numDataCols Then
            MsgBox "Notice: The Income Statement section contains " & maxISCols & " data columns, " & vbNewLine & _
                   "while the Balance Sheet header has " & numDataCols & " columns." & vbNewLine & vbNewLine & _
                   "Ratios will be calculated for the common " & numDataCols & " periods." & vbNewLine & _
                   "To calculate all " & maxISCols & " periods for P&L-only ratios, please run in 'Income Statement Mode' separately.", _
                   vbInformation, "Column Alignment Notice"
        End If
    End If

    Dim resultArray As Variant
    ReDim resultArray(1 To ratios.count + 1, 1 To numDataCols + 1)

    ' Header row
    resultArray(1, 1) = "Ratio (" & modeTitle & ")"
    Dim dCol As Long
    For dCol = 1 To numDataCols
        resultArray(1, dCol + 1) = dataArray(1, dataStartCol + dCol - 1)
    Next dCol

    Dim DQ As String
    DQ = Chr(34)

    Dim ratioIdx As Long
    ratioIdx = 2
    Dim ratioName As Variant
    For Each ratioName In ratios.Keys
        resultArray(ratioIdx, 1) = ratioName

        For dCol = 1 To numDataCols
            Dim colIdx As Long
            colIdx = dataStartCol + dCol - 1

            ' Cell references as text
            Dim refNetSales As String, refCOGS As String, refGrossProfit As String
            Dim refDepreciation As String, refSellingExp As String, refGAExp As String
            Dim refFinancialExp As String, refInterestExp As String
            Dim refOperatingProfit As String, refProfitBeforeTax As String, refNetProfit As String

            If includeIS Then
                refNetSales = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "NetSales"), colIdx)
                refCOGS = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "COGS"), colIdx)
                refGrossProfit = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "GrossProfit"), colIdx)
                refDepreciation = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "Depreciation"), colIdx)
                refSellingExp = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "SellingExp"), colIdx)
                refGAExp = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "GAExp"), colIdx)
                refFinancialExp = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "FinancialExp"), colIdx)
                refInterestExp = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "InterestExp"), colIdx)
                refOperatingProfit = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "OperatingProfit"), colIdx)
                refProfitBeforeTax = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "ProfitBeforeTax"), colIdx)
                refNetProfit = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "NetProfit"), colIdx)
            End If

            Dim refCash As String, refShortTermInvest As String, refReceivables As String
            Dim refInventory As String, refCurrentAssets As String, refFixedAssets As String, refTotalAssets As String
            Dim refPayables As String, refCurrentLiab As String, refTotalEquity As String
            Dim refTotalDebt As String

            If includeBS Then
                refCash = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "Cash"), colIdx)
                refShortTermInvest = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "ShortTermInvest"), colIdx)
                refReceivables = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "Receivables"), colIdx)
                refInventory = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "Inventory"), colIdx)
                refCurrentAssets = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "CurrentAssets"), colIdx)
                refFixedAssets = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "FixedAssets"), colIdx)
                refTotalAssets = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "TotalAssets"), colIdx)

                refPayables = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "Payables"), colIdx)
                refCurrentLiab = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "CurrentLiabilities"), colIdx)
                refTotalEquity = ItemRef(ws, baseRow, baseCol, SafeGetItem(itemRows, "TotalEquity"), colIdx)

                If debtRow = -1 Then
                    refTotalDebt = "(" & ItemRef(ws, baseRow, baseCol, itemRows("ShortTermDebt"), colIdx) & _
                                   "+" & ItemRef(ws, baseRow, baseCol, itemRows("LongTermDebt"), colIdx) & ")"
                ElseIf debtRow > 0 Then
                    refTotalDebt = ItemRef(ws, baseRow, baseCol, debtRow, colIdx)
                Else
                    refTotalDebt = "0"
                End If
            End If

            ' BUG10 FIX: EBIT priority corrected.
            ' In VAS, "OperatingProfit" (Loi nhuan tu HDKD) ALREADY deducts interest/financial expenses.
            ' EBIT (international standard) = Profit Before Tax + Interest Expenses.
            ' Priority: PBT+Interest > PBT+FinancialExp > OperatingProfit (approximation only) > 0
            Dim refEbit As String
            If includeIS Then
                If SafeGetItem(itemRows, "ProfitBeforeTax") > 0 And SafeGetItem(itemRows, "InterestExp") > 0 Then
                    ' Most accurate: EBIT = PBT + Interest Expense (IFRS standard)
                    refEbit = "(" & refProfitBeforeTax & "+" & refInterestExp & ")"
                ElseIf SafeGetItem(itemRows, "ProfitBeforeTax") > 0 And SafeGetItem(itemRows, "FinancialExp") > 0 Then
                    ' Fallback: EBIT ≈ PBT + Total Financial Expenses (overestimates slightly)
                    refEbit = "(" & refProfitBeforeTax & "+" & refFinancialExp & ")"
                ElseIf SafeGetItem(itemRows, "OperatingProfit") > 0 Then
                    ' Last resort: use Operating Profit as EBIT approximation (understates for VAS)
                    refEbit = refOperatingProfit
                Else
                    refEbit = "0"
                End If
            Else
                refEbit = "0"
            End If

            Dim refCapitalEmployed As String
            If includeBS Then
                ' International standard: Capital Employed = Total Assets - Current Liabilities
                If SafeGetItem(itemRows, "TotalAssets") > 0 And SafeGetItem(itemRows, "CurrentLiabilities") > 0 Then
                    refCapitalEmployed = "(" & refTotalAssets & "-" & refCurrentLiab & ")"
                Else
                    refCapitalEmployed = "(" & refTotalEquity & "+" & refTotalDebt & ")"
                End If
            Else
                refCapitalEmployed = "0"
            End If

            ' Build Excel formula
            Dim formulaStr As String
            Select Case ratioName
                Case "Cost of goods sold (% sales)"
                    ' Wrap ABS in case raw COGS is recorded as negative in financial reports
                    formulaStr = "=IFERROR(ABS(" & refCOGS & ")/" & refNetSales & "," & DQ & DQ & ")"
                Case "Gross profit margin (% sales)"
                    formulaStr = MakeRatioFormula(refGrossProfit, refNetSales, DQ)
                Case "Operating profit margin (%)"
                    formulaStr = MakeRatioFormula(refOperatingProfit, refNetSales, DQ)
                Case "Depreciation rate (% sales)"
                    If SafeGetItem(itemRows, "Depreciation") > 0 Then
                        formulaStr = MakeRatioFormula(refDepreciation, refNetSales, DQ)
                    Else
                        formulaStr = ""
                    End If
                Case "Selling expense/sales (%)"
                    formulaStr = MakeRatioFormula(refSellingExp, refNetSales, DQ)
                Case "Administrative expense/sales (%)"
                    formulaStr = MakeRatioFormula(refGAExp, refNetSales, DQ)
                Case "Financial expense/sales (%)"
                    formulaStr = MakeRatioFormula(refFinancialExp, refNetSales, DQ)
                Case "Net Profit margin (%)"
                    formulaStr = MakeRatioFormula(refNetProfit, refNetSales, DQ)
                Case "Return on total assets - ROA (%)"
                    formulaStr = MakeRatioFormula(refNetProfit, refTotalAssets, DQ)
                Case "Return on equity - ROE (%)"
                    formulaStr = MakeRatioFormula(refNetProfit, refTotalEquity, DQ)
                Case "Return on capital employed - ROCE (%)"
                    formulaStr = MakeRatioFormula(refEbit, refCapitalEmployed, DQ)

                Case "Inventory days"
                    ' Standardized formula: (Inventory * 365) / ABS(COGS)
                    formulaStr = "=IFERROR((" & refInventory & "*365)/ABS(" & refCOGS & ")," & DQ & DQ & ")"
                Case "Receivable days"
                    ' Standardized formula: (Receivables * 365) / NetSales
                    formulaStr = "=IFERROR((" & refReceivables & "*365)/" & refNetSales & "," & DQ & DQ & ")"
                Case "Payables days"
                    ' Standardized formula: (Payables * 365) / ABS(COGS)
                    formulaStr = "=IFERROR((" & refPayables & "*365)/ABS(" & refCOGS & ")," & DQ & DQ & ")"

                Case "Current ratio"
                    formulaStr = MakeRatioFormula(refCurrentAssets, refCurrentLiab, DQ)
                Case "Acid test ratio"
                    formulaStr = MakeRatioFormula("(" & refCurrentAssets & "-" & refInventory & ")", refCurrentLiab, DQ)
                Case "Cash ratio"
                    formulaStr = MakeRatioFormula(refCash, refCurrentLiab, DQ)
                Case "Quick ratio"
                    formulaStr = MakeRatioFormula("(" & refCash & "+" & refShortTermInvest & "+" & refReceivables & ")", refCurrentLiab, DQ)

                Case "Total debt/equity ratio"
                    formulaStr = MakeRatioFormula(refTotalDebt, refTotalEquity, DQ)
                Case "Total equity/debt ratio"
                    formulaStr = MakeRatioFormula(refTotalEquity, refTotalDebt, DQ)
                Case "Equity/(Debt+Equity)"
                    formulaStr = MakeRatioFormula(refTotalEquity, "(" & refTotalDebt & "+" & refTotalEquity & ")", DQ)
                Case "Debt/(Debt+Equity)"
                    formulaStr = MakeRatioFormula(refTotalDebt, "(" & refTotalDebt & "+" & refTotalEquity & ")", DQ)
                Case "Debt to assets ratio (%)"
                    formulaStr = MakeRatioFormula(refTotalDebt, refTotalAssets, DQ)
                Case "Interest coverage ratio"
                    formulaStr = MakeRatioFormula(refEbit, refInterestExp, DQ)
                Case Else
                    formulaStr = ""
            End Select

            resultArray(ratioIdx, dCol + 1) = formulaStr
        Next dCol

        ratioIdx = ratioIdx + 1
    Next ratioName

    '-----------------------------------------------------------
    ' 9. Ask destination and paste result AS FORMULAS
    '-----------------------------------------------------------
    Dim destRange As Range
    On Error Resume Next
    Set destRange = Application.InputBox("Select the top-left cell for the " & modeTitle & " ratios result:", "Destination", Type:=8)
    On Error GoTo ErrorHandler
    If destRange Is Nothing Then GoTo CleanExit
    Set destRange = destRange.Cells(1, 1)

    Dim outRows As Long, outCols As Long
    outRows = UBound(resultArray, 1) - LBound(resultArray, 1) + 1
    outCols = UBound(resultArray, 2) - LBound(resultArray, 2) + 1

    destRange.Resize(outRows, outCols).Formula = resultArray

    ' Apply formatting
    Dim i As Long, j As Long
    For i = 2 To outRows
        For j = 2 To outCols
            Dim cell As Range
            Set cell = destRange.offset(i - 1, j - 1)
            If IsNumeric(cell.Value) Then
                Dim ratioLabel As String
                ratioLabel = resultArray(i, 1)
                If InStr(1, ratioLabel, "days", vbTextCompare) > 0 Or _
                   InStr(1, ratioLabel, "ratio", vbTextCompare) > 0 Or _
                   InStr(1, ratioLabel, "Current", vbTextCompare) > 0 Or _
                   InStr(1, ratioLabel, "Acid", vbTextCompare) > 0 Or _
                   InStr(1, ratioLabel, "coverage", vbTextCompare) > 0 Then
                    cell.NumberFormat = "0.00"
                Else
                    cell.NumberFormat = "0.00%"
                End If
            End If
        Next j
    Next i

    MsgBox "Ratios calculation completed for " & modeTitle & "!" & vbNewLine & _
           "Results pasted at " & destRange.Address, vbInformation
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------
' Helper: Safely get item row from dictionary
'------------------------------------------------------------------
Function SafeGetItem(dict As Object, keyName As String) As Long
    If dict.Exists(keyName) Then
        SafeGetItem = dict(keyName)
    Else
        SafeGetItem = 0
    End If
End Function

'------------------------------------------------------------------
' Helper: Find row in array based on keywords and exclusions.
' Searches across description columns (Col 1 to maxLabelCol, e.g. VN and EN)
'------------------------------------------------------------------
Function FindRowByKeywords(dataArr As Variant, startRow As Long, endRow As Long, _
                           keywords As Variant, exclusions As Variant, _
                           Optional maxLabelCol As Long = 1) As Long
    Dim i As Long, kw As Variant, excl As Variant
    Dim cellText As String, normText As String
    Dim containsKw As Boolean, isExcluded As Boolean
    Dim scanEnd As Long
    Dim colIdx As Long

    scanEnd = IIf(endRow > 0 And endRow <= UBound(dataArr, 1), endRow, UBound(dataArr, 1))
    If maxLabelCol < 1 Then maxLabelCol = 1
    If maxLabelCol > 2 Then maxLabelCol = 2 ' Primary text scanning on Col 1 (VN) and Col 2 (EN)

    For i = startRow To scanEnd
        For colIdx = 1 To maxLabelCol
            cellText = CStr(dataArr(i, colIdx))
            If Len(cellText) > 0 Then
                normText = StripAccents(LCase(Trim(cellText)))
                containsKw = False
                For Each kw In keywords
                    If InStr(1, normText, StripAccents(LCase(CStr(kw))), vbTextCompare) > 0 Then
                        containsKw = True
                        Exit For
                    End If
                Next kw

                If containsKw Then
                    ' Check exclusions across both label columns of this row
                    isExcluded = False
                    Dim eCol As Long
                    For eCol = 1 To maxLabelCol
                        Dim eText As String
                        eText = StripAccents(LCase(Trim(CStr(dataArr(i, eCol)))))
                        For Each excl In exclusions
                            If InStr(1, eText, StripAccents(LCase(CStr(excl))), vbTextCompare) > 0 Then
                                isExcluded = True
                                Exit For
                            End If
                        Next excl
                        If isExcluded Then Exit For
                    Next eCol

                    If Not isExcluded Then
                        FindRowByKeywords = i
                        Exit Function
                    End If
                End If
            End If
        Next colIdx
    Next i
    FindRowByKeywords = 0
End Function

'------------------------------------------------------------------
' Helper: Build cell reference string
' Use absolute cell address Address(True, True) to avoid formula drift
'------------------------------------------------------------------
Function ItemRef(ws As Worksheet, baseRow As Long, baseCol As Long, _
                  itemRowInData As Long, colIdxInData As Long) As String
    If itemRowInData <= 0 Then
        ItemRef = "0"
    Else
        ItemRef = "'" & Replace(ws.Name, "'", "''") & "'!" & _
            ws.Cells(baseRow + itemRowInData - 1, baseCol + colIdxInData - 1).Address(True, True)
    End If
End Function

'------------------------------------------------------------------
' Helper: Build Excel formula "=IFERROR(num/den,"")"
'------------------------------------------------------------------
Function MakeRatioFormula(numRef As String, denRef As String, dq As String) As String
    MakeRatioFormula = "=IFERROR(" & numRef & "/" & denRef & "," & dq & dq & ")"
End Function

'------------------------------------------------------------------
' Helper: Strip Vietnamese accents
'------------------------------------------------------------------
Function StripAccents(ByVal txt As String) As String
    Dim src As String, dst As String

    ' a with tones
    src = src & ChrW(&HE0) & ChrW(&HE1) & ChrW(&H1EA3) & ChrW(&HE3) & ChrW(&H1EA1)
    dst = dst & "aaaaa"

    ' a-breve with tones
    src = src & ChrW(&H103) & ChrW(&H1EB1) & ChrW(&H1EAF) & ChrW(&H1EB3) & ChrW(&H1EB5) & ChrW(&H1EB7)
    dst = dst & "aaaaaa"

    ' a-circumflex with tones
    src = src & ChrW(&HE2) & ChrW(&H1EA7) & ChrW(&H1EA5) & ChrW(&H1EA9) & ChrW(&H1EAB) & ChrW(&H1EAD)
    dst = dst & "aaaaaa"

    ' e with tones
    src = src & ChrW(&HE8) & ChrW(&HE9) & ChrW(&H1EBB) & ChrW(&H1EBD) & ChrW(&H1EB9)
    dst = dst & "eeeee"

    ' e-circumflex with tones
    src = src & ChrW(&HEA) & ChrW(&H1EC1) & ChrW(&H1EBF) & ChrW(&H1EC3) & ChrW(&H1EC5) & ChrW(&H1EC7)
    dst = dst & "eeeeee"

    ' i with tones
    src = src & ChrW(&HEC) & ChrW(&HED) & ChrW(&H1EC9) & ChrW(&H129) & ChrW(&H1ECB)
    dst = dst & "iiiii"

    ' o with tones
    src = src & ChrW(&HF2) & ChrW(&HF3) & ChrW(&H1ECF) & ChrW(&HF5) & ChrW(&H1ECD)
    dst = dst & "ooooo"

    ' o-circumflex with tones
    src = src & ChrW(&HF4) & ChrW(&H1ED3) & ChrW(&H1ED1) & ChrW(&H1ED5) & ChrW(&H1ED7) & ChrW(&H1ED9)
    dst = dst & "oooooo"

    ' o-horn with tones
    src = src & ChrW(&H1A1) & ChrW(&H1EDD) & ChrW(&H1EDB) & ChrW(&H1EDF) & ChrW(&H1EE1) & ChrW(&H1EE3)
    dst = dst & "oooooo"

    ' u with tones
    src = src & ChrW(&HF9) & ChrW(&HFA) & ChrW(&H1EE7) & ChrW(&H169) & ChrW(&H1EE5)
    dst = dst & "uuuuu"

    ' u-horn with tones
    src = src & ChrW(&H1B0) & ChrW(&H1EEB) & ChrW(&H1EE9) & ChrW(&H1EED) & ChrW(&H1EEF) & ChrW(&H1EF1)
    dst = dst & "uuuuuu"

    ' y with tones
    src = src & ChrW(&H1EF3) & ChrW(&HFD) & ChrW(&H1EF7) & ChrW(&H1EF9) & ChrW(&H1EF5)
    dst = dst & "yyyyy"

    ' dj (Vietnamese d with stroke)
    src = src & ChrW(&H111)
    dst = dst & "d"

    Dim result As String
    Dim k As Long
    result = txt
    For k = 1 To Len(src)
        result = Replace(result, Mid(src, k, 1), Mid(dst, k, 1))
    Next k
    StripAccents = result
End Function
