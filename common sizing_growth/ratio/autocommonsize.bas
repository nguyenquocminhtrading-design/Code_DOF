Sub AutoCommonSize_SelectedRange()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '-----------------------------------------------------------
    ' 1. Validate & read data
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

    If rng.Rows.count < 2 Or rng.Columns.count < 2 Then
        MsgBox "Range must include header row, label column, and at least 1 data column.", vbExclamation
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

    ' Priority 1: Check header for dates, quarters, years (e.g. 2023, 2024, Q1, Quy, /, FY, Jun, Jan, etc.)
    For col = 2 To UBound(dataArray, 2)
        Dim headerVal As String
        headerVal = Trim(CStr(dataArray(1, col)))
        If Len(headerVal) > 0 Then
            ' Skip columns that are clearly Code/STT/Thuyet minh even if header looks strange
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
                ' Check if header represents a date/period
                If InStr(1, headerVal, "/", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "-", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Q", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Qu", vbTextCompare) > 0 Or _
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
                        ' If it's a small integer <= 100 in early columns, it might be a Code column
                        If col <= 4 And (vTest = CLng(Val(vTest))) And Val(vTest) >= 1 And Val(vTest) <= 99 Then
                            ' likely a code column like 01, 02
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
    ' 3. Ask user for report type
    '-----------------------------------------------------------
    Dim reportChoice As Variant
    reportChoice = Application.InputBox( _
        "Select Common Size base:" & vbNewLine & _
        "1 - Income Statement (Revenue base)" & vbNewLine & _
        "2 - Balance Sheet (Total Assets base)" & vbNewLine & _
        "3 - Auto BS + PL (Dual Base)" & vbNewLine & _
        "4 - Cash Flow Statement (Net Operating Cash Flow base)", _
        "Report Type", 3, Type:=1)
    If reportChoice = False Or reportChoice = "" Then Exit Sub
    If reportChoice < 1 Or reportChoice > 4 Then
        MsgBox "Invalid choice. Please enter 1, 2, 3, or 4.", vbExclamation
        Exit Sub
    End If

    '-----------------------------------------------------------
    ' 4. Define keywords (English & Vietnamese unaccented)
    '-----------------------------------------------------------
    ' BS Total Assets keywords: covers English & VN formats (e.g. "TOTAL ASSETS (270 = 100 + 200)", "TONG CONG TAI SAN", etc.)
    Dim bsKeywords As Variant
    bsKeywords = Array( _
        "total assets", "tong cong tai san", "tong tai san", "cong tai san", _
        "total asset", "assets total", "assets = 100 + 200", "tai san (270", "tai san 270")

    ' BS Exclude words (sub-totals)
    Dim bsExclude As Variant
    bsExclude = Array( _
        "current assets", "non-current assets", "short-term", "long-term", _
        "ngan han", "dai han", "co dinh", "fixed", "other assets", "tai san khac", "sinh hoc", "biological")

    ' PL Revenue keywords
    ' Priority: "Gross Operating Revenue" -> "Revenue" (gross) -> "Net Sales" / "Net Revenue" -> specific variants -> generic fallbacks
    ' Per user requirement: treat Revenue and Net Sales as equivalent; prioritize Revenue if present.
    Dim plKeywords As Variant
    plKeywords = Array( _
        "gross operating revenue", "revenue", "net sales", "net turnover", _
        "doanh thu ban hang va cung cap dich vu", "tong doanh thu", "doanh thu", _
        "net revenue from sales and provision of services", "net revenue from sales and service", _
        "net revenue from sales", "net revenue", _
        "doanh thu thuan ve ban hang va cung cap dich vu", "doanh thu thuan ve ban hang", "doanh thu thuan", _
        "revenue from sales and provision of services", "revenue from sales and services", _
        "revenue from sales", "gross sales", "gross revenue", "total revenue")

    Dim plExclude As Variant
    plExclude = Array( _
        "deferred", "unrealized", "financial", "tai chinh", "chua thuc hien", _
        "cho phan bo", "hoat dong tai chinh", "financial income", "other income", "other incomes", "thu nhap khac", _
        "deductions", "giam tru", "revenue deductions", "unearned")

    ' CF Cash Flow keywords (Net Cash Flow from Operating Activities)
    ' Supports both indirect IFRS method and Vietnamese TT99 / TT200 bilingual standards
    Dim cfKeywords As Variant
    cfKeywords = Array( _
        "net cash flow from operating activities", _
        "net cashflow from operating activities", _
        "net cash flows from operating activities", _
        "net cash flow from business activities", _
        "net cashflow from business activities", _
        "luu chuyen tien thuan tu hoat dong kinh doanh", _
        "luu chuyen thuan tu hoat dong kinh doanh", _
        "tien thuan tu hoat dong kinh doanh", _
        "operating activities", _
        "business activities")

    Dim cfExclude As Variant
    cfExclude = Array( _
        "investing", "financing", "dau tu", "tai chinh", _
        "indirect method", "direct method", _
        "before changes", "truoc thay doi", _
        "other incomes", "adjustments", "dieu chinh", _
        "working capital", "von luu dong", "contingencies", "du phong")

    ' Section header separator keywords between BS and PL (also matches PL distinctive row items)
    Dim plSectionKeywords As Variant
    plSectionKeywords = Array( _
        "income statement", "profit and loss", "p&l", "ket qua hoat dong kinh doanh", _
        "bao cao ket qua", "kqhdkd", "ket qua kinh doanh", "báo cáo kqhđkd", "ii. income statement", "ii. ket qua", _
        "net sales", "cost of goods sold", "gross profit", "doanh thu thuan", "gia von hang ban", "loi nhuan gop")

    '-----------------------------------------------------------
    ' 5. Detect PL section start row (if full financial report)
    ' Supports parallel checking of column 1 (VN) and column 2 (EN)
    '-----------------------------------------------------------
    Dim dataStartRow As Long
    dataStartRow = 2

    ' Identify description columns (all columns before dataStartCol, excluding purely numeric Code/STT columns)
    Dim maxLabelCol As Long
    If dataStartCol > 1 Then
        maxLabelCol = dataStartCol - 1
    Else
        maxLabelCol = 1
    End If
    If maxLabelCol > 2 Then maxLabelCol = 2 ' Focus primary matching on Col 1 (VN) and Col 2 (EN)

    Dim labelCol As Long
    labelCol = 1

    Dim plStartRow As Long
    plStartRow = 0
    Dim i As Long, kw As Variant, lCol As Long
    For i = dataStartRow To UBound(dataArray, 1)
        For lCol = 1 To maxLabelCol
            Dim cellText As String
            cellText = CStr(dataArray(i, lCol))
            If Len(cellText) > 0 Then
                Dim normText As String
                normText = CleanSearchText(cellText)
                For Each kw In plSectionKeywords
                    If InStr(1, normText, kw, vbTextCompare) > 0 Then
                        plStartRow = i
                        Exit For
                    End If
                Next kw
                If plStartRow > 0 Then Exit For
            End If
        Next lCol
        If plStartRow > 0 Then Exit For
    Next i

    ' Search ranges
    Dim bsSearchEnd As Long
    Dim plSearchStart As Long
    Dim hasPLSection As Boolean

    ' BUG4 FIX: plStartRow must be strictly greater than dataStartRow to constitute a real BS section above it.
    ' If plStartRow <= dataStartRow, the "PL section" starts at the very beginning — no BS section exists above.
    If plStartRow > 0 And plStartRow <= dataStartRow Then
        plStartRow = 0
    End If
    hasPLSection = (plStartRow > 0)

    If hasPLSection Then
        bsSearchEnd = plStartRow - 1
        plSearchStart = plStartRow
    Else
        bsSearchEnd = UBound(dataArray, 1)
        plSearchStart = dataStartRow
    End If

    '-----------------------------------------------------------
    ' 6. Search for Base rows
    '-----------------------------------------------------------
    Dim bsBaseRow As Long: bsBaseRow = 0
    Dim bsType As String
    Dim foundBS As Boolean

    Dim plBaseRow As Long: plBaseRow = 0
    Dim plType As String
    Dim foundPL As Boolean

    Dim cfBaseRow As Long: cfBaseRow = 0
    Dim cfType As String
    Dim foundCF As Boolean

    If reportChoice = 2 Or reportChoice = 3 Then
        foundBS = FindBaseRowInRange(dataArray, dataStartRow, bsSearchEnd, bsKeywords, bsExclude, bsBaseRow, bsType, maxLabelCol)
    End If

    If reportChoice = 1 Or reportChoice = 3 Then
        foundPL = FindBaseRowInRange(dataArray, plSearchStart, UBound(dataArray, 1), plKeywords, plExclude, plBaseRow, plType, maxLabelCol)
        ' Lenient second pass if not found
        If Not foundPL Then
            Dim plKeywordsLenient As Variant
            plKeywordsLenient = Array("revenue", "net sales", "net revenue", "doanh thu thuan", "doanh thu")
            Dim plExcludeLenient As Variant
            plExcludeLenient = Array("deferred", "chua thuc hien", "tai chinh", "financial", "financial income")
            foundPL = FindBaseRowInRange(dataArray, plSearchStart, UBound(dataArray, 1), plKeywordsLenient, plExcludeLenient, plBaseRow, plType, maxLabelCol)
        End If
    End If

    If reportChoice = 4 Then
        foundCF = FindBaseRowInRange(dataArray, dataStartRow, UBound(dataArray, 1), cfKeywords, cfExclude, cfBaseRow, cfType, maxLabelCol)
        If Not foundCF Then
            Dim cfKeywordsLenient As Variant
            cfKeywordsLenient = Array("operating activities", "business activities", "kinh doanh")
            Dim cfExcludeLenient As Variant
            cfExcludeLenient = Array("investing", "financing", "dau tu", "tai chinh", "indirect", "direct")
            foundCF = FindBaseRowInRange(dataArray, dataStartRow, UBound(dataArray, 1), cfKeywordsLenient, cfExcludeLenient, cfBaseRow, cfType, maxLabelCol)
        End If
    End If

    '-----------------------------------------------------------
    ' 7. Base Confirmation & Manual Fallback Handling
    '-----------------------------------------------------------
    Dim useDualBase As Boolean
    useDualBase = False
    Dim finalBaseRow As Long
    Dim finalBaseType As String
    Dim manualInput As String
    Dim rowNum As Long

    Select Case reportChoice
        Case 3 ' AUTO (Dual mode)
            If hasPLSection Then
                ' If BS not found, prompt specifically for BS
                If Not foundBS Then
                    manualInput = InputBox("Auto mode detected P&L section at row " & plStartRow & "," & vbNewLine & _
                                           "but could NOT automatically find TOTAL ASSETS for Balance Sheet." & vbNewLine & vbNewLine & _
                                           "Please enter row number or keyword for TOTAL ASSETS (Rows 2 to " & (plStartRow - 1) & "):", _
                                           "Select Balance Sheet Base", "Total Assets")
                    If manualInput = "" Then Exit Sub
                    bsBaseRow = ResolveRowFromInput(dataArray, dataStartRow, plStartRow - 1, manualInput, maxLabelCol)
                    If bsBaseRow > 0 Then foundBS = True
                End If

                ' If PL not found, prompt specifically for PL
                If Not foundPL Then
                    manualInput = InputBox("Auto mode could NOT automatically find REVENUE / NET SALES for Income Statement." & vbNewLine & vbNewLine & _
                                           "Please enter row number or keyword for REVENUE (Rows " & plStartRow & " to " & UBound(dataArray, 1) & "):", _
                                           "Select Income Statement Base", "Revenue")
                    If manualInput = "" Then Exit Sub
                    plBaseRow = ResolveRowFromInput(dataArray, plStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                    If plBaseRow > 0 Then foundPL = True
                End If

                If foundBS And foundPL Then
                    useDualBase = True
                    Dim confirmDual As String
                    confirmDual = "Auto Dual Base Confirmed:" & vbNewLine & vbNewLine & _
                                  "1. Balance Sheet Base (Rows 2 to " & (plStartRow - 1) & "):" & vbNewLine & _
                                  "   Row " & bsBaseRow & ": " & dataArray(bsBaseRow, 1) & vbNewLine & vbNewLine & _
                                  "2. Income Statement Base (Rows " & plStartRow & " to " & UBound(dataArray, 1) & "):" & vbNewLine & _
                                  "   Row " & plBaseRow & ": " & dataArray(plBaseRow, 1) & vbNewLine & vbNewLine & _
                                  "Click Yes to proceed, or No to choose single base."
                    If MsgBox(confirmDual, vbYesNo + vbQuestion, "Confirm Dual Base") = vbNo Then
                        useDualBase = False
                        manualInput = InputBox("Enter keyword or row number for whole report base:", "Single Base Entry", "Total Assets")
                        If manualInput = "" Then Exit Sub
                        finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                        If finalBaseRow = 0 Then
                            MsgBox "Invalid base selection. Operation cancelled.", vbExclamation
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox "Could not establish dual bases. Switching to manual single base.", vbInformation
                    manualInput = InputBox("Enter keyword or row number for single base:", "Manual Base Entry", "Total Assets")
                    If manualInput = "" Then Exit Sub
                    finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                    If finalBaseRow = 0 Then Exit Sub
                End If
            Else
                ' BUG5 FIX: No PL section separator detected → Ask user to clarify report type
                ' before blindly searching the entire range, to avoid BS rows being matched as PL base.
                If foundBS And Not foundPL Then
                    ' Only BS base found — confirm and use it as single base
                    finalBaseRow = bsBaseRow
                    finalBaseType = "Total Assets (BS)"
                    If MsgBox("No Income Statement section detected." & vbNewLine & _
                              "Using Total Assets (Row " & bsBaseRow & ") as single base." & vbNewLine & vbNewLine & _
                              "Proceed?", vbYesNo + vbQuestion, "Confirm Single Base") = vbNo Then Exit Sub
                ElseIf foundPL And Not foundBS Then
                    ' Only PL base found — use it
                    finalBaseRow = plBaseRow
                    finalBaseType = "Revenue (PL)"
                    If MsgBox("No Balance Sheet section detected." & vbNewLine & _
                              "Using Revenue (Row " & plBaseRow & ") as single base." & vbNewLine & vbNewLine & _
                              "Proceed?", vbYesNo + vbQuestion, "Confirm Single Base") = vbNo Then Exit Sub
                ElseIf foundBS And foundPL Then
                    ' Both found on same range — ask user which to use
                    Dim pickChoice As String
                    pickChoice = InputBox("Both Total Assets and Revenue found in the same range." & vbNewLine & _
                                          "Enter row number for the base to use (or leave blank to cancel):", _
                                          "Select Base", bsBaseRow)
                    If pickChoice = "" Then Exit Sub
                    finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), pickChoice, maxLabelCol)
                    If finalBaseRow = 0 Then Exit Sub
                Else
                    manualInput = InputBox("No section separator found and no standard base detected." & vbNewLine & _
                                           "Enter keyword (e.g., 'Total Assets' or 'Net Revenue') or row number:", "Manual Base Entry", "Total Assets")
                    If manualInput = "" Then Exit Sub
                    finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                    If finalBaseRow = 0 Then Exit Sub
                End If
            End If

        Case 1 ' Income Statement (PL single base)
            If foundPL Then
                finalBaseRow = plBaseRow
                finalBaseType = "Revenue"
            Else
                manualInput = InputBox("Could not detect Revenue automatically." & vbNewLine & _
                                       "Enter keyword or row number for Revenue base:", "Manual PL Base", "Revenue")
                If manualInput = "" Then Exit Sub
                finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                If finalBaseRow = 0 Then Exit Sub
            End If

        Case 2 ' Balance Sheet (BS single base)
            If foundBS Then
                finalBaseRow = bsBaseRow
                finalBaseType = "Total Assets"
            Else
                manualInput = InputBox("Could not detect Total Assets automatically." & vbNewLine & _
                                       "Enter keyword or row number for Total Assets base:", "Manual BS Base", "Total Assets")
                If manualInput = "" Then Exit Sub
                finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                If finalBaseRow = 0 Then Exit Sub
            End If

        Case 4 ' Cash Flow Statement (Net Operating Cash Flow base)
            If foundCF Then
                finalBaseRow = cfBaseRow
                finalBaseType = "Net Cash Flow from Operating Activities"
            Else
                manualInput = InputBox("Could not detect Net Operating Cash Flow automatically." & vbNewLine & _
                                       "Enter keyword or row number for Cash Flow base:", "Manual Cash Flow Base", "Operating Activities")
                If manualInput = "" Then Exit Sub
                finalBaseRow = ResolveRowFromInput(dataArray, dataStartRow, UBound(dataArray, 1), manualInput, maxLabelCol)
                If finalBaseRow = 0 Then Exit Sub
                finalBaseType = "Net Operating Cash Flow"
            End If
    End Select

    ' Single base confirmation if not in dual mode
    If Not useDualBase Then
        Dim confirmSingle As String
        confirmSingle = "Using Base Row " & finalBaseRow & ": " & dataArray(finalBaseRow, 1) & vbNewLine & _
                        "Proceed with Common Size calculation?"
        If MsgBox(confirmSingle, vbYesNo + vbQuestion, "Confirm Base") = vbNo Then
            Exit Sub
        End If
    End If

    '-----------------------------------------------------------
    ' 8. Calculate Common Size (Excel Dynamic Formulas)
    '-----------------------------------------------------------
    Dim outputArray As Variant
    outputArray = dataArray

    Dim cellVal As Variant
    Dim cellRef As String, baseCellRef As String
    Dim targetBaseRow As Long
    Dim DQ As String
    DQ = Chr(34)
    Dim safeWsName As String
    safeWsName = "'" & Replace(ws.Name, "'", "''") & "'!"

    For col = dataStartCol To UBound(dataArray, 2)
        For i = dataStartRow To UBound(dataArray, 1)
            cellVal = dataArray(i, col)
            If IsNumeric(cellVal) And Len(Trim(CStr(cellVal))) > 0 Then
                ' Determine target base row (Balance Sheet or Income Statement)
                If useDualBase Then
                    If i < plStartRow Then
                        targetBaseRow = bsBaseRow
                    Else
                        targetBaseRow = plBaseRow
                    End If
                Else
                    targetBaseRow = finalBaseRow
                End If

                If targetBaseRow > 0 Then
                    ' BUG1 FIX: Use absolute addresses (True, True) so formulas don't drift when copied
                    cellRef = safeWsName & ws.Cells(firstRow + i - 1, firstCol + col - 1).Address(True, True)
                    baseCellRef = safeWsName & ws.Cells(firstRow + targetBaseRow - 1, firstCol + col - 1).Address(True, True)
                    outputArray(i, col) = "=IFERROR(" & cellRef & "/" & baseCellRef & ", " & DQ & DQ & ")"
                Else
                    outputArray(i, col) = cellVal
                End If
            Else
                outputArray(i, col) = cellVal
            End If
        Next i
    Next col

    '-----------------------------------------------------------
    ' 9. Ask destination and paste (as Formulas)
    '-----------------------------------------------------------
    Dim destRange As Range
    On Error Resume Next
    Set destRange = Application.InputBox( _
        "Select top-left cell to paste Common Size results:", _
        "Destination Cell", Type:=8)
    On Error GoTo ErrorHandler
    If destRange Is Nothing Then Exit Sub
    Set destRange = destRange.Cells(1, 1)

    Dim rowsCount As Long, colsCount As Long
    rowsCount = UBound(outputArray, 1) - LBound(outputArray, 1) + 1
    colsCount = UBound(outputArray, 2) - LBound(outputArray, 2) + 1

    ' BUG11 FIX: Warn if destination overlaps with source (would overwrite raw data with formulas)
    If Not Intersect(destRange.Resize(rowsCount, colsCount), rng) Is Nothing Then
        If MsgBox("WARNING: Destination overlaps with source data!" & vbNewLine & _
                  "Proceeding will overwrite original values with formulas." & vbNewLine & vbNewLine & _
                  "Continue anyway?", vbYesNo + vbCritical, "Overlap Warning") = vbNo Then
            GoTo CleanExit
        End If
    End If

    destRange.Resize(rowsCount, colsCount).Formula = outputArray

    ' BUG3 FIX: Apply percentage format only to data columns in destination
    ' col index in outputArray is 1-based; dataStartCol is the first numeric data column
    ' offset within destRange = col - 1 (since outputArray col 1 = destRange col 1)
    For col = dataStartCol To UBound(dataArray, 2)
        ' Row 1 in outputArray = header row; skip it (Offset by 1 row)
        destRange.Offset(1, col - 1).Resize(rowsCount - 1, 1).NumberFormat = "0.00%"
    Next col

    Dim resultMsg As String
    If useDualBase Then
        resultMsg = "Common Size completed (Dual Base):" & vbNewLine & _
                    "• Balance Sheet base: Row " & bsBaseRow & " (" & dataArray(bsBaseRow, 1) & ")" & vbNewLine & _
                    "• Income Statement base: Row " & plBaseRow & " (" & dataArray(plBaseRow, 1) & ")"
    Else
        resultMsg = "Common Size completed:" & vbNewLine & _
                    "• Base: Row " & finalBaseRow & " (" & dataArray(finalBaseRow, 1) & ")"
    End If
    MsgBox resultMsg & vbNewLine & vbNewLine & "Results pasted at " & destRange.Address, vbInformation
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------
' Helper: Search Base Row within row range
' Searches across description columns (Col 1 to maxLabelCol, e.g. VN & EN)
'------------------------------------------------------------------
Function FindBaseRowInRange(dataArr As Variant, startRow As Long, endRow As Long, _
                            keywords As Variant, exclusions As Variant, _
                            ByRef baseRow As Long, ByRef baseType As String, _
                            Optional maxLabelCol As Long = 1) As Boolean
    Dim i As Long, kw As Variant, excl As Variant
    Dim rawText As String, normText As String
    Dim isExcluded As Boolean
    Dim colIdx As Long

    If startRow > endRow Or startRow < 1 Or endRow > UBound(dataArr, 1) Then
        FindBaseRowInRange = False
        Exit Function
    End If

    If maxLabelCol < 1 Then maxLabelCol = 1
    If maxLabelCol > 2 Then maxLabelCol = 2 ' Primary text scanning on Col 1 (VN) and Col 2 (EN)

    ' Step 1: Scan keywords in order of priority (exact/substring in clean text)
    For Each kw In keywords
        For i = startRow To endRow
            For colIdx = 1 To maxLabelCol
                rawText = CStr(dataArr(i, colIdx))
                If Len(rawText) > 0 Then
                    normText = CleanSearchText(rawText)
                    If InStr(1, normText, kw, vbTextCompare) > 0 Then
                        ' Check exclusions across both label columns of this row
                        isExcluded = False
                        Dim eCol As Long
                        For eCol = 1 To maxLabelCol
                            Dim eText As String
                            eText = CleanSearchText(CStr(dataArr(i, eCol)))
                            For Each excl In exclusions
                                If InStr(1, eText, excl, vbTextCompare) > 0 Then
                                    isExcluded = True
                                    Exit For
                                End If
                            Next excl
                            If isExcluded Then Exit For
                        Next eCol

                        If Not isExcluded Then
                            baseRow = i
                            baseType = "Matched keyword: " & kw
                            FindBaseRowInRange = True
                            Exit Function
                        End If
                    End If
                End If
            Next colIdx
        Next i
    Next kw

    FindBaseRowInRange = False
End Function

'------------------------------------------------------------------
' Helper: Resolve manual row input (either row number or keyword)
' Searches across description columns (Col 1 to maxLabelCol)
'------------------------------------------------------------------
Function ResolveRowFromInput(dataArr As Variant, minRow As Long, maxRow As Long, inputVal As String, _
                            Optional maxLabelCol As Long = 1) As Long
    Dim i As Long, cleanVal As String, colIdx As Long
    inputVal = Trim(inputVal)
    If Len(inputVal) = 0 Then
        ResolveRowFromInput = 0
        Exit Function
    End If

    If IsNumeric(inputVal) Then
        Dim r As Long
        r = CLng(inputVal)
        If r >= minRow And r <= maxRow Then
            ResolveRowFromInput = r
            Exit Function
        Else
            MsgBox "Row number " & r & " is out of range (" & minRow & " - " & maxRow & ").", vbExclamation
            ResolveRowFromInput = 0
            Exit Function
        End If
    End If

    If maxLabelCol < 1 Then maxLabelCol = 1
    If maxLabelCol > 2 Then maxLabelCol = 2

    cleanVal = CleanSearchText(inputVal)
    For i = minRow To maxRow
        For colIdx = 1 To maxLabelCol
            Dim rowText As String
            rowText = CStr(dataArr(i, colIdx))
            If InStr(1, CleanSearchText(rowText), cleanVal, vbTextCompare) > 0 Then
                ResolveRowFromInput = i
                Exit Function
            End If
        Next colIdx
    Next i

    MsgBox "Keyword '" & inputVal & "' not found in search range (rows " & minRow & " to " & maxRow & ").", vbExclamation
    ResolveRowFromInput = 0
End Function

'------------------------------------------------------------------
' Helper: Normalize text (lowercase, remove Vietnamese accents & symbols)
'------------------------------------------------------------------
Function CleanSearchText(ByVal txt As String) As String
    Dim result As String
    result = LCase(Trim(txt))

    ' Replace common punctuation to make matching resilient
    result = Replace(result, "_", " ")
    result = Replace(result, "-", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, ":", " ")
    result = Replace(result, "(", " ")
    result = Replace(result, ")", " ")
    result = Replace(result, "=", " ")
    result = Replace(result, "+", " ")

    ' Convert Vietnamese accented Unicode characters safely by character code
    Dim i As Long, code As Long, ch As String, cleanCh As String
    Dim outStr As String
    outStr = ""

    For i = 1 To Len(result)
        ch = Mid(result, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        Select Case code
            ' a / A variants
            Case 224, 225, 226, 227, 259, 7840, 7841, 7842, 7843, 7844, 7845, 7846, 7847, 7848, 7849, 7850, 7851, 7852, 7853, 7854, 7855, 7856, 7857, 7858, 7859, 7860, 7861, 7862, 7863
                cleanCh = "a"
            ' e / E variants
            Case 232, 233, 234, 7864, 7865, 7866, 7867, 7868, 7869, 7870, 7871, 7872, 7873, 7874, 7875, 7876, 7877, 7878, 7879
                cleanCh = "e"
            ' i / I variants
            Case 236, 237, 297, 7880, 7881, 7882, 7883
                cleanCh = "i"
            ' o / O variants
            Case 242, 243, 244, 245, 417, 7884, 7885, 7886, 7887, 7888, 7889, 7890, 7891, 7892, 7893, 7894, 7895, 7896, 7897, 7898, 7899, 7900, 7901, 7902, 7903, 7904, 7905, 7906, 7907
                cleanCh = "o"
            ' u / U variants
            Case 249, 250, 361, 432, 7908, 7909, 7910, 7911, 7912, 7913, 7914, 7915, 7916, 7917, 7918, 7919, 7920, 7921
                cleanCh = "u"
            ' y / Y variants
            Case 253, 7922, 7923, 7924, 7925, 7926, 7927, 7928, 7929
                cleanCh = "y"
            ' d / D variants
            Case 273, 272
                cleanCh = "d"
            Case Else
                cleanCh = ch
        End Select
        outStr = outStr & cleanCh
    Next i

    ' Collapse multiple spaces
    Do While InStr(outStr, "  ") > 0
        outStr = Replace(outStr, "  ", " ")
    Loop

    CleanSearchText = Trim(outStr)
End Function


