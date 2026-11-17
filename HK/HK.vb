Option Explicit

'================================================================
' Module-level state
'================================================================
Private m_wsNew As Worksheet, m_wsHK As Worksheet
Private m_colLetter As String
Private m_inputMonth As String
Private m_r As Long
Private m_boldRows As String

' Vi tri BS
Private m_caStart As Long, m_caTotal As Long
Private m_peStart As Long, m_peTotal As Long
Private m_ltStart As Long, m_ltTotal As Long
Private m_assetsTotal As Long
Private m_clStart As Long, m_clTotal As Long
Private m_ltlStart As Long, m_ltlTotal As Long
Private m_liabTotal As Long
Private m_capStart As Long, m_capTotal As Long
Private m_liabCapTotal As Long

' Vi tri PL
Private m_revStart As Long, m_revTotal As Long
Private m_cosStart As Long, m_cosTotal As Long
Private m_grossProfitRow As Long
Private m_expStart As Long, m_expTotal As Long
Private m_netIncomeRow As Long
Private m_checkPLRow As Long
Private m_reRow As Long
Private m_ebitdaRow As Long

' Vi tri dong dac biet cho EBITDA
Private m_rowOtherBankInt As Long, m_rowOtherInvInt As Long, m_rowOtherLoanInt As Long
Private m_rowIntBank As Long, m_rowIntLexus As Long
Private m_rowDepOE As Long, m_rowDepCE As Long, m_rowDepMV As Long, m_rowDepSS As Long
Private m_rowTaxation As Long

'================================================================
' MAIN
'================================================================
Sub Create004HK()
    Dim wb As Workbook
    Dim found As Range
    Dim targetCol As Long

    Set wb = ThisWorkbook

    m_inputMonth = InputBox("Enter the month exactly as in HK header (e.g. July 2026):", _
                            "Month Selection", "July 2026")
    If m_inputMonth = "" Then Exit Sub

    On Error Resume Next
    Set m_wsHK = wb.Sheets("HK")
    On Error GoTo 0
    If m_wsHK Is Nothing Then
        MsgBox "Khong tim thay sheet 'HK'.", vbCritical
        Exit Sub
    End If

    Set found = m_wsHK.Rows(1).Find(What:=m_inputMonth, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "Khong tim thay thang '" & m_inputMonth & "' o dong 1 cua sheet 'HK'.", vbCritical
        Exit Sub
    End If
    targetCol = found.Column
    m_colLetter = Left(m_wsHK.Cells(1, targetCol).Address(False, False), _
                       Len(m_wsHK.Cells(1, targetCol).Address(False, False)) - 1)

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("004-HK").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set m_wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    m_wsNew.Name = "004-HK"

    Application.ScreenUpdating = False

    m_r = 2
    m_boldRows = ""

    ' === BALANCE SHEET ===
    BS_Header
    BS_CurrentAssets
    BS_PropertyEquipment
    BS_LongTermAssets
    BS_TotalAssets
    BS_Header2
    BS_CurrentLiabilities
    BS_LongTermLiabilities
    BS_TotalLiabilities
    BS_Capital
    BS_TotalLiabCap

    ' === PROFIT AND LOSS ===
    PL_Header
    PL_Revenues
    PL_CostOfSales
    PL_GrossProfit
    PL_Expenses_Part1
    PL_Expenses_Part2
    PL_Expenses_Part3
    PL_Expenses_Part4
    PL_Expenses_Total
    PL_NetIncome
    PL_CheckRE_EBITDA

    ApplyFormatting

    Application.ScreenUpdating = True

    MsgBox "Da tao sheet 004-HK cho thang " & m_inputMonth & " thanh cong!" & vbCrLf & _
           "Moi cong thuc tu dong cap nhat khi ban them/bot dong.", vbInformation
End Sub

'================================================================
' BALANCE SHEET
'================================================================
Private Sub BS_Header()
    With m_wsNew
        .Range("A" & m_r).Value = "BALANCE SHEET"
        m_boldRows = m_boldRows & "," & m_r
        m_r = m_r + 1
        .Range("B" & m_r).Value = m_inputMonth
        m_r = m_r + 1
        .Range("A" & m_r).Value = "ASSETS"
        m_boldRows = m_boldRows & "," & m_r
        m_r = m_r + 2
    End With
End Sub

Private Sub BS_CurrentAssets()
    m_wsNew.Range("A" & m_r).Value = "Current Assets"
    m_caStart = m_r
    m_boldRows = m_boldRows & "," & m_caStart
    m_r = m_r + 1

    ProcessBS_Chunk Array( _
        "HSBC S/A HKD-411-781008-838", _
        "HSBC C/A HKD-411-781008-001", _
        "HSBC USD S/C-411-781008-838", _
        "HSBC CAD S/C-411-781008-838", _
        "C/A - Dan on USA", _
        "C/A - AN DIE FOODS JSC", _
        "C/A - Dan D Pak Co.,Ltd", _
        "C/A - Dan On Foods Corporation", _
        "C/A - Danway Holding", _
        "Inventories - Finished Goods")

    ProcessBS_Chunk Array( _
        "Accounts Receivable", _
        "Other accounts receivable", _
        "Deposit paid - non trade", _
        "Deposit paid - ERP system", _
        "Prepayments (Rent)", _
        "Prepayments")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_caStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Current Assets"
    m_caTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_PropertyEquipment()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Property and Equipment"
    m_peStart = m_r
    m_boldRows = m_boldRows & "," & m_peStart
    m_r = m_r + 1

    ProcessBS_Chunk Array( _
        "Furnitures, Fixtures and Fitti", _
        "Depreciation - FF&F", _
        "Office Equipments", _
        "Depreciation - OE", _
        "Computer Equipment", _
        "Depreciation - CE", _
        "Motor Vehicles", _
        "Depreciation - MV", _
        "Software & System", _
        "Depreciation - SS")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_peStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Property and Equipment"
    m_peTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_LongTermAssets()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Long Term Assets"
    m_ltStart = m_r
    m_boldRows = m_boldRows & "," & m_ltStart
    m_r = m_r + 1

    WriteBS m_wsNew, m_wsHK, m_colLetter, m_r, "Long Term Investment"
    m_r = m_r + 2

    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_ltStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Long Term  Assets"
    m_ltTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_TotalAssets()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(B" & m_caTotal & "+B" & m_peTotal & "+B" & m_ltTotal & ", 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Assets"
    m_assetsTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_Header2()
    m_r = m_r + 3
    m_wsNew.Range("A" & m_r).Value = "LIABILITIES AND CAPITAL"
    m_boldRows = m_boldRows & "," & m_r
    m_r = m_r + 1
End Sub

Private Sub BS_CurrentLiabilities()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Current Liabilities"
    m_clStart = m_r
    m_r = m_r + 1

    ProcessBS_Chunk Array( _
        "Accounts Payable", _
        "Other payable", _
        "Accrued expenses")

    WriteBS m_wsNew, m_wsHK, m_colLetter, m_r, "Provision for taxation", True
    m_r = m_r + 1

    ProcessBS_Chunk Array( _
        "Deposits received-Manulife", _
        "Loan from Lexus")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_clStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Current Liabilities"
    m_clTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_LongTermLiabilities()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Long-Term Liabilities"
    m_ltlStart = m_r
    m_boldRows = m_boldRows & "," & m_ltlStart
    m_r = m_r + 2

    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_ltlStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Long-Term Liabilities"
    m_ltlTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_TotalLiabilities()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=-(ROUND(-B" & m_clTotal & "+-B" & m_ltlTotal & ", 5))"
    m_wsNew.Range("A" & m_r).Value = "Total Liabilities"
    m_liabTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_Capital()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Capital"
    m_capStart = m_r
    m_boldRows = m_boldRows & "," & m_capStart
    m_r = m_r + 1

    ProcessBS_Chunk Array( _
        "Issued Share Capital", _
        "Retained earnings", _
        "Dividend", _
        "Net Income")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_capStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Capital"
    m_capTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub BS_TotalLiabCap()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=-(ROUND(-B" & m_liabTotal & "+-B" & m_capTotal & ", 5))"
    m_wsNew.Range("A" & m_r).Value = "Total Liabilities & Capital"
    m_liabCapTotal = m_r
    m_r = m_r + 1

    m_wsNew.Range("B" & m_r).Formula = "=B" & m_liabCapTotal & "-B" & m_assetsTotal
    m_wsNew.Range("A" & m_r).Value = "Check"
    m_r = m_r + 1
End Sub

'================================================================
' PROFIT AND LOSS
'================================================================
Private Sub PL_Header()
    m_r = m_r + 2
    m_wsNew.Range("A" & m_r).Value = "PROFIT AND LOSS"
    m_boldRows = m_boldRows & "," & m_r
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Value = m_inputMonth
    m_r = m_r + 1
End Sub

Private Sub PL_Revenues()
    m_wsNew.Range("A" & m_r).Value = "Revenues"
    m_revStart = m_r
    m_r = m_r + 1

    ProcessPLRev_Chunk Array( _
        "Income - Sales", _
        "Income - Freight", _
        "Sales - return", _
        "Sales - Claim", _
        "Discount Allowed - Sales", _
        "Other income - Bank interest", _
        "Other income - Inv.interest", _
        "Other income - Loan interest", _
        "Other income - Subsidies", _
        "Other income - Others", _
        "Gain on Disposal of fixed asse")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_revStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Revenues"
    m_revTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub PL_CostOfSales()
    m_r = m_r + 2
    m_wsNew.Range("A" & m_r).Value = "Cost of Sales"
    m_cosStart = m_r
    m_r = m_r + 1

    ProcessPL_Chunk Array( _
        "Opening stock", _
        "Purchase", _
        "Purchase Return", _
        "Purchase Discount", _
        "DCOS - Transportation", _
        "DCOS - freight charges", _
        "DCOS - Others", _
        "Less: closing stock")

    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_cosStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Cost of Sales"
    m_cosTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub PL_GrossProfit()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=-(ROUND(-B" & m_revTotal & "+B" & m_cosTotal & ", 5))"
    m_wsNew.Range("A" & m_r).Value = "Gross Profit"
    m_grossProfitRow = m_r
    m_r = m_r + 1
End Sub

Private Sub PL_Expenses_Part1()
    m_r = m_r + 1
    m_wsNew.Range("A" & m_r).Value = "Expenses"
    m_expStart = m_r
    m_r = m_r + 1

    ProcessPLExp_Chunk Array( _
        "Advertising and Promotion Fee", _
        "Commission", _
        "Entertainments", _
        "Exhibition", _
        "Trucking charges", _
        "Declaration & Certificate Fee", _
        "Freight Charges", _
        "Chain Store commission& rebate", _
        "Other selling expenses", _
        "Bank Charges", _
        "Exchange difference", _
        "Interest Expenses - Bank", _
        "Interest Expenses - Lexus", _
        "Depreciation - OE", _
        "Depreciation - CE", _
        "Depreciation - MV", _
        "Depreciation-SS")
End Sub

Private Sub PL_Expenses_Part2()
    ProcessPLExp_Chunk Array( _
        "Accounting Services Fee", _
        "Audit Fee", _
        "Building Cost - Rental", _
        "Building Cost - Storage", _
        "Business Registration Fees", _
        "Cleaning Charges", _
        "Communication  - Tel, fax, IDD", _
        "Computer Supplies", _
        "Chains store  Dis. Allowance", _
        "PNS Distribution Allowance", _
        "Distribution Allowance", _
        "Governments Charges", _
        "Insurance", _
        "Legal and professional fees", _
        "Loss on sale of investment", _
        "Listing fees")
End Sub

Private Sub PL_Expenses_Part3()
    ProcessPLExp_Chunk Array( _
        "Gate charges", _
        "Motor Vehicles Expenses", _
        "Motor Vehicles Expenses - Gas", _
        "Motor Vehicles Expense-Parking", _
        "Office Supplies", _
        "Postage and Courier Serivces", _
        "Printing and Stationery", _
        "Repair and Maintenance fee", _
        "Salary - Full Time", _
        "MPF (Employer)", _
        "Salary - Partime / freelance", _
        "Staff allowance", _
        "Staff allowance-Meal", _
        "Registration fee", _
        "Subscription fee", _
        "Travelling - Local")
End Sub

Private Sub PL_Expenses_Part4()
    ProcessPLExp_Chunk Array( _
        "Travelling - Overseas", _
        "Travelling - Hotel charges", _
        "Utilities - Electricity", _
        "Member fees", _
        "Product development fees", _
        "Sample charges", _
        "Depreciation - FF&F", _
        "Discount Allowed", _
        "Staff Bonus", _
        "Taxation", _
        "Utilities - Water", _
        "Terminal Handling Charges", _
        "Other Administrative Expenses")
End Sub

Private Sub PL_Expenses_Total()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=ROUND(SUBTOTAL(9, B" & m_expStart & ":B" & (m_r - 1) & "), 5)"
    m_wsNew.Range("A" & m_r).Value = "Total Expenses"
    m_expTotal = m_r
    m_r = m_r + 1
End Sub

Private Sub PL_NetIncome()
    m_r = m_r + 1
    m_wsNew.Range("B" & m_r).Formula = _
        "=-(ROUND(-B" & m_grossProfitRow & "+B" & m_expTotal & ", 5))"
    m_wsNew.Range("A" & m_r).Value = "Net Income"
    m_netIncomeRow = m_r
    m_r = m_r + 1
End Sub

Private Sub PL_CheckRE_EBITDA()
    m_checkPLRow = m_netIncomeRow + 1
    m_reRow = m_netIncomeRow + 3
    m_ebitdaRow = m_netIncomeRow + 5

    With m_wsNew
        ' Check
        .Range("B" & m_checkPLRow).Formula = _
            "=B" & m_reRow & "+B" & m_netIncomeRow & "-B" & m_capTotal
        .Range("A" & m_checkPLRow).Value = "Check"

        ' RE dau nam
        .Range("A" & m_reRow).Value = "At 31 Dec 25 - RE"
        .Range("B" & m_reRow).Value = 9853375.61

        ' EBITDA
        .Range("A" & m_ebitdaRow).Value = "EBITDA"
        .Range("B" & m_ebitdaRow).Formula = _
            "=B" & m_netIncomeRow & "+B" & m_rowTaxation & _
            "+SUM(B" & m_rowDepOE & ":B" & m_rowDepSS & ")" & _
            "+SUM(B" & m_rowIntBank & ":B" & m_rowIntLexus & ")" & _
            "-SUM(B" & m_rowOtherBankInt & ":B" & m_rowOtherLoanInt & ")"
    End With

    m_boldRows = m_boldRows & "," & m_ebitdaRow
End Sub

'================================================================
' HELPERS — nho, gon, khong vuot gioi han line-continuation
'================================================================
Private Sub ProcessBS_Chunk(arr As Variant)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        WriteBS m_wsNew, m_wsHK, m_colLetter, m_r, CStr(arr(i))
        m_r = m_r + 1
    Next i
End Sub

Private Sub ProcessPL_Chunk(arr As Variant)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        WritePL m_wsNew, m_wsHK, m_colLetter, m_r, CStr(arr(i))
        m_r = m_r + 1
    Next i
End Sub

Private Sub ProcessPLRev_Chunk(arr As Variant)
    Dim i As Long, lbl As String
    For i = LBound(arr) To UBound(arr)
        lbl = CStr(arr(i))
        WritePL m_wsNew, m_wsHK, m_colLetter, m_r, lbl
        Select Case lbl
            Case "Other income - Bank interest": m_rowOtherBankInt = m_r
            Case "Other income - Inv.interest":  m_rowOtherInvInt = m_r
            Case "Other income - Loan interest": m_rowOtherLoanInt = m_r
        End Select
        m_r = m_r + 1
    Next i
End Sub

Private Sub ProcessPLExp_Chunk(arr As Variant)
    Dim i As Long, lbl As String
    For i = LBound(arr) To UBound(arr)
        lbl = CStr(arr(i))
        WritePL m_wsNew, m_wsHK, m_colLetter, m_r, lbl
        Select Case lbl
            Case "Interest Expenses - Bank":  m_rowIntBank = m_r
            Case "Interest Expenses - Lexus": m_rowIntLexus = m_r
            Case "Depreciation - OE":         m_rowDepOE = m_r
            Case "Depreciation - CE":         m_rowDepCE = m_r
            Case "Depreciation - MV":         m_rowDepMV = m_r
            Case "Depreciation-SS":           m_rowDepSS = m_r
            Case "Taxation":                  m_rowTaxation = m_r
        End Select
        m_r = m_r + 1
    Next i
End Sub

Private Sub WriteBS(wsNew As Worksheet, wsHK As Worksheet, ByVal colLetter As String, _
                    ByVal r As Long, ByVal label As String, _
                    Optional ByVal negate As Boolean = False)
    wsNew.Cells(r, "A").Value = label
    Dim f As String
    f = "SUMIFS(HK!" & colLetter & "$81:" & colLetter & "$136,HK!$A$81:$A$136,""" & label & """)"
    If negate Then f = "-" & f
    wsNew.Cells(r, "B").Formula = "=" & f
End Sub

Private Sub WritePL(wsNew As Worksheet, wsHK As Worksheet, ByVal colLetter As String, _
                    ByVal r As Long, ByVal label As String)
    wsNew.Cells(r, "A").Value = label
    wsNew.Cells(r, "B").Formula = _
        "=SUMIFS(HK!" & colLetter & "$4:" & colLetter & "$69,HK!$A$4:$A$69,""" & label & """)"
End Sub

Private Sub ApplyFormatting()
    Dim arr() As String, k As Long

    With m_wsNew
        .Rows("2:3").Font.Bold = True
        If Len(m_boldRows) > 0 Then
            arr = Split(m_boldRows, ",")
            For k = LBound(arr) To UBound(arr)
                If Trim(arr(k)) <> "" Then
                    .Range("A" & CLng(Trim(arr(k)))).Font.Bold = True
                End If
            Next k
        End If
        .Columns("A:B").AutoFit
        .Activate
    End With
End Sub
