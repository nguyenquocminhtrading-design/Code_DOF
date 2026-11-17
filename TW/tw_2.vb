Option Explicit

Private Const MAX_ROW_003 As Long = 198   ' 003-TW: noi dung chinh khong duoc vuot qua hang nay
Private Const MAX_SCAN_COLS As Long = 18  ' quet toi da bao nhieu cot ben phai nhan de tim gia tri so

Sub Create003TW()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsPL As Worksheet, wsFmt As Worksheet
    Dim inputMonth As String
    Dim destRow As Long
    Dim rowOf As Object          ' Luu so dong (tren 003-TW) cua cac muc quan trong
    Dim bsLabels As Variant      ' Danh sach khoan muc BALANCE SHEET (co dinh)
    Dim plLabels As Variant      ' Danh sach khoan muc PROFIT AND LOSS (co dinh)
    Dim i As Long
    Dim lbl As String, cleanLbl As String
    Dim valRef As String
    Dim checkRow As Long, ebitdaRow As Long
    Dim bsStr As String, plStr As String

    Set wb = ThisWorkbook

    ' ---- Nhap thang ----
    inputMonth = InputBox("Enter the month (chi de ghi chu, vd 2026.07):", "Month Selection", "2026.07")
    If inputMonth = "" Then Exit Sub

    ' ---- Kiem tra sheet nguon ----
    On Error Resume Next
    Set wsBS = wb.Sheets("TW BS")
    Set wsPL = wb.Sheets("TW PL")
    Set wsFmt = wb.Sheets("TW_format")
    On Error GoTo 0
    If wsBS Is Nothing Or wsPL Is Nothing Or wsFmt Is Nothing Then
        MsgBox "Thieu sheet 'TW BS', 'TW PL' hoac 'TW_format'.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- Xoa sheet cu, tao moi ----
    On Error Resume Next
    wb.Sheets("003-TW").Delete
    On Error GoTo 0
    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = "003-TW"
    Application.DisplayAlerts = True

    ' ---- Ghi tieu de thang ----
    wsNew.Range("B3").Value = inputMonth

    ' ---- Khoi tao Dictionary luu so dong (tren 003-TW) cua cac muc ----
    Set rowOf = CreateObject("Scripting.Dictionary")
    rowOf.Add "Total Assets", 0
    rowOf.Add "Total Liabilities", 0
    rowOf.Add "Total Equity", 0
    rowOf.Add "Interest revenue", 0
    rowOf.Add "Interest expense", 0
    rowOf.Add "Every depreciation", 0
    rowOf.Add "Various amortization", 0
    rowOf.Add "Taxable income", 0
    rowOf.Add "Net income(loss)", 0
    ' Them cac muc can cho phan CHECK PL (truoc day bi tham chieu sai sang sheet nguon)
    rowOf.Add "Sales revenue", 0
    rowOf.Add "minus:Sales return", 0
    rowOf.Add "minus:Sales discounts", 0
    rowOf.Add "Net sales", 0
    rowOf.Add "Total cost of goods sold", 0
    rowOf.Add "Gross profit(loss)", 0
    rowOf.Add "Total operating expenses", 0
    rowOf.Add "Income (loss) from operations", 0

    ' ================================================================
    ' 1. DANH SACH KHOAN MUC BALANCE SHEET (co dinh) - dung Split
    ' ================================================================
    bsStr = "Current assets|Cash|Cash in banks-Hua|" & _
            "Cash in banks-HSBC|Deposit account|Notes receivable|Accounts|" & _
            "Merchandise|Other prepayments|Prepaid insurance|Prepaid rents|" & _
            "Temporary payments|Prepaid sales tax|Prepaid expenses|Overpaid sales tax|" & _
            "Advance payments|Should return the amount of tax to be paid|" & _
            "Long-term|Long-term investments|Total Long-term investment|" & _
            "Property, plant and equipment|Land|Buildings|" & _
            "Minus:Accumulated_ depreciatio n-buildings|Transportation|" & _
            "Minus : Accumulat|Other fixed| Minus : Accumulat Other fixed|" & _
            "Other non-current assets|Refundable|Have not amortized|" & _
            "Payment on behalf|Total Other non-current assets|Total Assets|" & _
            "Liabilities and Equity|Current liabilities|Accounts payable|Accrued expenses|" & _
            "Advances from|Temporary receipts|Other unearned revenue|" & _
            "Stockholders' current account|Current account|Bank loan|" & _
            "Sales tax payable|Receipts under|Total Current liabilities|" & _
            "Other non-current liabilities|Withhold taxes|Long-term liabilities|" & _
            "Long-term debt|Total Long-term liabilities|Total Liabilities|" & _
            "Capital|Total Capital|Retained earnings|" & _
            "Accumulated |Net income or loss|Total Retained earnings|" & _
            "Legal reserve|Total Equity|check"
    bsLabels = Split(bsStr, "|")

    ' ================================================================
    ' 2. DANH SACH KHOAN MUC PROFIT AND LOSS (co dinh) - dung Split
    ' ================================================================
    plStr = "PROFIT AND LOSS|Account Name|Operating revenue|Sales revenue|" & _
            "minus:Sales return|minus:Sales discounts|Net sales|" & _
            "Cost of goods sold|Merchandise inventory|Purchases|" & _
            "minus:Purchase discounts|minus:Purchase return|The goods are changed over to|" & _
            "minus:The goods producing|Adjustment debit of the|" & _
            "minus:Merchandise|Total cost of goods sold|" & _
            "Total operating cost|Rate of operating costs|Gross profit(loss)|" & _
            "Rate of gross profit|Operating expenses|Payroll expense|" & _
            "Payroll expense-The company carrying and-|Payroll expense-Emply's Bonus|" & _
            "Rent expense|Supplies expense|Travelling expense|Shipping expenses|" & _
            "Postage expenses|Repair(s) and maintenance expense|" & _
            "Advertisement|Utilities expense|Insurance expense|" & _
            "Donation expense (adjustment)|Entertainment expense|Taxes|Every depreciation|" & _
            "Various amortization|Meal expenses|Employee benefits/welfare|Overtime pay|" & _
            "Other expense|Purchase miscellaneously|Export expenses|Service charge|" & _
            "Inspection fee|Transportation expenses|Packing expenses|Professional service fees|" & _
            "Import the freight charge|Sample expense|Total operating expenses|" & _
            "Rate of total operating expenses|Income (loss) from operations|" & _
            "Rate of income (loss) from operations|Non-operating revenue|" & _
            "General dividend and dividend|Interest revenue|Commission revenue/income|" & _
            "Foreign exchange gain|Other incomes|Total non-operating revenue|" & _
            "Loss and expenses that do not open|Interest expense|Foreign exchange loss|" & _
            "Other losses|Total loss and expenses that do not open|" & _
            "Income (loss) before income tax|Rate of income (loss) before tax|" & _
            "Taxable income|Net income(loss)"
    plLabels = Split(plStr, "|")

    ' ================================================================
    ' 3. DIEN DU LIEU VAO SHEET 003-TW
    ' ================================================================
    destRow = 4   ' bat dau tu dong 4

    ' ---- 3a. Xu ly BALANCE SHEET ----
    For i = LBound(bsLabels) To UBound(bsLabels)
        lbl = bsLabels(i)
        cleanLbl = CleanTxt(lbl)
        wsNew.Cells(destRow, "A").Value = cleanLbl

        ' Gan co dinh cho 2 dong bi wrap sai vi nguon TW BS nam cheo 2 ben.
        If cleanLbl = "Minus : Accumulated" Then
            wsNew.Cells(destRow, "B").Formula = "='TW BS'!F51"
        ElseIf cleanLbl = "Accumulated" Then
            wsNew.Cells(destRow, "B").Formula = "='TW BS'!M53"
        ' ?n giá tr? ? Total Liabilities, nhưng v?n gi? d?ng đ? layout không v?
        ElseIf cleanLbl = "Total Liabilities" Then
            wsNew.Cells(destRow, "B").Value = ""
        ' Cac dong tieu de/nhom chi hien thi text, khong lookup so
        ElseIf IsNoValueHeader(cleanLbl) Then
            wsNew.Cells(destRow, "B").Value = ""
        Else
            ' Quet toan bo sheet TW BS de tim nhan, roi do sang phai de tim o so
            ' (thay vi co dinh cot F -> ho tro ca 2 khoi Tai san / No & VCSH nam canh nhau)
            valRef = FindValueRef(wb, wsBS, cleanLbl)
            If valRef <> "" Then
                If Left$(valRef, 1) = "=" Then
                    wsNew.Cells(destRow, "B").Formula = valRef
                Else
                    wsNew.Cells(destRow, "B").Formula = "=" & valRef
                End If
            Else
                wsNew.Cells(destRow, "B").Value = ""   ' khong tim thay
            End If
        End If

        If IsHeaderLabel(cleanLbl) Or Left(cleanLbl, 5) = "Total" Or cleanLbl = "Liabilities and Equity" Then
            wsNew.Cells(destRow, "A").Font.Bold = True
        End If
        If rowOf.Exists(cleanLbl) Then rowOf(cleanLbl) = destRow
        destRow = destRow + 1
    Next i

    ' ---- 3b. Them dong CHECK cho BALANCE SHEET ----
    checkRow = destRow
    If rowOf("Total Assets") > 0 And rowOf("Total Equity") > 0 Then
        valRef = FindValueRef(wb, wsBS, "Total Liabilities")
        wsNew.Cells(checkRow, "A").Value = "check"
        If valRef <> "" Then
            If Left$(valRef, 1) = "=" Then
                wsNew.Cells(checkRow, "B").Formula = valRef & "+B" & rowOf("Total Equity") & "-B" & rowOf("Total Assets")
            Else
                wsNew.Cells(checkRow, "B").Formula = "=" & valRef & "+B" & rowOf("Total Equity") & "-B" & rowOf("Total Assets")
            End If
        Else
            wsNew.Cells(checkRow, "B").Value = "check (KHONG tim duoc Total Liabilities)"
        End If
    Else
        wsNew.Cells(checkRow, "A").Value = "check (KHONG du du lieu de tinh - kiem tra lai TW BS)"
    End If
    destRow = checkRow + 2

    ' ---- 3c. Xu ly PROFIT AND LOSS ----
    For i = LBound(plLabels) To UBound(plLabels)
        lbl = plLabels(i)
        cleanLbl = CleanTxt(lbl)
        wsNew.Cells(destRow, "A").Value = cleanLbl

        If cleanLbl = "PROFIT AND LOSS" Or cleanLbl = "Account Name" Then
            If cleanLbl = "PROFIT AND LOSS" Then wsNew.Cells(destRow, "A").Font.Bold = True
        Else
            ' Cac dong tieu de/nhom chi hien thi text, khong lookup so
            If IsNoValueHeader(cleanLbl) Then
                wsNew.Cells(destRow, "B").Value = ""
            Else
                ' Quet nhieu cot: dong chi tiet nam o cot "Amount", dong subtotal nam o cot
                ' "Summation", dong tong lon nam o cot "Total" -> khong the co dinh 1 cot
                valRef = FindValueRef(wb, wsPL, cleanLbl)
                If valRef <> "" Then
                    If Left$(valRef, 1) = "=" Then
                        wsNew.Cells(destRow, "B").Formula = valRef
                    Else
                        wsNew.Cells(destRow, "B").Formula = "=" & valRef
                    End If
                Else
                    wsNew.Cells(destRow, "B").Value = ""
                End If
            End If

            If IsHeaderLabel(cleanLbl) Or Left(cleanLbl, 5) = "Total" Or Left(cleanLbl, 8) = "Income (" Or _
               cleanLbl = "Operating revenue" Or cleanLbl = "Cost of goods sold" Or cleanLbl = "Operating expenses" Then
                wsNew.Cells(destRow, "A").Font.Bold = True
            End If
            If rowOf.Exists(cleanLbl) Then rowOf(cleanLbl) = destRow
        End If
        destRow = destRow + 1
    Next i

    ' ---- 3d. Them EBITDA ----
    ebitdaRow = destRow
    If ebitdaRow > MAX_ROW_003 Then ebitdaRow = MAX_ROW_003
    If rowOf("Net income(loss)") > 0 And rowOf("Interest revenue") > 0 And _
       rowOf("Interest expense") > 0 And rowOf("Every depreciation") > 0 And _
       rowOf("Various amortization") > 0 And rowOf("Taxable income") > 0 Then
        wsNew.Cells(ebitdaRow, "A").Value = "EBITDA"
        wsNew.Cells(ebitdaRow, "A").Font.Bold = True
        wsNew.Cells(ebitdaRow, "B").Formula = _
            "=B" & rowOf("Net income(loss)") & _
            "-B" & rowOf("Interest revenue") & _
            "+B" & rowOf("Interest expense") & _
            "+SUM(B" & rowOf("Every depreciation") & ":B" & rowOf("Various amortization") & ")" & _
            "+B" & rowOf("Taxable income")
        destRow = ebitdaRow + 1
    Else
        wsNew.Cells(ebitdaRow, "A").Value = "EBITDA (thieu 1 hoac nhieu moc trong TW PL - kiem tra lai)"
        destRow = ebitdaRow + 1
    End If

    ' ---- 3e. Them cac CHECK cho PROFIT AND LOSS ----
    ' LUU Y: ban goc lay so dong tu sheet nguon (TW PL) roi ghep vao cot B cua
    ' sheet dich (003-TW) -> sai lech vi 2 sheet co so dong khac nhau.
    ' Ban duoi day dung dung so dong TREN 003-TW (rowOf), la noi cac gia tri
    ' da duoc dien vao that su.
    wsNew.Cells(destRow, "A").Value = "CHECK PL"
    wsNew.Cells(destRow, "A").Font.Bold = True
    destRow = destRow + 1

    If rowOf("Sales revenue") > 0 And rowOf("minus:Sales return") > 0 And _
       rowOf("minus:Sales discounts") > 0 And rowOf("Net sales") > 0 Then
        wsNew.Cells(destRow, "A").Value = "Check Net sales"
        wsNew.Cells(destRow, "B").Formula = "=B" & rowOf("Net sales") & " - (B" & rowOf("Sales revenue") & _
            "-B" & rowOf("minus:Sales return") & "-B" & rowOf("minus:Sales discounts") & ")"
    Else
        wsNew.Cells(destRow, "A").Value = "Check Net sales (thieu du lieu)"
    End If
    destRow = destRow + 1

    If rowOf("Net sales") > 0 And rowOf("Total cost of goods sold") > 0 And rowOf("Gross profit(loss)") > 0 Then
        wsNew.Cells(destRow, "A").Value = "Check Gross profit"
        wsNew.Cells(destRow, "B").Formula = "=B" & rowOf("Gross profit(loss)") & " - (B" & rowOf("Net sales") & _
            "-B" & rowOf("Total cost of goods sold") & ")"
    Else
        wsNew.Cells(destRow, "A").Value = "Check Gross profit (thieu du lieu)"
    End If
    destRow = destRow + 1

    If rowOf("Gross profit(loss)") > 0 And rowOf("Total operating expenses") > 0 And _
       rowOf("Income (loss) from operations") > 0 Then
        wsNew.Cells(destRow, "A").Value = "Check Income from operations"
        wsNew.Cells(destRow, "B").Formula = "=B" & rowOf("Income (loss) from operations") & " - (B" & _
            rowOf("Gross profit(loss)") & "-B" & rowOf("Total operating expenses") & ")"
    Else
        wsNew.Cells(destRow, "A").Value = "Check Income from operations (thieu du lieu)"
    End If
    destRow = destRow + 2

    With wsNew
        .Columns("A:B").AutoFit
        .Activate
    End With

    Application.ScreenUpdating = True

    Dim lastRow As Long
    lastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row
    If lastRow > MAX_ROW_003 Then
        MsgBox "Da tao sheet 003-TW cho thang " & inputMonth & "." & vbCrLf & _
               "CANH BAO: noi dung dang chiem den hang " & lastRow & _
               ", vuot qua gioi han " & MAX_ROW_003 & "." & vbCrLf & _
               "Kiem tra lai 'TW BS'/'TW PL' xem co dong moi phat sinh khong.", vbExclamation
    Else
        MsgBox "Da tao sheet 003-TW cho thang " & inputMonth & " thanh cong!" & vbCrLf & _
               "Noi dung ket thuc o hang " & lastRow & " (gioi han " & MAX_ROW_003 & ").", vbInformation
    End If
End Sub

' ================================================================
' Ham tim gia tri:
' - uu tien tim bang label "mo rong" lay tu TW_format
' - sau do quet source sheet theo nhan + cac dong wrap ben duoi
' - tra ve chuoi tham chieu dang "'TenSheet'!$X$Y", hoac "" neu khong thay.
' Tra ve chuoi tham chieu dang "'TenSheet'!$X$Y", hoac "" neu khong thay.
' ================================================================
Public Function FindValueRef(wb As Workbook, ws As Worksheet, ByVal searchLabel As String) As String
    Dim wsFmt As Worksheet
    Dim primaryLabel As String, secondaryLabel As String
    Dim bestRef As String, bestScore As Double

    If Trim$(searchLabel) = "" Then Exit Function

    Set wsFmt = Nothing
    On Error Resume Next
    Set wsFmt = wb.Sheets("TW_format")
    On Error GoTo 0

    primaryLabel = CleanTxt(searchLabel)
    secondaryLabel = ResolveLookupLabelFromFormat(wsFmt, searchLabel)

    bestRef = ""
    bestScore = 0

    If StrComp(ws.Name, "TW BS", vbTextCompare) = 0 Then
        If NormalizeKey(primaryLabel) = NormalizeKey("Minus : Accumulat") _
           Or NormalizeKey(primaryLabel) = NormalizeKey("Minus : Accumulat Other fixed") Then
            FindValueRef = "'" & ws.Name & "'!F57"
            Exit Function
        End If
        If NormalizeKey(primaryLabel) = NormalizeKey("Accumulated") Then
            FindValueRef = "'" & ws.Name & "'!M53"
            Exit Function
        End If
    End If

    If StrComp(ws.Name, "TW BS", vbTextCompare) = 0 Then
        bestRef = GetBsValueRef(primaryLabel)
        If bestRef = "" And CleanTxt(secondaryLabel) <> "" And CleanTxt(secondaryLabel) <> primaryLabel Then
            bestRef = GetBsValueRef(CleanTxt(secondaryLabel))
        End If
        If bestRef = "" Then
            ScanSheetForLabel ws, primaryLabel, bestRef, bestScore
            If CleanTxt(secondaryLabel) <> "" And CleanTxt(secondaryLabel) <> primaryLabel Then
                ScanSheetForLabel ws, secondaryLabel, bestRef, bestScore
            End If
        End If
        FindValueRef = bestRef
        Exit Function
    End If

    ScanSheetForLabel ws, primaryLabel, bestRef, bestScore
    If CleanTxt(secondaryLabel) <> "" And CleanTxt(secondaryLabel) <> primaryLabel Then
        ScanSheetForLabel ws, secondaryLabel, bestRef, bestScore
    End If

    FindValueRef = bestRef
End Function

' ================================================================
' BS co layout co dinh nen map thang theo dia chi de tranh lay nham.
' ================================================================
Private Function GetBsValueRef(ByVal searchLabel As String) As String
    Dim k As String
    k = NormalizeKey(searchLabel)

    Select Case k
        Case NormalizeKey("Cash")
            GetBsValueRef = "'TW BS'!F6"
        Case NormalizeKey("Cash in banks-Hua")
            GetBsValueRef = "='TW BS'!F7+'TW BS'!F9"
        Case NormalizeKey("Cash in banks-HSBC")
            GetBsValueRef = "='TW BS'!F12+'TW BS'!F13"
        Case NormalizeKey("Deposit account")
            GetBsValueRef = "'TW BS'!F16"
        Case NormalizeKey("Notes receivable")
            GetBsValueRef = "'TW BS'!F10"
        Case NormalizeKey("Accounts")
            GetBsValueRef = "'TW BS'!F17"
        Case NormalizeKey("Merchandise")
            GetBsValueRef = "'TW BS'!F19"
        Case NormalizeKey("Other prepayments")
            GetBsValueRef = "'TW BS'!F24"
        Case NormalizeKey("Prepaid insurance")
            GetBsValueRef = "'TW BS'!F25"
        Case NormalizeKey("Prepaid rents")
            GetBsValueRef = ""
        Case NormalizeKey("Temporary payments")
            GetBsValueRef = "'TW BS'!F26"
        Case NormalizeKey("Prepaid sales tax")
            GetBsValueRef = "'TW BS'!F27"
        Case NormalizeKey("Prepaid expenses")
            GetBsValueRef = "'TW BS'!F21"
        Case NormalizeKey("Overpaid sales tax")
            GetBsValueRef = "'TW BS'!F28"
        Case NormalizeKey("Advance payments")
            GetBsValueRef = "'TW BS'!F22"
        Case NormalizeKey("Should return the amount of tax to be paid")
            GetBsValueRef = "'TW BS'!F29"
        Case NormalizeKey("Long-term investments")
            GetBsValueRef = "'TW BS'!F37"
        Case NormalizeKey("Total Long-term investment")
            GetBsValueRef = "'TW BS'!F44"
        Case NormalizeKey("Transportation")
            GetBsValueRef = "'TW BS'!F49"
        Case NormalizeKey("Minus : Accumulated")
            GetBsValueRef = "'TW BS'!F51"
        Case NormalizeKey("Other fixed")
            GetBsValueRef = "'TW BS'!F55"
        Case NormalizeKey("Minus : Accumulat Other fixed")
            GetBsValueRef = "'TW BS'!F57"
        Case NormalizeKey("Refundable")
            GetBsValueRef = "'TW BS'!F69"
        Case NormalizeKey("Have not amortized")
            GetBsValueRef = "'TW BS'!F71"
        Case NormalizeKey("Payment on behalf")
            GetBsValueRef = "'TW BS'!F74"
        Case NormalizeKey("Total Other non-current assets")
            GetBsValueRef = "'TW BS'!F81"
        Case NormalizeKey("Total Assets")
            GetBsValueRef = "'TW BS'!F85"
        Case NormalizeKey("Accounts payable")
            GetBsValueRef = "'TW BS'!M7"
        Case NormalizeKey("Accrued expenses")
            GetBsValueRef = "'TW BS'!M8"
        Case NormalizeKey("Advances from")
            GetBsValueRef = "'TW BS'!M9"
        Case NormalizeKey("Temporary receipts")
            GetBsValueRef = "'TW BS'!M11"
        Case NormalizeKey("Current account")
            GetBsValueRef = "'TW BS'!M12"
        Case NormalizeKey("Bank loan")
            GetBsValueRef = "'TW BS'!M6"
        Case NormalizeKey("Sales tax payable")
            GetBsValueRef = "'TW BS'!M14"
        Case NormalizeKey("Receipts under")
            GetBsValueRef = "'TW BS'!M15"
        Case NormalizeKey("Total Current liabilities")
            GetBsValueRef = "'TW BS'!M17"
        Case NormalizeKey("Withhold taxes")
            GetBsValueRef = "'TW BS'!M31"
        Case NormalizeKey("Long-term debt")
            GetBsValueRef = "'TW BS'!M24"
        Case NormalizeKey("Total Long-term liabilities")
            GetBsValueRef = "'TW BS'!M26"
        Case NormalizeKey("Total Liabilities")
            GetBsValueRef = "'TW BS'!M43"
        Case NormalizeKey("Capital")
            GetBsValueRef = "'TW BS'!M46"
        Case NormalizeKey("Total Capital")
            GetBsValueRef = "'TW BS'!M47"
        Case NormalizeKey("Accumulated")
            GetBsValueRef = "'TW BS'!M53"
        Case NormalizeKey("Net income or loss")
            GetBsValueRef = "'TW BS'!M57"
        Case NormalizeKey("Total Retained earnings")
            GetBsValueRef = "'TW BS'!M61"
        Case NormalizeKey("Legal reserve")
            GetBsValueRef = "'TW BS'!M50"
        Case NormalizeKey("Total Equity")
            GetBsValueRef = "'TW BS'!M64"
        Case Else
            GetBsValueRef = ""
    End Select
End Function

' ================================================================
' Quet source sheet theo nhan, co xu ly label bi wrap nhieu dong.
' ================================================================
Private Sub ScanSheetForLabel(ws As Worksheet, ByVal searchLabel As String, ByRef bestRef As String, ByRef bestScore As Double)
    Dim searchKey As String

    searchKey = NormalizeKey(searchLabel)
    If searchKey = "" Then Exit Sub

    If StrComp(ws.Name, "TW BS", vbTextCompare) = 0 Then
        ScanBSSection ws, 1, 6, searchKey, bestRef, bestScore
        ScanBSSection ws, 8, 13, searchKey, bestRef, bestScore
    Else
        ScanPLSection ws, searchKey, bestRef, bestScore
    End If
End Sub

' ================================================================
' Duyet 1 nua cua TW BS (cot nhan + cot gia tri)
' ================================================================
Private Sub ScanBSSection(ws As Worksheet, ByVal labelCol As Long, ByVal valueCol As Long, _
                          ByVal searchKey As String, ByRef bestRef As String, ByRef bestScore As Double)
    Dim lastRow As Long, r As Long, endRow As Long
    Dim lbl As String, fullLbl As String
    Dim addr As String, candKey As String, score As Double

    lastRow = ws.Cells(ws.Rows.Count, labelCol).End(xlUp).Row
    r = 1
    Do While r <= lastRow
        lbl = CleanTxt(CStr(ws.Cells(r, labelCol).Value & ""))
        If lbl <> "" Then
            fullLbl = lbl
            endRow = r
            Do While endRow + 1 <= lastRow
                If CleanTxt(CStr(ws.Cells(endRow + 1, labelCol).Value & "")) <> "" _
                   And Not HasNumericOnRow(ws, endRow + 1, valueCol) Then
                    fullLbl = fullLbl & " " & CleanTxt(CStr(ws.Cells(endRow + 1, labelCol).Value & ""))
                    endRow = endRow + 1
                Else
                    Exit Do
                End If
            Loop

            If IsNumericValue(ws.Cells(r, valueCol).Value) Then
                addr = "'" & ws.Name & "'!" & ws.Cells(r, valueCol).Address(False, False)
                candKey = NormalizeKey(fullLbl)
                score = LabelMatchScore(searchKey, candKey)
                If score > bestScore Then
                    bestScore = score
                    bestRef = addr
                End If
            End If

            r = endRow + 1
        Else
            r = r + 1
        End If
    Loop
End Sub

' ================================================================
' Duyet TW PL: cot A + cot gia tri dau tien trong C/F/I
' ================================================================
Private Sub ScanPLSection(ws As Worksheet, ByVal searchKey As String, ByRef bestRef As String, ByRef bestScore As Double)
    Dim lastRow As Long, r As Long, endRow As Long
    Dim lbl As String, fullLbl As String
    Dim addr As String, candKey As String, score As Double

    If searchKey = NormalizeKey("Payroll expense-The company carrying and-") Then
        bestRef = "'" & ws.Name & "'!C33"
        bestScore = 10000
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    r = 1
    Do While r <= lastRow
        lbl = CleanTxt(CStr(ws.Cells(r, "A").Value & ""))
        If lbl <> "" Then
            fullLbl = lbl
            endRow = r
            Do While endRow + 1 <= lastRow
                If CleanTxt(CStr(ws.Cells(endRow + 1, "A").Value & "")) <> "" _
                   And Not HasNumericOnRow(ws, endRow + 1, 3) _
                   And Not HasNumericOnRow(ws, endRow + 1, 6) _
                   And Not HasNumericOnRow(ws, endRow + 1, 9) Then
                    fullLbl = fullLbl & " " & CleanTxt(CStr(ws.Cells(endRow + 1, "A").Value & ""))
                    endRow = endRow + 1
                Else
                    Exit Do
                End If
            Loop

            addr = FirstNumericAddrPL(ws, r)
            If addr <> "" Then
                candKey = NormalizeKey(fullLbl)
                score = LabelMatchScore(searchKey, candKey)
                If score > bestScore Then
                    bestScore = score
                    bestRef = "'" & ws.Name & "'!" & addr
                End If
            End If

            r = endRow + 1
        Else
            r = r + 1
        End If
    Loop
End Sub

' ================================================================
' Tim nhan thay the tu TW_format.
' ================================================================
Private Function ResolveLookupLabelFromFormat(wsFmt As Worksheet, ByVal searchLabel As String) As String
    Dim searchKey As String
    Dim r As Long, lastRow As Long
    Dim a As String, p As String

    If wsFmt Is Nothing Then
        ResolveLookupLabelFromFormat = searchLabel
        Exit Function
    End If

    searchKey = NormalizeKey(searchLabel)
    If searchKey = "" Then
        ResolveLookupLabelFromFormat = searchLabel
        Exit Function
    End If

    ' Mot vai alias ro rang giup BS/PL map dung hon
    Select Case searchKey
        Case "totalnonoperatingrevenue"
            ResolveLookupLabelFromFormat = "Total non-operating income"
            Exit Function
        Case "nonoperatingrevenue"
            ResolveLookupLabelFromFormat = ""
            Exit Function
    End Select

    lastRow = wsFmt.Cells(wsFmt.Rows.Count, "A").End(xlUp).Row
    For r = 1 To lastRow
        a = CleanTxt(CStr(wsFmt.Cells(r, "A").Value & ""))
        p = CleanTxt(CStr(wsFmt.Cells(r, "P").Value & ""))
        If a <> "" Then
            If LabelMatchScore(searchKey, NormalizeKey(a)) > 0 Then
                If p <> "" And Not IsNoValueHeader(a) Then
                    ResolveLookupLabelFromFormat = p
                Else
                    ResolveLookupLabelFromFormat = a
                End If
                Exit Function
            End If
        End If
        If p <> "" Then
            If LabelMatchScore(searchKey, NormalizeKey(p)) > 0 Then
                ResolveLookupLabelFromFormat = p
                Exit Function
            End If
        End If
    Next r

    ResolveLookupLabelFromFormat = searchLabel
End Function

' ================================================================
' Lay dia chi o so dau tien tren dong PL (cot C/F/I).
' ================================================================
Private Function FirstNumericAddrPL(ws As Worksheet, ByVal r As Long) As String
    Dim c As Long
    For c = 3 To 9 Step 3
        If IsNumericValue(ws.Cells(r, c).Value) Then
            FirstNumericAddrPL = ws.Cells(r, c).Address(False, False)
            Exit Function
        End If
    Next c
    FirstNumericAddrPL = ""
End Function

' ================================================================
' Kiem tra 1 dong co gia tri so o cot chi dinh hay khong.
' ================================================================
Private Function HasNumericOnRow(ws As Worksheet, ByVal r As Long, ByVal c As Long) As Boolean
    HasNumericOnRow = IsNumericValue(ws.Cells(r, c).Value)
End Function

' ================================================================
' Diem khop label:
' - exact > contains > contained-by
' ================================================================
Private Function LabelMatchScore(ByVal searchKey As String, ByVal candKey As String) As Double
    If searchKey = "" Or candKey = "" Then Exit Function

    If StrComp(searchKey, candKey, vbTextCompare) = 0 Then
        LabelMatchScore = 1000
    ElseIf InStr(1, candKey, searchKey, vbTextCompare) > 0 Then
        LabelMatchScore = 800 - (Abs(Len(candKey) - Len(searchKey)) / 10#)
    ElseIf InStr(1, searchKey, candKey, vbTextCompare) > 0 Then
        LabelMatchScore = 700 - (Abs(Len(candKey) - Len(searchKey)) / 10#)
    Else
        LabelMatchScore = 0
    End If
End Function

' ================================================================
' Tra ve True neu gia tri co the xem la so.
' ================================================================
Private Function IsNumericValue(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v & "")) = "" Then Exit Function
    IsNumericValue = IsNumeric(v)
End Function

' ================================================================
' Ham chuan hoa chuoi (trim, bo khoang trang thua, chuyen ky tu full-width)
' ================================================================
Public Function CleanTxt(ByVal s As String) As String
    Dim i As Long, code As Long, ch As String, outp As String
    outp = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 9, 10, 13, 160
                outp = outp & " "
            Case 65306
                outp = outp & ":"
            Case 65307
                outp = outp & ";"
            Case 65292
                outp = outp & ","
            Case 65281
                outp = outp & "!"
            Case 65288
                outp = outp & "("
            Case 65289
                outp = outp & ")"
            Case Else
                outp = outp & ch
        End Select
    Next i
    Do While InStr(outp, "  ") > 0
        outp = Replace(outp, "  ", " ")
    Loop
    CleanTxt = Trim$(outp)
End Function

' ================================================================
' Chuan hoa de so khop label (bo dau cach / dau cau / full-width).
' ================================================================
Private Function NormalizeKey(ByVal s As String) As String
    Dim t As String

    t = LCase$(CleanTxt(s))
    t = Replace(t, " ", "")
    t = Replace(t, "-", "")
    t = Replace(t, "_", "")
    t = Replace(t, "/", "")
    t = Replace(t, "'", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    t = Replace(t, ".", "")
    t = Replace(t, ",", "")
    t = Replace(t, ":", "")
    t = Replace(t, ";", "")
    t = Replace(t, "?", "")
    t = Replace(t, "?", "")
    NormalizeKey = t
End Function

' ================================================================
' Cac label chi la header / tieu de, khong can lookup so.
' ================================================================
Private Function IsNoValueHeader(ByVal cleanLbl As String) As Boolean
    Select Case cleanLbl
        Case "BALANCE SHEET", "Assets Account", "Liabilities and Equity", _
             "PROFIT AND LOSS", "Account Name", _
             "Current assets", "Long-term", "Long-term investment", "Property, plant and equipment", _
             "Other non-current assets", "Current liabilities", _
             "Other non-current liabilities", "Long-term liabilities", _
             "Capital", "Retained earnings", _
             "Operating revenue", "Cost of goods sold", "Operating expenses", _
             "Non-operating revenue", "Loss and expenses that do not open", _
             "Non-operating income", "Non-operating expenses", "check"
            IsNoValueHeader = True
        Case Else
            IsNoValueHeader = False
    End Select
End Function

' ================================================================
' Ham xac dinh tieu de nhom (de in dam)
' ================================================================
Private Function IsHeaderLabel(ByVal cleanLbl As String) As Boolean
    Select Case cleanLbl
        Case "Assets Account", "Current assets", "Long-term", "Non-current assets", _
             "Liabilities and Equity", "Current liabilities", "Non-current liabilities", _
             "Capital", "Retained earnings", _
             "Operating revenue", "Cost of goods sold", "Operating expenses", _
             "Non-operating income", "Non-operating revenue", "Non-operating expenses", _
             "PROFIT AND LOSS"
            IsHeaderLabel = True
        Case Else
            IsHeaderLabel = False
    End Select
End Function


