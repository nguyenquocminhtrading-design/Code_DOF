Option Explicit

' ================================================================
' PHIEN BAN MOI: COT A HARD-CODE CO DINH (row 2 -> row 196),
' KHONG CON DUNG TW_format, KHONG CON VBA scan/GetBsValueRef/GetPlValueRef.
' Gia tri duoc dien bang cong thuc XLOOKUP ngay trong o (giong Excel thuong):
'   - Ben ASSET  (BS, cot A):        =XLOOKUP(A{r},'TW BS'!$A:$A,'TW BS'!F:F,0)
'   - Ben LIAB&EQUITY (BS, cot H):   =XLOOKUP(A{r},'TW BS'!$H:$H,'TW BS'!$M:$M,0)
'   - PROFIT AND LOSS (PL, cot A):   =XLOOKUP(A{r},'TW PL'!$A:$A,'TW PL'!$C:$C,0)
' EBITDA (row 196) dung cong thuc rieng, cong/tru cac dong da co san.
'
' CANH BAO: cong thuc XLOOKUP se tra ve 0 (khong bao loi) neu chu trong
' cot A cua 003-TW KHONG khop chinh xac tung ky tu voi chu trong TW BS /
' TW PL thang do (ke ca cach xuong dong do wrap chu). Neu thang sau
' TW BS/TW PL wrap chu khac di, dong tuong ung se tra ve 0 ma KHONG
' co canh bao. Nen doi chieu lai moi thang truoc khi tin so lieu.
'
' GHI CHU KY THUAT: danh sach nhan/section duoc dung bang cach noi
' chuoi (Dim s As String: s = s & "..."), roi Split("|") - CACH NAY
' TRANH duoc loi "too many line continuations" cua VBA (moi cau lenh
' phai <= 24 dau noi dong " _", trong khi mang 195 phan tu can rat
' nhieu dong neu viet Array(...) truc tiep).
' ================================================================

Sub Create003TW_XLOOKUP()
    Dim wb As Workbook
    Dim wsNew As Worksheet, wsBS As Worksheet, wsPL As Worksheet
    Dim inputMonth As String
    Dim labels As Variant, sections As Variant
    Dim i As Long, r As Long
    Dim n As Long

    Set wb = ThisWorkbook

    ' ---- Nhap thang ----
    inputMonth = InputBox("Enter the month (vd 2026.07):", "Month Selection", "2026.07")
    If inputMonth = "" Then Exit Sub

    ' ---- Kiem tra sheet nguon (khong con can TW_format) ----
    On Error Resume Next
    Set wsBS = wb.Sheets("TW BS")
    Set wsPL = wb.Sheets("TW PL")
    On Error GoTo 0
    If wsBS Is Nothing Or wsPL Is Nothing Then
        MsgBox "Thieu sheet 'TW BS' hoac 'TW PL'.", vbCritical
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

    ' ---- Ghi thang o hang 1 ----
    wsNew.Range("A1").Value = "Month:"
    wsNew.Range("B1").Value = inputMonth

    ' ================================================================
    ' DANH SACH NHAN COT A (195 phan tu, dien tu row 2 -> row 196)
    ' Xay dung bang noi chuoi + Split de tranh loi "too many line
    ' continuations".
    ' ================================================================
    Dim sLabels As String
    sLabels = sLabels & "BALANCE SHEET|Assets Account|Current assets|  Cash|  Cash in banks-Hua |  Nan Bank|  Cash in banks-HSBC|  Deposit account|  Notes receivable|  Accounts |"
    sLabels = sLabels & "  receivable|  Merchandise |  inventory|  Other prepayments|  Prepaid insurance|  Prepaid rents|  Temporary payments|  Prepaid sales tax|  Prepaid expenses|  Overpaid sales tax|"
    sLabels = sLabels & "  Advance payments |  Should return the |Total Current assets||Total Long-term investment|    Long-term |  investments|Total Long-term |investment||"
    sLabels = sLabels & "Property, plant and |equipment|  Land|  Buildings|  Minus:Accumulated|         depreciatio|         n-buildings|    Transportation |    facilities|     Minus:Accumulated_|"
    sLabels = sLabels & "         depreciatio|    Other fixed |    Minus:Accumulat|Other non-current |assets|    Refundable |  deposits|    Have not |  the expenses|    Payment on |"
    sLabels = sLabels & "  of others|Total Other |non-current assets|Total Assets|Liabilities and Equity|Current liabilities|  Accounts payable|  Accrued expenses|  Taxes payable|  Advances from |"
    sLabels = sLabels & "  Temporary receipts|  Other unearned revenue|  Stockholders' |  current account|  Current account |  with others|  Bank loan|  Sales tax payable|  Receipts under |  custody|"
    sLabels = sLabels & "Total Current |liabilities||  Other non-current |  liabilities|    Withhold taxes||Long-term -|liabilities|    Long-term debt |"
    sLabels = sLabels & "  payable|Total Long-term |liabilities|||Total Liabilities||Capital|  Capital|Total Capital_|"
    sLabels = sLabels & "|Retained earnings|  Accumulated |  deficit(the |  remaining sum |  after 99)|  Net income or loss|Total Retained |  Legal reserve(the |Total Equity|"
    sLabels = sLabels & "|||PROFIT AND LOSS|Account Name|Operating revenue|  Sales revenue|  minus:Sales return|  minus:Sales discounts |        and allowances|"
    sLabels = sLabels & "Net sales||Cost of goods sold|  Merchandise inventory|  (beginning inventory)|  Purchases|  minus:Purchase discounts|  minus:Purchase return|  The goods are changed over to|  minus:The goods producing|"
    sLabels = sLabels & "  Adjustment debit of the |  minus:Merchandise |        inventory(final |        inventory)|Total cost of good sold|Total operating cost|Rate of operating costs|Gross profit(loss)|Rate of gross profit|Operating expenses|"
    sLabels = sLabels & "  Payroll expense|  Payroll expense-The |  Payroll expense-Emply's Bonus|  Rent expense|  Supplies expense|  Travelling expense|  Shipping expenses|  Postage expenses|  Repair(s) and maintenance expense|  Advertisement |"
    sLabels = sLabels & "  expense,advertisement|  Utilities expense|  Insurance expense|  Donation expense (adjustment)|  Entertainment expense|  Taxes|  Every depreciation|  Various amortization|  Meal expenses|  Employee |"
    sLabels = sLabels & "  Employee benefits/welfare|  Overtime pay|  Other expense|  Purchase miscellaneously|  Export expenses|  Service charge|  Inspection fee|  Transportation expenses|  Packing expenses|  Professional service |"
    sLabels = sLabels & "Import the freight charge|  Sample expense|Total operating expenses|Rate of total operating |expenses|Income (loss) from |operations|Rate of income (loss) |from operations|Non-operating revenue|"
    sLabels = sLabels & "  General dividend and |  dividend|  Interest revenue|  Commission |  revenue/income|  Foreign exchange gain|  Other incomes|Total non-operating |revenue|Loss and expenses that |"
    sLabels = sLabels & "do not open|  Interest expense|  Foreign exchange loss|  Other losses|Total loss and expenses |that do not open|Income (loss) before |income tax|Rate of income (loss) |before tax|"
    sLabels = sLabels & "Taxable income|Net income(loss)|||EBITDA|"
    labels = Split(Left$(sLabels, Len(sLabels) - 1), "|")

    ' Nhan section tuong ung cho tung dong (ASSET / LIAB / PL / EBITDA / NONE)
    ' NONE = dong trong, khong dien cong thuc.
    Dim sSections As String
    sSections = sSections & "ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|"
    sSections = sSections & "ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|NONE|ASSET|ASSET|ASSET|ASSET|ASSET|NONE|"
    sSections = sSections & "ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|"
    sSections = sSections & "ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|ASSET|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|"
    sSections = sSections & "LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|NONE|LIAB|LIAB|"
    sSections = sSections & "LIAB|NONE|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|NONE|NONE|LIAB|NONE|LIAB|LIAB|LIAB|"
    sSections = sSections & "NONE|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|LIAB|NONE|NONE|NONE|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|NONE|PL|PL|PL|PL|PL|PL|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|"
    sSections = sSections & "PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|PL|NONE|NONE|EBITDA|"
    sections = Split(Left$(sSections, Len(sSections) - 1), "|")

    n = UBound(labels) - LBound(labels)   ' so dong (0-based length - 1)

    For i = 0 To n
        r = i + 2   ' row 2 la dong dau tien
        wsNew.Cells(r, "A").Value = labels(i)

        Select Case sections(i)
            Case "ASSET"
                wsNew.Cells(r, "B").Formula = _
                    "=SUMIFS('TW BS'!F:F,'TW BS'!$A:$A,A" & r & ")"
            Case "LIAB"
                wsNew.Cells(r, "B").Formula = _
                    "=XLOOKUP(A" & r & ",'TW BS'!$H:$H,'TW BS'!$M:$M,0)"
            Case "PL"
                wsNew.Cells(r, "B").Formula = _
                    "=XLOOKUP(A" & r & ",'TW PL'!$A:$A,'TW PL'!$C:$C,0)"
            Case "EBITDA"
                ' EBITDA = Net income(loss) - Interest revenue + Interest expense
                '        + SUM(Every depreciation : Various amortization) + Taxable income
                ' Cac dong duoi day co dinh theo dung thu tu danh sach nhan o tren:
                '   Net income(loss)     -> row 193
                '   Interest revenue     -> row 174
                '   Interest expense     -> row 183
                '   Every depreciation   -> row 148
                '   Various amortization -> row 149
                '   Taxable income       -> row 192
                wsNew.Cells(r, "B").Formula = _
                    "=B193-B174+B183+SUM(B148:B149)+B192"
            Case Else
                ' NONE: dong trong hoac tieu de khong can gia tri
        End Select

        ' In dam cac dong header / Total
        If IsBoldLabel(CStr(labels(i))) Then
            wsNew.Cells(r, "A").Font.Bold = True
        End If
    Next i

    With wsNew
        .Columns("A:B").AutoFit
        .Activate
    End With

    Application.ScreenUpdating = True

    MsgBox "Da tao sheet 003-TW cho thang " & inputMonth & " (dung XLOOKUP)." & vbCrLf & _
           "Noi dung tu hang 2 den hang " & (n + 2) & "." & vbCrLf & _
           "NHO doi chieu lai chu trong cot A voi TW BS/TW PL thang nay truoc khi tin so lieu.", vbInformation
End Sub

' ================================================================
' In dam cac dong tieu de / Total (chi de hien thi dep, khong anh
' huong toi cong thuc).
' ================================================================
Private Function IsBoldLabel(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        IsBoldLabel = False
        Exit Function
    End If
    If t = "BALANCE SHEET" Or t = "PROFIT AND LOSS" Or t = "Liabilities and Equity" _
       Or t = "EBITDA" Or Left$(t, 5) = "Total" Then
        IsBoldLabel = True
    Else
        IsBoldLabel = False
    End If
End Function

