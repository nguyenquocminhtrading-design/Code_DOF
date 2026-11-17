Attribute VB_Name = "Module_IC_Elimination"
Option Explicit

' ======================================================================
' INTERCOMPANY ELIMINATION MACRO - DDF GROUP
' ----------------------------------------------------------------------
' Chuc nang:
'   1. Tu dong quet toan bo cac sheet trong workbook, nhan dien sheet
'      nao la sheet du lieu ban hang cua tung phap nhan (DOF, ADF,
'      DOAU, CAN...) dua tren TEN SHEET.
'   2. Tren moi sheet, tu dong tim cot "Name"/"Customer Name" (khach
'      hang), cot Doanh thu (Net Sales/Amount/Sales) va cot Gia von
'      (Total Cost/COGS), hoac tinh Gia von = Doanh thu - GP$ neu
'      sheet chi co cot loi nhuan gop.
'   3. Doi chieu ten khach hang voi danh sach "ben lien quan" (cac
'      phap nhan/chi nhanh trong Group) khai bao trong RELATED_PARTIES.
'   4. Cong don Doanh thu & Gia von cua cac giao dich noi bo, xuat ra
'      1 sheet ket qua "IC_Elimination" gom:
'         - Bang tong hop theo tung sheet (phap nhan ban): Tong Doanh
'           thu/Gia von BAN DAU (ca sheet), phan BI LOAI TRU (noi bo),
'           va phan SAU LOAI TRU (= ban dau - bi loai tru)
'         - Bang chi tiet theo tung doi tac noi bo trong moi sheet
'
' Cach dung:
'   - Mo workbook du lieu (vi du DDF_Group_-_sale_report_FY26.xlsx)
'   - Alt+F11 -> Insert Module -> dan toan bo code nay vao (hoac
'     File > Import File... chon file .bas nay)
'   - Chinh sua mang RELATED_PARTIES ben duoi cho dung voi cong ty cua
'     ban (them/bot ten cho khop 100% voi cot Name trong du lieu -
'     macro da tu dong TRIM khoang trang va bo qua chu HOA/thuong nen
'     khong can lo bien the vi du "DAN-D FOODS LTD." vs "DAN-D FOODS LTD")
'   - Bam F5 hoac Alt+F8 -> chon EliminateIntercompany -> Run
' ======================================================================

Const OUTPUT_SHEET_NAME As String = "IC_Elimination"

' Tu khoa de nhan dien sheet du lieu ban hang theo TEN SHEET (khong phan
' biet hoa/thuong). Them ten sheet moi vao day neu Group co them cong ty.
Private Function GetSheetKeywords() As Variant
    GetSheetKeywords = Array("DOF", "ADF", "DOAU", "CAN")
End Function

' Danh sach ten "ben lien quan" (cong ty me / chi nhanh / cong ty con
' trong Group DDF) can loai tru khi xuat hien o cot khach hang.
' -> SUA/THEM/XOA dong trong mang nay khi co thay doi co cau Group.
Private Function GetRelatedParties() As Variant
    GetRelatedParties = Array( _
        "CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN", _
        "CÔNG TY CP THỰC PHẨM DÂN ÔN - CHI NHÁNH MỸ PHƯỚC 3", _
        "CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN - CHI NHÁNH MỸ PHƯỚC 3", _
        "CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN-CN MỸ PHƯỚC 3", _
        "CÔNG TY CỔ PHẦN THỰC PHẨM AN ĐIỀN", _
        "QUỸ XÃ HỘI - TỪ THIỆN DÂN ÔN", _
        "DAN D PAK CO., LTD", _
        "DAN-D FOODS HONG KONG LIMITED", _
        "DAN-D FOODS HONGKONG LTD.", _
        "DAN-D FOODS LTD", _
        "DAN-D FOODS LTD.", _
        "DAN-D-PAK USA", _
        "DANWAY HOLDINGS LIMITED", _
        "DOF - VIETNAM", _
        "DOF MY PHUOC 3", _
        "DAN-D-FOODS, LTD-CANADA", _
        "DAN ON FOODS CORPORATION-CAD", _
        "DAN-D FOODS( HK LTD)-US$" _
    )
End Function

' Cac ten cot co the dung cho cot Ten khach hang / Doanh thu / Gia von,
' macro se thu lan luot theo thu tu uu tien cho toi khi tim thay.
Private Function GetNameHeaderCandidates() As Variant
    GetNameHeaderCandidates = Array("Name", "Customer Name")
End Function
Private Function GetRevenueHeaderCandidates() As Variant
    GetRevenueHeaderCandidates = Array("Net Sales", "Amount", "Sales")
End Function
Private Function GetCogsHeaderCandidates() As Variant
    GetCogsHeaderCandidates = Array("Total Cost", "COGS")
End Function
Private Function GetGPHeaderCandidates() As Variant
    GetGPHeaderCandidates = Array("GP$", "GM")
End Function

' ======================================================================
' MAIN
' ======================================================================
Sub EliminateIntercompany()

    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim relatedParties As Variant
    Dim sheetKeywords As Variant
    Dim outRow As Long
    Dim sheetSummary As Object   ' Scripting.Dictionary: sheetName -> Array(revTotal, cogsTotal)
    Dim k As Variant

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Dang chay Intercompany Elimination..."

    relatedParties = GetRelatedParties()
    sheetKeywords = GetSheetKeywords()

    ' --- Xoa sheet ket qua cu (neu co) va tao moi ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    wsOut.Name = OUTPUT_SHEET_NAME

    Set sheetSummary = CreateObject("Scripting.Dictionary")

    ' --- Tieu de ---
    With wsOut
        .Range("A1").Value = "BANG LOAI TRU GIAO DICH NOI BO (INTERCOMPANY ELIMINATION)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Tu dong tao boi macro EliminateIntercompany - chay luc: " & Format(Now, "dd/mm/yyyy hh:nn")
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Size = 9
    End With

    outRow = 4
    Dim summaryHeaderRow As Long
    summaryHeaderRow = outRow
    wsOut.Cells(outRow, 1).Value = "1. TONG HOP THEO SHEET (PHAP NHAN BAN) - TRUOC / LOAI TRU / SAU LOAI TRU"
    wsOut.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = "Sheet"
    wsOut.Cells(outRow, 2).Value = "So dong noi bo"
    wsOut.Cells(outRow, 3).Value = "Tong Doanh thu ban dau"
    wsOut.Cells(outRow, 4).Value = "Tong Gia von ban dau"
    wsOut.Cells(outRow, 5).Value = "Doanh thu bi loai tru"
    wsOut.Cells(outRow, 6).Value = "Gia von bi loai tru"
    wsOut.Cells(outRow, 7).Value = "Doanh thu sau loai tru"
    wsOut.Cells(outRow, 8).Value = "Gia von sau loai tru"
    wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 8)).Font.Bold = True
    outRow = outRow + 1
    Dim summaryFirstRow As Long
    summaryFirstRow = outRow

    ' Cho tru lai o phia duoi (chi tiet), can biet vi tri roi quay lai dien
    Dim summaryRows As Object
    Set summaryRows = CreateObject("Scripting.Dictionary")

    ' --- Duyet toan bo sheet, xu ly cac sheet khop tu khoa ---
    Dim detailStartRow As Long
    ' Tinh truoc so sheet du lieu de biet cho chua bang tong hop (so dong = so sheet khop)
    Dim matchCount As Long
    matchCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> OUTPUT_SHEET_NAME Then
            If SheetMatchesKeyword(ws.Name, sheetKeywords) Then matchCount = matchCount + 1
        End If
    Next ws

    detailStartRow = summaryFirstRow + matchCount + 4   ' + dong Tong cong + dong ghi chu + 1 dong trong

    Dim curDetailRow As Long
    curDetailRow = detailStartRow
    Dim curSummaryRow As Long
    curSummaryRow = summaryFirstRow

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> OUTPUT_SHEET_NAME Then
            If SheetMatchesKeyword(ws.Name, sheetKeywords) Then
                Dim revTotal As Double, cogsTotal As Double, rowCount As Long
                Dim grandRevTotal As Double, grandCogsTotal As Double
                ProcessSheet ws, relatedParties, wsOut, curDetailRow, revTotal, cogsTotal, rowCount, grandRevTotal, grandCogsTotal

                wsOut.Cells(curSummaryRow, 1).Value = ws.Name
                wsOut.Cells(curSummaryRow, 2).Value = rowCount
                wsOut.Cells(curSummaryRow, 3).Value = grandRevTotal                  ' Tong Doanh thu ban dau
                wsOut.Cells(curSummaryRow, 4).Value = grandCogsTotal                 ' Tong Gia von ban dau
                wsOut.Cells(curSummaryRow, 5).Value = revTotal                       ' Doanh thu bi loai tru
                wsOut.Cells(curSummaryRow, 6).Value = cogsTotal                      ' Gia von bi loai tru
                wsOut.Cells(curSummaryRow, 7).Value = grandRevTotal - revTotal       ' Doanh thu sau loai tru
                wsOut.Cells(curSummaryRow, 8).Value = grandCogsTotal - cogsTotal     ' Gia von sau loai tru
                wsOut.Range(wsOut.Cells(curSummaryRow, 3), wsOut.Cells(curSummaryRow, 8)).NumberFormat = "#,##0;(#,##0);""-"""
                curSummaryRow = curSummaryRow + 1
            End If
        End If
    Next ws

    ' --- Dong TONG CONG (luu y: chi cong don duoc neu cac sheet cung don vi tien te) ---
    wsOut.Cells(curSummaryRow, 1).Value = "TONG CONG (*)"
    wsOut.Cells(curSummaryRow, 1).Font.Bold = True
    Dim gTotRow As Long
    gTotRow = curSummaryRow
    Dim col As Long
    For col = 3 To 8
        wsOut.Cells(gTotRow, col).Formula = "=SUM(" & wsOut.Cells(summaryFirstRow, col).Address(False, False) & ":" & wsOut.Cells(gTotRow - 1, col).Address(False, False) & ")"
        wsOut.Cells(gTotRow, col).NumberFormat = "#,##0;(#,##0);""-"""
        wsOut.Cells(gTotRow, col).Font.Bold = True
    Next col
    wsOut.Range(wsOut.Cells(gTotRow, 1), wsOut.Cells(gTotRow, 8)).Interior.Color = RGB(217, 225, 242)
    curSummaryRow = curSummaryRow + 1
    wsOut.Cells(curSummaryRow, 1).Value = "(*) Chi cong hop le neu cac sheet dang cung 1 don vi tien te. Neu co sheet o ngoai te (vd DOAU, CAN), quy doi truoc khi cong tong."
    wsOut.Cells(curSummaryRow, 1).Font.Italic = True
    wsOut.Cells(curSummaryRow, 1).Font.Size = 9

    ' --- Dinh dang chung ---
    With wsOut
        .Range(.Cells(summaryHeaderRow, 1), .Cells(summaryHeaderRow, 8)).Interior.Color = RGB(252, 228, 214)
        .Range(.Cells(summaryFirstRow - 1, 1), .Cells(summaryFirstRow - 1, 8)).Interior.Color = RGB(31, 78, 120)
        .Range(.Cells(summaryFirstRow - 1, 1), .Cells(summaryFirstRow - 1, 8)).Font.Color = RGB(255, 255, 255)
        .Columns("A:H").AutoFit
        .Activate
        .Range("A1").Select
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Da chay xong Intercompany Elimination." & vbCrLf & _
           "Xem ket qua trong sheet: " & OUTPUT_SHEET_NAME, vbInformation

End Sub

' ======================================================================
' Xu ly 1 sheet: tim cot, loc giao dich noi bo, cong don theo doi tac,
' ghi bang chi tiet ra wsOut bat dau tu curRow (se tu tang curRow len).
' Tra ve (qua tham so ByRef):
'   - revTotal / cogsTotal / rowCount : Doanh thu, Gia von, so dong CUA
'     RIENG cac giao dich noi bo (bi loai tru)
'   - grandRevTotal / grandCogsTotal  : Tong Doanh thu, Gia von CUA
'     TOAN BO sheet (ca noi bo lan ben ngoai) - dung de tinh "truoc/
'     sau loai tru"
' ======================================================================
Private Sub ProcessSheet(ws As Worksheet, relatedParties As Variant, _
        wsOut As Worksheet, ByRef curRow As Long, _
        ByRef revTotal As Double, ByRef cogsTotal As Double, ByRef rowCount As Long, _
        ByRef grandRevTotal As Double, ByRef grandCogsTotal As Double)

    revTotal = 0: cogsTotal = 0: rowCount = 0
    grandRevTotal = 0: grandCogsTotal = 0

    Dim headerCell As Range
    Set headerCell = FindHeaderCell(ws, GetNameHeaderCandidates())
    If headerCell Is Nothing Then
        wsOut.Cells(curRow, 1).Value = "[" & ws.Name & "] Khong tim thay cot Name/Customer Name - bo qua sheet nay"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 2
        Exit Sub
    End If

    Dim headerRow As Long, nameCol As Long
    headerRow = headerCell.Row
    nameCol = headerCell.Column

    Dim revCol As Long, cogsCol As Long, gpCol As Long
    revCol = FindColInRow(ws, headerRow, GetRevenueHeaderCandidates())
    cogsCol = FindColInRow(ws, headerRow, GetCogsHeaderCandidates())
    gpCol = 0
    If cogsCol = 0 Then gpCol = FindColInRow(ws, headerRow, GetGPHeaderCandidates())

    If revCol = 0 Then
        wsOut.Cells(curRow, 1).Value = "[" & ws.Name & "] Khong tim thay cot Doanh thu - bo qua sheet nay"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 2
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    ' Doc du lieu vao mang (nhanh hon doc tung o) - tu cot 1 den cot lon nhat can dung
    Dim maxCol As Long
    maxCol = nameCol
    If revCol > maxCol Then maxCol = revCol
    If cogsCol > maxCol Then maxCol = cogsCol
    If gpCol > maxCol Then maxCol = gpCol

    Dim dataArr As Variant
    dataArr = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, maxCol)).Value2

    ' Dictionary: ten doi tac -> Array(revenue, cogs, count)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, nm As String
    Dim rev As Double, cogs As Double
    For i = 1 To UBound(dataArr, 1)
        nm = Trim(CStr(dataArr(i, nameCol)))

        ' --- Tinh rev/cogs cua DONG NAY (du la noi bo hay ben ngoai) ---
        rev = 0: cogs = 0
        If IsNumeric(dataArr(i, revCol)) Then rev = CDbl(dataArr(i, revCol))
        If cogsCol > 0 Then
            If IsNumeric(dataArr(i, cogsCol)) Then cogs = CDbl(dataArr(i, cogsCol))
        ElseIf gpCol > 0 Then
            Dim gp As Double: gp = 0
            If IsNumeric(dataArr(i, gpCol)) Then gp = CDbl(dataArr(i, gpCol))
            cogs = rev - gp   ' COGS = Doanh thu - Loi nhuan gop
        End If

        ' --- Cong vao tong ban dau cua CA SHEET (moi dong deu tinh) ---
        grandRevTotal = grandRevTotal + rev
        grandCogsTotal = grandCogsTotal + cogs

        ' --- Neu la giao dich noi bo thi cong them vao phan bi loai tru ---
        If Len(nm) > 0 Then
            If IsRelatedParty(nm, relatedParties) Then
                revTotal = revTotal + rev
                cogsTotal = cogsTotal + cogs
                rowCount = rowCount + 1

                If dict.Exists(nm) Then
                    Dim arr As Variant
                    arr = dict(nm)
                    arr(0) = arr(0) + rev
                    arr(1) = arr(1) + cogs
                    arr(2) = arr(2) + 1
                    dict(nm) = arr
                Else
                    dict.Add nm, Array(rev, cogs, 1)
                End If
            End If
        End If
    Next i

    ' --- Ghi bang chi tiet theo doi tac cho sheet nay ---
    wsOut.Cells(curRow, 1).Value = "Chi tiet - sheet """ & ws.Name & """"
    wsOut.Cells(curRow, 1).Font.Bold = True
    wsOut.Cells(curRow, 1).Interior.Color = RGB(252, 228, 214)
    wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 4)).Interior.Color = RGB(252, 228, 214)
    curRow = curRow + 1

    wsOut.Cells(curRow, 1).Value = "Doi tac noi bo"
    wsOut.Cells(curRow, 2).Value = "Doanh thu bi loai tru"
    wsOut.Cells(curRow, 3).Value = "Gia von bi loai tru"
    wsOut.Cells(curRow, 4).Value = "So dong"
    wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 4)).Font.Bold = True
    wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 4)).Interior.Color = RGB(217, 225, 242)
    curRow = curRow + 1

    Dim key As Variant
    For Each key In dict.Keys
        Dim v As Variant
        v = dict(key)
        wsOut.Cells(curRow, 1).Value = key
        wsOut.Cells(curRow, 2).Value = v(0)
        wsOut.Cells(curRow, 3).Value = v(1)
        wsOut.Cells(curRow, 4).Value = v(2)
        wsOut.Range(wsOut.Cells(curRow, 2), wsOut.Cells(curRow, 3)).NumberFormat = "#,##0;(#,##0);""-"""
        curRow = curRow + 1
    Next key

    If dict.Count = 0 Then
        wsOut.Cells(curRow, 1).Value = "(Khong co giao dich noi bo nao khop trong sheet nay)"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 1
    End If

    curRow = curRow + 1   ' dong trong ngan cach

End Sub

' ======================================================================
' Ham tien ich
' ======================================================================

' Sheet co ten khop 1 trong cac tu khoa khong (khong phan biet hoa/thuong)
Private Function SheetMatchesKeyword(sheetName As String, keywords As Variant) As Boolean
    Dim k As Variant
    Dim upperName As String
    upperName = UCase(Trim(sheetName))
    For Each k In keywords
        If upperName = UCase(k) Then
            SheetMatchesKeyword = True
            Exit Function
        End If
    Next k
    SheetMatchesKeyword = False
End Function

' Ten khach hang co nam trong danh sach ben lien quan khong (Trim + UCase)
Private Function IsRelatedParty(nm As String, relatedParties As Variant) As Boolean
    Dim p As Variant
    Dim target As String
    target = UCase(Trim(nm))
    For Each p In relatedParties
        If target = UCase(Trim(CStr(p))) Then
            IsRelatedParty = True
            Exit Function
        End If
    Next p
    IsRelatedParty = False
End Function

' Tim o trong 10 dong dau / 60 cot dau co gia tri (Trim, khong phan biet
' hoa/thuong) khop voi 1 trong cac ten trong candidates. Tra ve o dau
' tien tim thay (dung lam header cell), hoac Nothing neu khong thay.
Private Function FindHeaderCell(ws As Worksheet, candidates As Variant) As Range
    Dim r As Long, c As Long
    Dim cell As Range
    Dim cand As Variant
    For r = 1 To 10
        For c = 1 To 60
            Set cell = ws.Cells(r, c)
            If Len(Trim(CStr(cell.Value))) > 0 Then
                For Each cand In candidates
                    If UCase(Trim(CStr(cell.Value))) = UCase(cand) Then
                        Set FindHeaderCell = cell
                        Exit Function
                    End If
                Next cand
            End If
        Next c
    Next r
    Set FindHeaderCell = Nothing
End Function

' Tim cot trong dong headerRow co ten khop 1 trong candidates (thu lan
' luot theo thu tu uu tien). Tra ve so cot, hoac 0 neu khong thay.
Private Function FindColInRow(ws As Worksheet, headerRow As Long, candidates As Variant) As Long
    Dim cand As Variant
    Dim c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 60 Then lastCol = 60
    For Each cand In candidates
        For c = 1 To lastCol
            If UCase(Trim(CStr(ws.Cells(headerRow, c).Value))) = UCase(cand) Then
                FindColInRow = c
                Exit Function
            End If
        Next c
    Next cand
    FindColInRow = 0
End Function
