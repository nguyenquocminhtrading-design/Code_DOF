Option Explicit

Sub AutoSumBCTC()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ----- Lay pham vi dong dang duoc BOI DEN (Selection) ----------
    If TypeName(Selection) <> "Range" Then
        MsgBox "Ban hay boi den (chon) cac dong can tinh SUM truoc khi chay macro.", vbExclamation
        Exit Sub
    End If

    Dim firstSelRow As Long, lastSelRow As Long
    firstSelRow = Selection.Rows(1).Row
    lastSelRow = Selection.Rows(Selection.Rows.Count).Row

    If lastSelRow <= firstSelRow Then
        MsgBox "Ban can boi den it nhat 2 dong (dong cha va dong con) roi moi chay macro.", vbExclamation
        Exit Sub
    End If
    '----------------------------------------------------------------

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Const LABEL_COL As Long = 1        ' Cot ten chi tieu - luon la cot A trong file nay

    ' [FIX 1] Tu dong do cot "Ma so" thay vi hard-code cot sai
    Dim CODE_COL As Long
    CODE_COL = DetectCodeCol(ws, firstSelRow, lastSelRow, LABEL_COL)

    ' [FIX 3] Tu dong do cot du lieu dau tien (bo qua cot "Thuyet minh" neu co)
    Dim DATA_START_COL As Long
    DATA_START_COL = DetectDataStartCol(ws, CODE_COL, firstSelRow, lastSelRow)

    ' [FIX 2] Tim cot cuoi cung THUC SU co du lieu tren sheet (khong phu thuoc dong 1)
    Dim DATA_END_COL As Long
    DATA_END_COL = DetectLastUsedCol(ws)

    If DATA_START_COL = 0 Or DATA_END_COL < DATA_START_COL Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Khong xac dinh duoc vung du lieu (cot bat dau/cot ket thuc) tren sheet nay." & vbCrLf & _
               "Da do: CODE_COL=" & CODE_COL & ", DATA_START_COL=" & DATA_START_COL & _
               ", DATA_END_COL=" & DATA_END_COL & "." & vbCrLf & _
               "Vui long kiem tra lai cau truc sheet.", vbCritical
        Exit Sub
    End If
    '----------------------------------------------------------------

    ' Doc lvl/dash/ma so cho toan bo pham vi da chon (chi can trong vung nay)
    Dim i As Long, j As Long, c As Long
    Dim lvl() As Integer, dashRow() As Boolean, codeNum() As Double
    ReDim lvl(firstSelRow To lastSelRow)
    ReDim dashRow(firstSelRow To lastSelRow)
    ReDim codeNum(firstSelRow To lastSelRow)

    Dim lbl As String
    For i = firstSelRow To lastSelRow
        lbl = NormalizeDash(Trim(CStr(ws.Cells(i, LABEL_COL).Value)))
        dashRow(i) = (Left(lbl, 1) = "-")
        lvl(i) = GetLevel(lbl, dashRow(i))
        codeNum(i) = ParseCode(ws.Cells(i, CODE_COL).Value)
    Next i

    ' ----- Khu vuc loai tru theo Ma so (co the them/bot tai day) ----
    ' 220-229: khoi Tai san co dinh (Nguyen gia / Hao mon) -> KHONG tu dong SUM
    ' 400-440: toan bo khoi Von chu so huu -> KHONG tu dong SUM
    ' (Neu ban muon macro tu tinh 2 khoi nay, xoa hoac sua lai 2 dong duoi day)
    Dim exLow(1 To 2) As Double, exHigh(1 To 2) As Double
    exLow(1) = 220: exHigh(1) = 229    ' Tai san co dinh
    exLow(2) = 400: exHigh(2) = 440    ' Von chu so huu (+ tong 440)
    '------------------------------------------------------------------

    Dim countTinh As Long
    countTinh = 0

    ' Duyet tu duoi len tren (trong pham vi da chon) de dong cha
    ' luon duoc tinh SAU dong con
    For i = lastSelRow To firstSelRow Step -1

        If lvl(i) >= 0 And lvl(i) <= 3 And Not IsExcluded(codeNum(i), exLow, exHigh) Then

            Dim childRows As Collection
            Set childRows = New Collection

            j = i + 1
            Do While j <= lastSelRow
                If lvl(j) <> 9 And lvl(j) <= lvl(i) Then Exit Do
                If lvl(j) = lvl(i) + 1 Then childRows.Add j
                j = j + 1
            Loop

            If childRows.Count > 0 Then
                For c = DATA_START_COL To DATA_END_COL
                    Dim refs As String, ck As Variant
                    refs = ""
                    For Each ck In childRows
                        If refs <> "" Then refs = refs & ","
                        refs = refs & ws.Cells(CLng(ck), c).Address(False, False)
                    Next ck
                    ws.Cells(i, c).Formula = "=SUM(" & refs & ")"
                Next c
                countTinh = countTinh + 1
            End If

        End If
    Next i

    ' ----- Truong hop dac biet 1: TONG CONG TAI SAN (280 = 100 + 200) --
    ' (chi ap dung neu ca 3 dong 100, 200, 280 deu nam trong vung da chon)
    Dim r100 As Long, r200 As Long, r280 As Long
    r100 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 100)
    r200 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 200)
    r280 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 280)
    If r100 > 0 And r200 > 0 And r280 > 0 Then
        For c = DATA_START_COL To DATA_END_COL
            ws.Cells(r280, c).Formula = "=SUM(" & _
                ws.Cells(r100, c).Address(False, False) & "," & _
                ws.Cells(r200, c).Address(False, False) & ")"
        Next c
        countTinh = countTinh + 1
    End If

    ' ----- Truong hop dac biet 2 [FIX MOI]: TONG CONG NGUON VON --------
    ' (440 = 300 + 400) - truoc day bi thieu, gio bo sung tuong tu dong 280
    ' (chi ap dung neu ca 3 dong 300, 400, 440 deu nam trong vung da chon)
    Dim r300 As Long, r400 As Long, r440 As Long
    r300 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 300)
    r400 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 400)
    r440 = FindRowByCode(codeNum, firstSelRow, lastSelRow, 440)
    If r300 > 0 And r400 > 0 And r440 > 0 Then
        For c = DATA_START_COL To DATA_END_COL
            ws.Cells(r440, c).Formula = "=SUM(" & _
                ws.Cells(r300, c).Address(False, False) & "," & _
                ws.Cells(r400, c).Address(False, False) & ")"
        Next c
        countTinh = countTinh + 1
    End If
    '------------------------------------------------------------------

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Da tinh xong " & countTinh & " dong tong hop trong vung ban da chon (dong " & _
           firstSelRow & " -> " & lastSelRow & ")." & vbCrLf & _
           "Cot Ma so duoc nhan dien: " & Split(ws.Cells(1, CODE_COL).Address, "$")(1) & _
           " | Vung du lieu: " & Split(ws.Cells(1, DATA_START_COL).Address, "$")(1) & _
           " -> " & Split(ws.Cells(1, DATA_END_COL).Address, "$")(1) & vbCrLf & _
           "Khoi Tai san co dinh (220-229) va Von chu so huu (400-440) neu nam trong vung" & vbCrLf & _
           "van duoc BO QUA (khong bi tu dong SUM), tru rieng dong 280 va 440 duoc tinh" & vbCrLf & _
           "theo cong thuc dac biet (280=100+200, 440=300+400) neu ca 2 dong con va dong" & vbCrLf & _
           "tong deu nam trong vung ban chon.", vbInformation

End Sub

' ================= HAM PHU ======================================

Private Function NormalizeDash(ByVal s As String) As String
    ' [FIX 4] Doi en-dash (Chr 8211) va em-dash (Chr 8212) thanh dau
    ' gach ngang thuong "-" de viec so khop mau nhan cap 0 ("A - ...")
    ' hoat dong dung du nguoi nhap dung loai dau gach ngang nao.
    Dim t As String
    t = Replace(s, Chr(8211), "-")   ' en dash "–"
    t = Replace(t, Chr(8212), "-")   ' em dash "—"
    NormalizeDash = t
End Function

Private Function GetLevel(ByVal lbl As String, ByVal isDash As Boolean) As Integer
    If isDash Then
        GetLevel = 9
        Exit Function
    End If
    If Len(lbl) = 0 Then
        GetLevel = -1
        Exit Function
    End If

    If lbl Like "[A-D] -*" Or lbl Like "[A-D]-*" Then
        GetLevel = 0
        Exit Function
    End If

    Dim firstPart As String
    firstPart = Split(lbl, ".")(0)

    If IsRoman(firstPart) Then
        GetLevel = 1
        Exit Function
    End If

    If lbl Like "[a-z])*" Then
        GetLevel = 3
        Exit Function
    End If

    If IsNumeric(firstPart) Then
        GetLevel = 2
        Exit Function
    End If

    GetLevel = -1
End Function

Private Function IsRoman(ByVal s As String) As Boolean
    Dim validChars As String, k As Long
    validChars = "IVXLCDM"
    If Len(s) = 0 Then Exit Function
    For k = 1 To Len(s)
        If InStr(validChars, Mid(s, k, 1)) = 0 Then Exit Function
    Next k
    IsRoman = True
End Function

Private Function ParseCode(ByVal v As Variant) As Double
    Dim s As String, k As Long, numPart As String
    s = Trim(CStr(v))
    numPart = ""
    For k = 1 To Len(s)
        If Mid(s, k, 1) Like "[0-9]" Then
            numPart = numPart & Mid(s, k, 1)
        Else
            Exit For
        End If
    Next k
    If Len(numPart) > 0 Then
        ParseCode = CDbl(numPart)
    Else
        ParseCode = -1
    End If
End Function

Private Function IsExcluded(ByVal code As Double, ByRef lows() As Double, ByRef highs() As Double) As Boolean
    Dim k As Long
    For k = LBound(lows) To UBound(lows)
        If code >= lows(k) And code <= highs(k) Then
            IsExcluded = True
            Exit Function
        End If
    Next k
    IsExcluded = False
End Function

Private Function FindRowByCode(ByRef codeArr() As Double, ByVal lo As Long, ByVal hi As Long, ByVal target As Double) As Long
    Dim k As Long
    For k = lo To hi
        If codeArr(k) = target Then
            FindRowByCode = k
            Exit Function
        End If
    Next k
    FindRowByCode = 0
End Function

' -----------------------------------------------------------------
' [FIX 2] Tim cot cuoi cung THUC SU co du lieu tren toan bo sheet,
' dung Cells.Find thay vi chi nhin vao dong 1 (dong 1 o nhieu sheet
' chi co 1 o tieu de nen cach lam cu tra ve ket qua sai).
' -----------------------------------------------------------------
Private Function DetectLastUsedCol(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not foundCell Is Nothing Then
        DetectLastUsedCol = foundCell.Column
    Else
        DetectLastUsedCol = ws.UsedRange.Columns.Count
    End If
End Function

' -----------------------------------------------------------------
' [FIX 1] Tu dong tim cot "Ma so": uu tien tim tieu de "Ma so"/"Code"
' trong vai dong dau sheet; neu khong thay, do quet cac cot gan cot
' nhan de tim cot co nhieu gia tri dang "ma so" nhat (vd 110, 111,
' 411a...). Neu van khong xac dinh duoc, mac dinh dung cot ngay ben
' phai cot nhan (truong hop pho bien nhat trong file nay).
' -----------------------------------------------------------------
Private Function DetectCodeCol(ByVal ws As Worksheet, ByVal firstSelRow As Long, _
                                ByVal lastSelRow As Long, ByVal labelCol As Long) As Long
    Dim r As Long, c As Long
    Dim headerText As String

    ' Buoc 1: tim tieu de trong 6 dong dau sheet
    For r = 1 To 6
        For c = labelCol + 1 To labelCol + 5
            headerText = LCase(Trim(CStr(ws.Cells(r, c).Value)))
            headerText = Replace(headerText, Chr(160), " ") ' bo non-breaking space neu co
            If InStr(headerText, "ma so") > 0 Or InStr(headerText, ChrW(227) & " s" & ChrW(7889)) > 0 _
               Or InStr(headerText, "code") > 0 Then
                DetectCodeCol = c
                Exit Function
            End If
        Next c
    Next r

    ' Buoc 2: khong thay tieu de -> dem so o "giong ma so" trong tung cot
    ' (vd: "110", "111", "411a" - chuoi ngan, bat dau bang 1-4 chu so)
    Dim bestCol As Long, bestCount As Long, cnt As Long
    Dim v As String
    bestCol = 0: bestCount = -1
    For c = labelCol + 1 To labelCol + 5
        cnt = 0
        For r = firstSelRow To lastSelRow
            v = Trim(CStr(ws.Cells(r, c).Value))
            If Len(v) > 0 And Len(v) <= 6 Then
                If v Like "[0-9]" Or v Like "[0-9][0-9]" Or v Like "[0-9][0-9][0-9]" _
                   Or v Like "[0-9][0-9][0-9][a-z]" Or v Like "[0-9][0-9][0-9][0-9]" Then
                    cnt = cnt + 1
                End If
            End If
        Next r
        If cnt > bestCount Then
            bestCount = cnt
            bestCol = c
        End If
    Next c

    If bestCount > 0 Then
        DetectCodeCol = bestCol
    Else
        ' Buoc 3: khong xac dinh duoc gi -> mac dinh cot ngay sau cot nhan
        DetectCodeCol = labelCol + 1
    End If
End Function

' -----------------------------------------------------------------
' [FIX 3] Tu dong tim cot du lieu dau tien: bat dau tu ben phai cot
' Ma so, bo qua cac cot chu (vd "Thuyet minh") khong co so lieu,
' chon cot dau tien co ty le o dang so >= 50% trong vung da chon.
' -----------------------------------------------------------------
Private Function DetectDataStartCol(ByVal ws As Worksheet, ByVal codeCol As Long, _
                                     ByVal firstSelRow As Long, ByVal lastSelRow As Long) As Long
    Dim c As Long, r As Long
    Dim numCount As Long, nonBlankCount As Long
    Dim v As Variant

    For c = codeCol + 1 To codeCol + 6
        numCount = 0: nonBlankCount = 0
        For r = firstSelRow To lastSelRow
            v = ws.Cells(r, c).Value
            If Len(Trim(CStr(v))) > 0 Then
                nonBlankCount = nonBlankCount + 1
                If IsNumeric(v) Then numCount = numCount + 1
            End If
        Next r
        If nonBlankCount > 0 Then
            If numCount / nonBlankCount >= 0.5 Then
                DetectDataStartCol = c
                Exit Function
            End If
        End If
    Next c

    ' Khong tim thay cot nao du dieu kien -> mac dinh cot ngay sau cot ma so
    DetectDataStartCol = codeCol + 1
End Function
