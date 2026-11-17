Option Explicit
Sub BuildFSDimension()
    Dim wsRaw As Worksheet, wsOut As Worksheet
    Dim lastRow As Long
    Dim colCompany As Long, colFS As Long, colAmt As Long, colCondition As Long
    Dim dictCompanies As Object, dictFS As Object
    Dim r As Long, k As Variant
    Dim outR As Long, outC As Long

    Set wsRaw = ActiveSheet

    ' --- Tim cac cot can thiet ---
    colCompany = FindHeader(wsRaw, "Company")
    colFS = FindHeader(wsRaw, "FS code")
    colAmt = FindHeader(wsRaw, "Transaction Amount")
    colCondition = FindHeader(wsRaw, "Condition")

    If colCompany = 0 Or colFS = 0 Or colAmt = 0 Then
        MsgBox "Khong tim thay cot. Can co: 'Company', 'FS code', 'Transaction Amount'."
        Exit Sub
    End If

    ' --- Dong cuoi that su co du lieu (quet nguoc) ---
    lastRow = 0
    On Error Resume Next
    lastRow = wsRaw.Cells.Find(What:="*", LookIn:=xlFormulas, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lastRow < 2 Then
        MsgBox "Khong co du lieu."
        Exit Sub
    End If

    ' --- Thu thap unique Company (giu thu tu xuat hien) va FS code ---
    Set dictCompanies = CreateObject("Scripting.Dictionary")
    dictCompanies.CompareMode = vbTextCompare
    Set dictFS = CreateObject("Scripting.Dictionary")

    Dim cName As String, fsCode As String
    For r = 2 To lastRow
        cName = Trim(CStr(wsRaw.Cells(r, colCompany).Value))
        fsCode = Trim(CStr(wsRaw.Cells(r, colFS).Value))
        If Len(cName) > 0 Then
            If Not dictCompanies.Exists(cName) Then dictCompanies.Add cName, True
        End If
        If Len(fsCode) > 0 Then
            If Not dictFS.Exists(fsCode) Then dictFS.Add fsCode, True
        End If
    Next r

    If dictCompanies.Count = 0 Or dictFS.Count = 0 Then
        MsgBox "Khong co du lieu hop le."
        Exit Sub
    End If

    ' --- Sort FS codes (so thi sort so, chuoi thi sort chuoi) ---
    Dim arrFS() As Variant
    arrFS = dictFS.Keys
    Call SortVariantArray(arrFS)

    ' --- Tao / reset sheet output ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("FS_Dimension").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsRaw)
    wsOut.Name = "FS_Dimension"

    ' --- Header ---
    wsOut.Cells(1, 1).Value = "FS code"
    outC = 2
    For Each k In dictCompanies.Keys
        wsOut.Cells(1, outC).Value = k
        outC = outC + 1
    Next k
    wsOut.Cells(1, outC).Value = "Total"
    Dim totalCol As Long
    totalCol = outC

    ' --- Chuan bi dia chi cot nguon ---
    Dim srcName As String
    srcName = "'" & wsRaw.Name & "'"
    Dim colComLetter As String, colFSLetter As String, colAmtLetter As String
    colComLetter = Split(wsRaw.Cells(1, colCompany).Address(True, False), "$")(0)
    colFSLetter = Split(wsRaw.Cells(1, colFS).Address(True, False), "$")(0)
    colAmtLetter = Split(wsRaw.Cells(1, colAmt).Address(True, False), "$")(0)

    Dim extraCriteria As String
    If colCondition > 0 Then
        Dim colCondLetter As String
        colCondLetter = Split(wsRaw.Cells(1, colCondition).Address(True, False), "$")(0)
        extraCriteria = "," & srcName & "!" & colCondLetter & ":" & colCondLetter & ",TRUE"
    Else
        extraCriteria = ""
    End If

    ' --- Duyet tung FS code, sinh SUMIFS cho tung entity ---
    outR = 2
    For Each k In arrFS
        wsOut.Cells(outR, 1).Value = k
        For outC = 2 To totalCol - 1
            wsOut.Cells(outR, outC).Formula = _
                "=SUMIFS(" & srcName & "!" & colAmtLetter & ":" & colAmtLetter & _
                "," & srcName & "!" & colComLetter & ":" & colComLetter & _
                "," & wsOut.Cells(1, outC).Address(True, True) & _
                "," & srcName & "!" & colFSLetter & ":" & colFSLetter & _
                ",$A" & outR & extraCriteria & ")"
        Next outC
        ' Total row
        wsOut.Cells(outR, totalCol).Formula = _
            "=SUM(" & wsOut.Cells(outR, 2).Address(False, False) & ":" & _
            wsOut.Cells(outR, totalCol - 1).Address(False, False) & ")"
        outR = outR + 1
    Next k

    ' --- Grand total row ---
    Dim grandRow As Long
    grandRow = outR
    wsOut.Cells(grandRow, 1).Value = "Total"
    For outC = 2 To totalCol
        wsOut.Cells(grandRow, outC).Formula = _
            "=SUM(" & wsOut.Cells(2, outC).Address(False, False) & ":" & _
            wsOut.Cells(grandRow - 1, outC).Address(False, False) & ")"
    Next outC

    ' --- Format ---
    With wsOut
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(217, 217, 217)
        .Rows(1).HorizontalAlignment = xlCenter
        .Rows(grandRow).Font.Bold = True
        .Rows(grandRow).Interior.Color = RGB(242, 242, 242)
        .Columns(1).ColumnWidth = 10
        .Columns.AutoFit
        .Range(.Cells(2, 2), .Cells(grandRow, totalCol)).NumberFormat = _
            "#,##0;-#,##0;""-"""
        ' Dong bang cho de nhin
        .Range(.Cells(1, 1), .Cells(grandRow, totalCol)).Borders.LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(grandRow, totalCol)).Borders.Color = RGB(191, 191, 191)
    End With

    MsgBox "Da tao sheet 'FS_Dimension'!" & vbCrLf & _
           "Entities: " & dictCompanies.Count & " | FS codes: " & dictFS.Count, _
           vbInformation
End Sub


'============================================================
' Helpers
'============================================================
Function FindHeader(ws As Worksheet, headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim(CStr(ws.Cells(1, c).Value)), headerName, vbTextCompare) = 0 Then
            FindHeader = c
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

Private Sub SortVariantArray(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    Dim isNumAll As Boolean
    isNumAll = True
    For i = LBound(arr) To UBound(arr)
        If Not IsNumeric(arr(i)) Then
            isNumAll = False
            Exit For
        End If
    Next i

    Dim doSwap As Boolean
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            doSwap = False
            If isNumAll Then
                If CDbl(arr(i)) > CDbl(arr(j)) Then doSwap = True
            Else
                If CStr(arr(i)) > CStr(arr(j)) Then doSwap = True
            End If
            If doSwap Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub