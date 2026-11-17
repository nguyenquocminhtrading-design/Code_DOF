Sub CAN_merge()
    Dim wsCust As Worksheet, wsVend As Worksheet, wsResult As Worksheet
    Dim lastRowCust As Long, lastRowVend As Long
    Dim i As Long, destRow As Long

    Const CUST_HEADER_ROW As Long = 1
    Const VEND_HEADER_ROW As Long = 1

    On Error Resume Next
    Set wsCust = ThisWorkbook.Sheets("AR")
    Set wsVend = ThisWorkbook.Sheets("AP")
    On Error GoTo 0

    If wsCust Is Nothing Then
        MsgBox "Không t?m th?y sheet 'AR'. Vui l?ng ki?m tra l?i tên."
        Exit Sub
    End If
    If wsVend Is Nothing Then
        MsgBox "Không t?m th?y sheet 'AP'. Vui l?ng ki?m tra l?i tên."
        Exit Sub
    End If

    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "Merged"

    With wsResult
        .Cells(1, 1).Value = "Kind"
        .Cells(1, 2).Value = "Entity"
        .Cells(1, 3).Value = "Cus.no"
        .Cells(1, 4).Value = "Name"
        .Cells(1, 5).Value = "Ending Balance"
    End With

    destRow = 2

    ' --- L?y d? li?u t? sheet AR ---
    lastRowCust = wsCust.Cells(wsCust.Rows.Count, "A").End(xlUp).Row
    For i = CUST_HEADER_ROW + 1 To lastRowCust
        If Trim(wsCust.Cells(i, 1).Value) <> "" Then
            wsResult.Cells(destRow, 1).Value = "AR"          ' <--- T? T?O KIND
            wsResult.Cells(destRow, 2).Value = "CAN"         ' Entity
            wsResult.Cells(destRow, 3).Value = wsCust.Cells(i, 1).Value
            wsResult.Cells(destRow, 4).Value = wsCust.Cells(i, 2).Value
            wsResult.Cells(destRow, 5).Value = wsCust.Cells(i, 5).Value
            destRow = destRow + 1
        End If
    Next i

    ' --- L?y d? li?u t? sheet AP ---
    lastRowVend = wsVend.Cells(wsVend.Rows.Count, "A").End(xlUp).Row
    For i = VEND_HEADER_ROW + 1 To lastRowVend
        If Trim(wsVend.Cells(i, 1).Value) <> "" Then
            wsResult.Cells(destRow, 1).Value = "AP"          ' <--- T? T?O KIND
            wsResult.Cells(destRow, 2).Value = "CAN"         ' Entity
            wsResult.Cells(destRow, 3).Value = wsVend.Cells(i, 1).Value
            wsResult.Cells(destRow, 4).Value = wsVend.Cells(i, 2).Value
            wsResult.Cells(destRow, 5).Value = wsVend.Cells(i, 4).Value
            destRow = destRow + 1
        End If
    Next i

    wsResult.Columns("A:E").AutoFit
    MsgBox "Ð? g?p xong! K?t qu? n?m trong sheet '" & wsResult.Name & "'."
End Sub
