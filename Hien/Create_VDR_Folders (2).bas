Option Explicit

' ============================================================
'  MACRO: Create_VDR_Folder_Structure
'  Cau truc cot theo dung anh ban gui (KHONG dua vao cot H/G nua):
'     Cot A = so Section    (cap 1)   -> vd 1
'     Cot B = so Subsection (cap 2)   -> vd 1
'     Cot C = so Item       (cap 3)   -> vd 1
'     Cot I = ten VDR topics          (dung cho ten cap 1 va cap 2)
'     Cot J = ten Key documents       (dung cho ten cap 3)
'  Header o dong 6, du lieu bat dau tu dong 7.
'
'  Cach xac dinh cap cua 1 dong (dua thang vao A/B/C, ro rang va
'  on dinh, khong can do H hay "nhin xuong" dong duoi nua):
'     - Cap 1 (Section):    A co gia tri, B rong,      C rong
'     - Cap 2 (Subsection): A co gia tri, B co gia tri, C rong
'     - Cap 3 (Item):       A co gia tri, B co gia tri, C co gia tri
'
'  Ket qua tao thu muc, so duoc dem 2 chu so (01, 02 ... 15) de
'  Windows Explorer sap xep DUNG THU TU:
'     01 - General corporate/
'       01.01 - Group overview & structure/
'         01.01.01 - Legal entity org structure/
' ============================================================

Sub Create_VDR_Folder_Structure()

    Dim ws As Worksheet
    Dim rootPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim lvl As Integer

    Dim secNum As Long, secName As String, secCode As String
    Dim subNum As Long, subName As String, subCode As String
    Dim itemNum As Long, itemName As String, itemCode As String
    Dim path1 As String, path2 As String, path3 As String
    Dim folderName As String
    Dim countCreated As Long
    Dim skippedRows As String
    Dim skipCount As Long

    ' >>> Doi ten sheet o day neu file cua ban dat ten khac <<<
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("VDR Tracker")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Khong tim thay sheet 'VDR Tracker'. Vui long kiem tra lai ten sheet.", vbCritical
        Exit Sub
    End If

    ' --- Chon thu muc goc de tao cay thu muc ---
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Chon thu muc goc de tao cau truc VDR"
    If fd.Show <> -1 Then
        MsgBox "Da huy - chua chon thu muc goc.", vbExclamation
        Exit Sub
    End If
    rootPath = fd.SelectedItems(1)
    If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"

    lastRow = GetLastDataRow(ws)
    path1 = "": path2 = ""
    countCreated = 0
    skipCount = 0

    For i = 7 To lastRow
        lvl = ClassifyRow(ws, i)

        Select Case lvl

            Case 1  ' ---- Section (cap 1): A co gia tri, B/C rong ----
                secNum = CLng(ws.Cells(i, "A").Value)
                secName = Trim(CStr(ws.Cells(i, "I").Value))
                secCode = Format(secNum, "00")
                folderName = SanitizeName(secCode & " - " & secName)
                path1 = rootPath & folderName
                If EnsureFolder(path1) Then countCreated = countCreated + 1
                path2 = ""

            Case 2  ' ---- Subsection (cap 2): A va B co gia tri, C rong ----
                secNum = CLng(ws.Cells(i, "A").Value)
                subNum = CLng(ws.Cells(i, "B").Value)
                subName = Trim(CStr(ws.Cells(i, "I").Value))
                subCode = Format(secNum, "00") & "." & Format(subNum, "00")
                folderName = SanitizeName(subCode & " - " & subName)
                If path1 = "" Then path1 = rootPath  ' du phong an toan
                path2 = path1 & "\" & folderName
                If EnsureFolder(path2) Then countCreated = countCreated + 1

            Case 3  ' ---- Item (cap 3): A, B, C deu co gia tri ----
                secNum = CLng(ws.Cells(i, "A").Value)
                subNum = CLng(ws.Cells(i, "B").Value)
                itemNum = CLng(ws.Cells(i, "C").Value)
                itemName = Trim(CStr(ws.Cells(i, "J").Value))
                itemCode = Format(secNum, "00") & "." & Format(subNum, "00") & "." & Format(itemNum, "00")
                folderName = SanitizeName(itemCode & " - " & itemName)
                If path2 = "" Then path2 = path1
                path3 = path2 & "\" & folderName
                If EnsureFolder(path3) Then countCreated = countCreated + 1

            Case Else
                ' Dong khong xac dinh duoc cap (thieu du lieu A/B/C hoac trong)
                If Not IsRowFullyEmpty(ws, i) Then
                    skipCount = skipCount + 1
                    If skipCount <= 15 Then
                        skippedRows = skippedRows & i & ", "
                    End If
                End If

        End Select
    Next i

    Dim summaryMsg As String
    summaryMsg = "Hoan tat!" & vbCrLf & "Da tao " & countCreated & " thu muc moi tai:" & vbCrLf & rootPath
    If skipCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & vbCrLf & "Luu y: co " & skipCount & _
            " dong khong xac dinh duoc cap (thieu du lieu o cot A/B/C), da bo qua. Dong: " & skippedRows
    End If
    MsgBox summaryMsg, vbInformation

End Sub


' ============================================================
'                       HAM PHU TRO
' ============================================================

' Phan loai 1 dong du lieu dua THANG vao cot A, B, C:
'   1 = Section (A co, B/C rong)
'   2 = Subsection (A,B co, C rong)
'   3 = Item (A,B,C deu co)
'   0 = khong xac dinh / dong trong -> bo qua
Function ClassifyRow(ws As Worksheet, r As Long) As Integer
    Dim a As Variant, b As Variant, c As Variant
    a = ws.Cells(r, "A").Value
    b = ws.Cells(r, "B").Value
    c = ws.Cells(r, "C").Value

    If Not IsEmptyCell(a) And IsEmptyCell(b) And IsEmptyCell(c) Then
        ClassifyRow = 1
    ElseIf Not IsEmptyCell(a) And Not IsEmptyCell(b) And IsEmptyCell(c) Then
        ClassifyRow = 2
    ElseIf Not IsEmptyCell(a) And Not IsEmptyCell(b) And Not IsEmptyCell(c) Then
        ClassifyRow = 3
    Else
        ClassifyRow = 0
    End If
End Function

Function IsEmptyCell(v As Variant) As Boolean
    If IsNull(v) Then
        IsEmptyCell = True
    ElseIf VarType(v) = vbEmpty Then
        IsEmptyCell = True
    ElseIf VarType(v) = vbString Then
        IsEmptyCell = (Trim(v) = "")
    Else
        IsEmptyCell = False
    End If
End Function

' Kiem tra 1 dong co hoan toan trong khong (de phan biet dong "bo qua that su"
' voi dong "thieu du lieu can canh bao")
Function IsRowFullyEmpty(ws As Worksheet, r As Long) As Boolean
    Dim c As Long
    Dim allEmpty As Boolean
    allEmpty = True
    For c = 1 To 10 ' cot A den J
        If Not IsEmptyCell(ws.Cells(r, c).Value) Then
            allEmpty = False
            Exit For
        End If
    Next c
    IsRowFullyEmpty = allEmpty
End Function

' Tim dong cuoi cung co du lieu (dua tren cot A, I, J)
Function GetLastDataRow(ws As Worksheet) As Long
    Dim r1 As Long, r2 As Long, r3 As Long
    r1 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    r2 = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    r3 = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    GetLastDataRow = Application.WorksheetFunction.Max(r1, r2, r3)
End Function

' Tao thu muc neu chua ton tai. Tra ve True neu vua tao moi.
Function EnsureFolder(ByVal p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    If Dir(p, vbDirectory) = "" Then
        On Error Resume Next
        MkDir p
        On Error GoTo 0
        EnsureFolder = True
    Else
        EnsureFolder = False
    End If
End Function

' Loai bo ky tu khong hop le trong ten thu muc Windows
Function SanitizeName(ByVal s As String) As String
    Dim badChars As String
    Dim i As Integer
    badChars = "\/:*?" & Chr(34) & "<>|"
    For i = 1 To Len(badChars)
        s = Replace(s, Mid(badChars, i, 1), "-")
    Next i
    Do While Len(s) > 0 And (Right(s, 1) = " " Or Right(s, 1) = ".")
        s = Left(s, Len(s) - 1)
    Loop
    If Len(s) > 100 Then s = Left(s, 100) ' tranh ten qua dai gay loi duong dan
    SanitizeName = s
End Function
