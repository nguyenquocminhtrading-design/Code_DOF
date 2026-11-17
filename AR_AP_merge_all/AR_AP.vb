' ============================================================
'  HELPER FUNCTION: CONVERT VALUE TO NUMBER IF POSSIBLE
' ============================================================
Private Function ConvertToNumber(ByVal val As Variant) As Variant
    ' Check if value is Nothing or Empty
    If IsNull(val) Or IsEmpty(val) Then
        ConvertToNumber = val
        Exit Function
    End If

    ' Convert to string and remove non-printable characters + trim
    Dim strVal As String
    strVal = Application.Clean(Trim(val))
    strVal = Trim(strVal)

    If strVal = "" Then
        ConvertToNumber = strVal
        Exit Function
    End If

    ' Try to convert to number
    If IsNumeric(strVal) Then
        ConvertToNumber = CDbl(strVal)
    Else
        ConvertToNumber = strVal
    End If
End Function

' ============================================================
'  HELPER FUNCTION: GET WORKBOOK (OPEN IF NOT ALREADY OPEN)
' ============================================================
Private Function GetWorkbook(ByVal filePath As String, ByRef currentWB As Workbook) As Workbook
    If filePath = "" Then
        Set GetWorkbook = currentWB
    Else
        On Error Resume Next
        Set GetWorkbook = Workbooks(filePath)
        If GetWorkbook Is Nothing Then
            Set GetWorkbook = Workbooks.Open(filePath)
        End If
        On Error GoTo 0
    End If
End Function

' ============================================================
'  MAIN MACRO: MERGE THREE SOURCES WITH ACCOUNT COLUMN
' ============================================================
Sub MergeThreeSourcesWithAccount_Fixed()
    ' ------------------------------------------------------------
    '  CONFIGURATION FOR SOURCE 1 (AR from CAN Report)
    ' ------------------------------------------------------------
    Const SOURCE1_FILE As String = "C:\Users\Acer\Desktop\Dan on food\DOF(Dan on Food)\Company FS structure consolidatede\Auto_consol\CAN.xlsm"
    Const SOURCE1_SHEET As String = "Merged"
    Const SOURCE1_HEADER_ROW As Long = 1
    Const SOURCE1_COL_KIND As Long = 1
    Const SOURCE1_COL_ENTITY As Long = 2
    Const SOURCE1_COL_NAME As Long = 3
    Const SOURCE1_COL_CURRENT As Long = 4

    ' ------------------------------------------------------------
    '  CONFIGURATION FOR SOURCE 2 (AP from DOAU)
    ' ------------------------------------------------------------
    Const SOURCE2_FILE As String = "C:\Users\Acer\Desktop\Dan on food\DOF(Dan on Food)\Company FS structure consolidatede\Auto_consol\DOAU.xlsm"
    Const SOURCE2_SHEET As String = "Merged"
    Const SOURCE2_HEADER_ROW As Long = 1
    Const SOURCE2_COL_KIND As Long = 1
    Const SOURCE2_COL_ENTITY As Long = 2
    Const SOURCE2_COL_NAME As Long = 3
    Const SOURCE2_COL_CURRENT As Long = 4

    ' ------------------------------------------------------------
    '  CONFIGURATION FOR SOURCE 3 (AR AP loan)
    ' ------------------------------------------------------------
    Const SOURCE3_FILE As String = "C:\Users\Acer\Desktop\Dan on food\DOF(Dan on Food)\Company FS structure consolidatede\AR AP eli\AR AP loan  - DOF ADF - 2025.05.xlsx"
    Const SOURCE3_SHEET As String = "AR_AP_DOF"
    Const SOURCE3_HEADER_ROW As Long = 5
    Const SOURCE3_COL_KIND As Long = 1
    Const SOURCE3_COL_ENTITY As Long = 2
    Const SOURCE3_COL_NAME As Long = 4
    Const SOURCE3_COL_CURRENT As Long = 5
    Const SOURCE3_COL_ACCOUNT As Long = 6

    ' ------------------------------------------------------------
    '  RESULT SHEET NAME
    ' ------------------------------------------------------------
    Const RESULT_SHEET As String = "Merged"

    ' ------------------------------------------------------------
    '  START PROCESSING
    ' ------------------------------------------------------------
    Dim wb1 As Workbook, wb2 As Workbook, wb3 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, wsRes As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, destRow As Long
    Dim currentWB As Workbook

    Set currentWB = ThisWorkbook

    ' ---- Get Source 1 ----
    Set wb1 = GetWorkbook(SOURCE1_FILE, currentWB)
    If wb1 Is Nothing Then
        MsgBox "Cannot access Source 1. Please check the file path."
        Exit Sub
    End If
    On Error Resume Next
    Set ws1 = wb1.Sheets(SOURCE1_SHEET)
    On Error GoTo 0
    If ws1 Is Nothing Then
        MsgBox "Sheet '" & SOURCE1_SHEET & "' not found in Source 1."
        Exit Sub
    End If

    ' ---- Get Source 2 ----
    Set wb2 = GetWorkbook(SOURCE2_FILE, currentWB)
    If wb2 Is Nothing Then
        MsgBox "Cannot access Source 2. Please check the file path."
        Exit Sub
    End If
    On Error Resume Next
    Set ws2 = wb2.Sheets(SOURCE2_SHEET)
    On Error GoTo 0
    If ws2 Is Nothing Then
        MsgBox "Sheet '" & SOURCE2_SHEET & "' not found in Source 2."
        Exit Sub
    End If

    ' ---- Get Source 3 ----
    Set wb3 = GetWorkbook(SOURCE3_FILE, currentWB)
    If wb3 Is Nothing Then
        MsgBox "Cannot access Source 3. Please check the file path."
        Exit Sub
    End If
    On Error Resume Next
    Set ws3 = wb3.Sheets(SOURCE3_SHEET)
    On Error GoTo 0
    If ws3 Is Nothing Then
        MsgBox "Sheet '" & SOURCE3_SHEET & "' not found in Source 3."
        Exit Sub
    End If

    ' ---- Delete existing result sheet and create new ----
    On Error Resume Next
    Application.DisplayAlerts = False
    currentWB.Sheets(RESULT_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsRes = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.Count))
    wsRes.Name = RESULT_SHEET

    ' ---- Write headers (5 columns) ----
    With wsRes
        .Cells(1, 1).Value = "Kind"
        .Cells(1, 2).Value = "Entity"
        .Cells(1, 3).Value = "Name"
        .Cells(1, 4).Value = "Ending Balance"
        .Cells(1, 5).Value = "Account"
    End With

    ' ---- Format columns ----
    wsRes.Columns("A:C").NumberFormat = "@"        ' Text for Kind, Entity, Name
    wsRes.Columns("D:E").NumberFormat = "General"  ' Numbers for Balance and Account

    destRow = 2

    ' ---- Copy from Source 1 (no Account) ----
    lastRow1 = ws1.Cells(ws1.Rows.Count, SOURCE1_COL_KIND).End(xlUp).Row
    For i = SOURCE1_HEADER_ROW + 1 To lastRow1
        If Trim(ws1.Cells(i, SOURCE1_COL_KIND).Value) <> "" Then
            wsRes.Cells(destRow, 1).Value = ws1.Cells(i, SOURCE1_COL_KIND).Value
            wsRes.Cells(destRow, 2).Value = ws1.Cells(i, SOURCE1_COL_ENTITY).Value
            wsRes.Cells(destRow, 3).Value = ws1.Cells(i, SOURCE1_COL_NAME).Value
            wsRes.Cells(destRow, 4).Value = ConvertToNumber(ws1.Cells(i, SOURCE1_COL_CURRENT).Value)
            wsRes.Cells(destRow, 5).Value = ""   ' leave Account empty
            destRow = destRow + 1
        End If
    Next i

    ' ---- Copy from Source 2 (no Account) ----
    lastRow2 = ws2.Cells(ws2.Rows.Count, SOURCE2_COL_KIND).End(xlUp).Row
    For i = SOURCE2_HEADER_ROW + 1 To lastRow2
        If Trim(ws2.Cells(i, SOURCE2_COL_KIND).Value) <> "" Then
            wsRes.Cells(destRow, 1).Value = ws2.Cells(i, SOURCE2_COL_KIND).Value
            wsRes.Cells(destRow, 2).Value = ws2.Cells(i, SOURCE2_COL_ENTITY).Value
            wsRes.Cells(destRow, 3).Value = ws2.Cells(i, SOURCE2_COL_NAME).Value
            wsRes.Cells(destRow, 4).Value = ConvertToNumber(ws2.Cells(i, SOURCE2_COL_CURRENT).Value)
            wsRes.Cells(destRow, 5).Value = ""
            destRow = destRow + 1
        End If
    Next i

    ' ---- Copy from Source 3 (with Account) ----
    lastRow3 = ws3.Cells(ws3.Rows.Count, SOURCE3_COL_KIND).End(xlUp).Row
    For i = SOURCE3_HEADER_ROW + 1 To lastRow3
        If Trim(ws3.Cells(i, SOURCE3_COL_KIND).Value) <> "" Then
            wsRes.Cells(destRow, 1).Value = ws3.Cells(i, SOURCE3_COL_KIND).Value
            wsRes.Cells(destRow, 2).Value = ws3.Cells(i, SOURCE3_COL_ENTITY).Value
            wsRes.Cells(destRow, 3).Value = ws3.Cells(i, SOURCE3_COL_NAME).Value
            wsRes.Cells(destRow, 4).Value = ConvertToNumber(ws3.Cells(i, SOURCE3_COL_CURRENT).Value)
            wsRes.Cells(destRow, 5).Value = ConvertToNumber(ws3.Cells(i, SOURCE3_COL_ACCOUNT).Value)
            destRow = destRow + 1
        End If
    Next i

    ' ---- Auto-fit columns ----
    wsRes.Columns("A:E").AutoFit

    MsgBox "Merge completed! Result is in sheet '" & RESULT_SHEET & "' of the current workbook."
End Sub
