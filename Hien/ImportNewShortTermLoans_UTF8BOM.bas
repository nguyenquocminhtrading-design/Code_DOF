Option Explicit

'===================== CONFIG =====================
Private Const LOAN_FILE_PATH As String = "C:\Users\Acer\Desktop\Dan on food\Loans\me\Loan.xlsm"
Private Const LOAN_SHEET_NAME As String = "ST Loan"
Private Const LOAN_TABLE_NAME As String = "Short_term_loans"

Private Const CONV_SHEET_NAME As String = "Sheet1"
Private Const CONV_TABLE_NAME As String = "Table1"

' Tên c?t trong b?ng Conversion (Table1) - Đ? S?A THEO TÊN TH?C T?
Private Const COL_ENTITY As String = "Entity"
Private Const COL_BANK As String = "Bank"
Private Const COL_CURR As String = "Original Currency"
Private Const COL_PRINCIPAL As String = "Loan Amount "      ' Có 1 d?u cách cu?i
Private Const COL_ISSUE As String = "Issue Date"
Private Const COL_MATURITY As String = "Maturity Date"
Private Const COL_RATE As String = "Interest Rate (%/nam)"  ' Đ? S?A: %/nam
Private Const COL_CHECK As String = "Check ngày"            ' Đ? S?A: b? d?u cách cu?i

' Tên c?t trong b?ng Short_term_loans - Đ? S?A THEO TÊN TH?C T?
Private Const LOAN_COL_ENTITY As String = "Entity"
Private Const LOAN_COL_BANK As String = "Bank"
Private Const LOAN_COL_CURR As String = "Original Currency"
Private Const LOAN_COL_PRINCIPAL As String = "Loan Amount " ' Có 1 d?u cách cu?i
Private Const LOAN_COL_ISSUE As String = "Issue Date"
Private Const LOAN_COL_MATURITY As String = "Maturity Date"
Private Const LOAN_COL_RATE As String = "Interest Rate (%/nam)"  ' Đ? S?A: %/nam
Private Const LOAN_COL_FLAG As String = "Type"

Private Const LOAN_FLAG_VALUE As String = "Ngày thanh toán"
Private Const AUTO_SAVE_LOAN As Boolean = True
Private Const MARK_FILL As Long = 13561798
Private Const MARK_FONT As Long = 3937500

'======================================================================

Sub ImportNewShortTermLoans()
    Dim wbConv As Workbook, wbLoan As Workbook
    Dim wsConv As Worksheet, wsLoan As Worksheet
    Dim loConv As ListObject, loLoan As ListObject
    Dim rw As Range, rngTemplate As Range, newRows As Range
    Dim i As Long, r As Long
    Dim colEntityC As Long, colBankC As Long, colCurrC As Long, colPrinC As Long
    Dim colIssueC As Long, colMatC As Long, colRateC As Long, colCheckC As Long
    Dim sEntity As String, sBank As String, sCurr As String
    Dim dPrincipal As Double, dIssue As Date, dMat As Date, dRate As Double
    Dim dictKeys As Object
    Dim sKey As String
    Dim nAdded As Long, nSkipped As Long, nDup As Long
    Dim origApp As Boolean, origCalc As XlCalculation

    On Error GoTo ErrHandler

    origApp = Application.ScreenUpdating
    origCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbConv = ThisWorkbook
    Set wsConv = wbConv.Sheets(CONV_SHEET_NAME)
    Set loConv = wsConv.ListObjects(CONV_TABLE_NAME)

    wsConv.Calculate

    '===== KI?M TRA TÊN C?T (Đ? CHUY?N VÀO ĐÂY) =====
    On Error Resume Next
    colEntityC = loConv.ListColumns(COL_ENTITY).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_ENTITY: GoTo CleanExit

    colBankC = loConv.ListColumns(COL_BANK).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_BANK: GoTo CleanExit

    colCurrC = loConv.ListColumns(COL_CURR).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_CURR: GoTo CleanExit

    colPrinC = loConv.ListColumns(COL_PRINCIPAL).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_PRINCIPAL: GoTo CleanExit

    colIssueC = loConv.ListColumns(COL_ISSUE).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_ISSUE: GoTo CleanExit

    colMatC = loConv.ListColumns(COL_MATURITY).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_MATURITY: GoTo CleanExit

    colRateC = loConv.ListColumns(COL_RATE).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_RATE: GoTo CleanExit

    colCheckC = loConv.ListColumns(COL_CHECK).Index
    If Err.Number <> 0 Then MsgBox "L?i c?t: " & COL_CHECK: GoTo CleanExit
    On Error GoTo 0
    '================================================

    Set wbLoan = GetOrOpenLoanWorkbook()
    If wbLoan Is Nothing Then
        MsgBox "Cannot open Loan file at: " & vbCrLf & LOAN_FILE_PATH & vbCrLf & _
               "Please check the path or open the file manually.", vbCritical
        GoTo CleanExit
    End If

    Set wsLoan = wbLoan.Sheets(LOAN_SHEET_NAME)
    Set loLoan = wsLoan.ListObjects(LOAN_TABLE_NAME)

    If loLoan.DataBodyRange Is Nothing Or loLoan.DataBodyRange.Rows.Count < 3 Then
        MsgBox "Table '" & LOAN_TABLE_NAME & "' in Loan file has insufficient data (need at least 3 rows for template).", vbCritical
        GoTo CleanExit
    End If

    Set rngTemplate = loLoan.DataBodyRange.Rows(loLoan.DataBodyRange.Rows.Count - 2).Resize(3)
    Set dictKeys = CreateObject("Scripting.Dictionary")
    BuildExistingKeys loLoan, dictKeys

    nAdded = 0: nSkipped = 0: nDup = 0

    If Not loConv.DataBodyRange Is Nothing Then
        For r = 1 To loConv.DataBodyRange.Rows.Count
            Set rw = loConv.DataBodyRange.Rows(r)

            sEntity = CStr(rw.Cells(1, colEntityC).Value)
            If Len(Trim(sEntity)) = 0 Then GoTo NextRow

            ' Already processed (marked)
            If rw.Cells(1, colEntityC).Interior.Color = MARK_FILL Then GoTo NextRow

            ' Only process rows with #N/A error in Check column
            If Not IsError(rw.Cells(1, colCheckC).Value) Then GoTo NextRow

            sBank = CStr(rw.Cells(1, colBankC).Value)
            sCurr = CStr(rw.Cells(1, colCurrC).Value)
            dPrincipal = CDbl(rw.Cells(1, colPrinC).Value)
            dIssue = CDate(rw.Cells(1, colIssueC).Value)
            dMat = CDate(rw.Cells(1, colMatC).Value)
            dRate = CDbl(rw.Cells(1, colRateC).Value)

            sKey = BuildKey(sEntity, sBank, sCurr, dPrincipal, dIssue)

            If dictKeys.Exists(sKey) Then
                MarkRow rw
                nDup = nDup + 1
                GoTo NextRow
            End If

            ' Add 3 rows to Loan table
            loLoan.ListRows.Add AlwaysInsert:=True
            loLoan.ListRows.Add AlwaysInsert:=True
            loLoan.ListRows.Add AlwaysInsert:=True

            Set newRows = loLoan.DataBodyRange.Rows(loLoan.DataBodyRange.Rows.Count - 2).Resize(3)

            ' Copy format and formulas from template
            rngTemplate.Copy
            newRows.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            ' Paste data into all 3 rows
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_ENTITY).DataBodyRange).Value = sEntity
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_BANK).DataBodyRange).Value = sBank
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_CURR).DataBodyRange).Value = sCurr
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_PRINCIPAL).DataBodyRange).Value = dPrincipal
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_ISSUE).DataBodyRange).Value = dIssue
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_MATURITY).DataBodyRange).Value = dMat
            Intersect(newRows, loLoan.ListColumns(LOAN_COL_RATE).DataBodyRange).Value = dRate

            ' Ensure correct flag names
            newRows.Cells(1, loLoan.ListColumns(LOAN_COL_FLAG).Index).Value = "Ngày thanh toán"
            newRows.Cells(2, loLoan.ListColumns(LOAN_COL_FLAG).Index).Value = "Chi tr? phí L?i vay"
            newRows.Cells(3, loLoan.ListColumns(LOAN_COL_FLAG).Index).Value = "Chi tr? n? g?c vay"

            dictKeys.Add sKey, True
            MarkRow rw
            nAdded = nAdded + 1

NextRow:
        Next r
    End If

    Application.Calculation = xlCalculationAutomatic
    wbLoan.Application.CalculateFullRebuild

    If AUTO_SAVE_LOAN Then wbLoan.Save

CleanExit:
    Application.ScreenUpdating = origApp
    Application.EnableEvents = True
    Application.Calculation = origCalc

    MsgBox "Completed." & vbCrLf & _
           "- New loans added to ST Loan: " & nAdded & vbCrLf & _
           "- Duplicate loans (already exist, skipped): " & nDup, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'--------------------------------------------------------------------
Private Function GetOrOpenLoanWorkbook() As Workbook
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In Application.Workbooks
        If Not wb.Sheets(LOAN_SHEET_NAME) Is Nothing Then
            Set GetOrOpenLoanWorkbook = wb
            Exit Function
        End If
    Next wb
    On Error GoTo 0

    If Len(Dir(LOAN_FILE_PATH)) = 0 Then Exit Function
    Set GetOrOpenLoanWorkbook = Workbooks.Open(LOAN_FILE_PATH)
End Function

'--------------------------------------------------------------------
Private Sub BuildExistingKeys(loLoan As ListObject, dictKeys As Object)
    Dim i As Long, n As Long
    Dim colEntity As Long, colBank As Long, colCurr As Long, colPrin As Long, colIssue As Long, colFlag As Long
    Dim sKey As String

    If loLoan.DataBodyRange Is Nothing Then Exit Sub

    colEntity = loLoan.ListColumns(LOAN_COL_ENTITY).Index
    colBank = loLoan.ListColumns(LOAN_COL_BANK).Index
    colCurr = loLoan.ListColumns(LOAN_COL_CURR).Index
    colPrin = loLoan.ListColumns(LOAN_COL_PRINCIPAL).Index
    colIssue = loLoan.ListColumns(LOAN_COL_ISSUE).Index
    colFlag = loLoan.ListColumns(LOAN_COL_FLAG).Index

    n = loLoan.DataBodyRange.Rows.Count
    For i = 1 To n
        With loLoan.DataBodyRange.Rows(i)
            If CStr(.Cells(1, colFlag).Value) = LOAN_FLAG_VALUE Then
                sKey = BuildKey(CStr(.Cells(1, colEntity).Value), CStr(.Cells(1, colBank).Value), _
                                CStr(.Cells(1, colCurr).Value), CDbl(.Cells(1, colPrin).Value), _
                                CDate(.Cells(1, colIssue).Value))
                If Not dictKeys.Exists(sKey) Then dictKeys.Add sKey, True
            End If
        End With
    Next i
End Sub

'--------------------------------------------------------------------
Private Function BuildKey(sEntity As String, sBank As String, sCurr As String, _
                           dPrincipal As Double, dIssue As Date) As String
    BuildKey = UCase(Trim(sEntity)) & "|" & UCase(Trim(sBank)) & "|" & UCase(Trim(sCurr)) & _
               "|" & Format(dPrincipal, "0.00") & "|" & CLng(dIssue)
End Function

'--------------------------------------------------------------------
Private Sub MarkRow(rw As Range)
    With rw
        .Interior.Color = MARK_FILL
        .Font.Color = MARK_FONT
    End With
End Sub

'--------------------------------------------------------------------
' Hàm này dùng đ? ki?m tra tên c?t (ch?y riêng n?u c?n)
'--------------------------------------------------------------------
Sub ListColumnNames()
    Dim lo As ListObject
    Dim i As Long
    
    On Error Resume Next
    Set lo = ThisWorkbook.Sheets("Sheet1").ListObjects("Table1")
    If Not lo Is Nothing Then
        Debug.Print "--- Table1 Columns ---"
        For i = 1 To lo.ListColumns.Count
            Debug.Print i, lo.ListColumns(i).Name
        Next i
    Else
        Debug.Print "Table1 not found"
    End If
    
    Dim wbLoan As Workbook, wsLoan As Worksheet
    Set wbLoan = Workbooks("Loan.xlsm")
    If Not wbLoan Is Nothing Then
        Set wsLoan = wbLoan.Sheets("ST Loan")
        Set lo = wsLoan.ListObjects("Short_term_loans")
        If Not lo Is Nothing Then
            Debug.Print "--- Short_term_loans Columns ---"
            For i = 1 To lo.ListColumns.Count
                Debug.Print i, lo.ListColumns(i).Name
            Next i
        End If
    End If
End Sub

