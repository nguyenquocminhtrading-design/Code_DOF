Sub CalculateGrowth_SelectedRange()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '-----------------------------------------------------------
    ' 1. Validate selected range & read data into array
    '-----------------------------------------------------------
    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "No range selected. Please select a data range.", vbExclamation
        GoTo CleanExit
    End If
    Set rng = rng.Areas(1)

    Dim ws As Worksheet
    Set ws = rng.Parent

    Dim firstRow As Long, firstCol As Long
    Dim lastRow As Long, lastCol As Long
    firstRow = rng.Row
    firstCol = rng.Column
    lastRow = rng.Row + rng.Rows.count - 1
    lastCol = rng.Column + rng.Columns.count - 1

    If rng.Rows.count < 2 Or rng.Columns.count < 3 Then
        MsgBox "Range must include header row, label column, and at least 2 data columns.", vbExclamation
        GoTo CleanExit
    End If

    ' Read data into array
    Dim dataArray As Variant
    dataArray = rng.Value

    '-----------------------------------------------------------
    ' 2. Determine first data column (skip VN/EN text, code, note, STT)
    ' Supports Format 1 (1 label col + FY cols) and Format 2 (VN, EN, Code, Note + Date cols)
    '-----------------------------------------------------------
    Dim dataStartCol As Long
    Dim found As Boolean
    Dim col As Long

    dataStartCol = 0
    found = False

    ' Priority 1: Check header for dates, quarters, years (e.g. 2023, 2024, Q1, Quy, /, FY, Jun, Jan, etc.)
    For col = 2 To UBound(dataArray, 2)
        Dim headerVal As String
        headerVal = Trim(CStr(dataArray(1, col)))
        If Len(headerVal) > 0 Then
            Dim hLower As String
            hLower = LCase(headerVal)
            Dim isDescCol As Boolean
            isDescCol = (InStr(1, hLower, "ma so", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "thuyet minh", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "chi tieu", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "line item", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "item code", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "stt", vbTextCompare) > 0 Or _
                         InStr(1, hLower, "note", vbTextCompare) > 0)

            If Not isDescCol Then
                If InStr(1, headerVal, "/", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "-", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Q", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Qu", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "FY", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Jun", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Dec", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Mar", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Sep", vbTextCompare) > 0 Or _
                   (IsNumeric(headerVal) And Len(headerVal) = 4 And Val(headerVal) >= 1990 And Val(headerVal) <= 2100) Then
                    dataStartCol = col
                    found = True
                    Exit For
                End If
            End If
        End If
    Next col

    ' Priority 2: Check by inspecting data contents (finding first column with majority numeric values)
    If Not found Then
        For col = 2 To UBound(dataArray, 2)
            Dim numCount As Long, totalCheck As Long, checkR As Long
            numCount = 0: totalCheck = 0
            For checkR = 2 To Application.Min(UBound(dataArray, 1), 15)
                Dim vTest As Variant
                vTest = dataArray(checkR, col)
                If Len(Trim(CStr(vTest))) > 0 Then
                    totalCheck = totalCheck + 1
                    If IsNumeric(vTest) Then
                        If col <= 4 And (vTest = CLng(Val(vTest))) And Val(vTest) >= 1 And Val(vTest) <= 99 Then
                            ' likely code col
                        Else
                            numCount = numCount + 1
                        End If
                    End If
                End If
            Next checkR
            If totalCheck > 0 And (numCount / totalCheck) >= 0.5 Then
                dataStartCol = col
                found = True
                Exit For
            End If
        Next col
    End If

    ' Priority 3: Fallback verification prompt if confidence is not 100% or multi-column layout detected
    If Not found Or dataStartCol > 2 Then
        Dim promptDefault As Long
        If found Then promptDefault = dataStartCol Else promptDefault = 2
        Dim colPromptMsg As String
        colPromptMsg = "Confirm Data Start Column:" & vbNewLine & vbNewLine & _
                       "Detected column " & promptDefault & " ('" & dataArray(1, promptDefault) & "') as first data column." & vbNewLine & _
                       "Columns 1 to " & (promptDefault - 1) & " will be treated as item descriptions/codes." & vbNewLine & vbNewLine & _
                       "Enter column index (e.g. 2, 3, 5):"
        Dim userColInput As String
        userColInput = InputBox(colPromptMsg, "Data Start Column Confirmation", CStr(promptDefault))
        If userColInput = "" Then GoTo CleanExit
        If IsNumeric(userColInput) Then
            dataStartCol = CLng(userColInput)
            If dataStartCol >= 2 And dataStartCol <= UBound(dataArray, 2) Then
                found = True
            Else
                MsgBox "Column index out of range.", vbExclamation
                GoTo CleanExit
            End If
        End If
    End If

    If Not found Or dataStartCol > UBound(dataArray, 2) Then
        MsgBox "Could not determine data columns.", vbExclamation
        GoTo CleanExit
    End If

    '-----------------------------------------------------------
    ' 3. Define data rows to process (skip rows with non-numeric in first data col)
    '-----------------------------------------------------------
    Dim dataStartRow As Long
    dataStartRow = 2   ' row 1 is header

    ' Create a list of valid data rows (numeric in first data column)
    Dim validRows As New Collection
    Dim i As Long
    For i = dataStartRow To UBound(dataArray, 1)
        Dim firstDataVal As Variant
        firstDataVal = dataArray(i, dataStartCol)
        If IsNumeric(firstDataVal) Then
            validRows.Add i
        End If
    Next i

    If validRows.count < 1 Then
        MsgBox "No numeric data rows found in the first data column (column " & dataStartCol & ")." & vbNewLine & vbNewLine & _
               "Please check:" & vbNewLine & _
               "  • Your selection includes at least one row with numeric values." & vbNewLine & _
               "  • The first data column is not a text/code column (e.g. 'VNDS001')." & vbNewLine & _
               "  • If column 2 is still a code column, start your selection from the data columns only.", vbExclamation
        GoTo CleanExit
    End If

    '-----------------------------------------------------------
    ' 4. Calculate Growth as Excel Dynamic Formulas (for valid rows only)
    '-----------------------------------------------------------
    Dim outputArray As Variant
    outputArray = dataArray   ' start with a copy

    Dim j As Long
    Dim validRow As Variant
    Dim DQ As String
    DQ = Chr(34)
    Dim safeWsName As String
    safeWsName = "'" & Replace(ws.Name, "'", "''") & "'!"
    Dim currRef As String, prevRef As String

    For Each validRow In validRows
        i = validRow
        ' First data column: no previous period to compare, leave empty
        outputArray(i, dataStartCol) = ""

        ' Subsequent data columns (starting from dataStartCol+1)
        For j = dataStartCol + 1 To UBound(dataArray, 2)
            ' BUG2 FIX: Use absolute addresses (True, True) so formulas don't drift when copied
            currRef = safeWsName & ws.Cells(firstRow + i - 1, firstCol + j - 1).Address(True, True)
            prevRef = safeWsName & ws.Cells(firstRow + i - 1, firstCol + j - 2).Address(True, True)

            outputArray(i, j) = "=IFERROR((" & currRef & "-" & prevRef & ")/" & prevRef & ", " & DQ & DQ & ")"
        Next j
    Next validRow

    '-----------------------------------------------------------
    ' 5. Ask user for destination range and paste result (as Formulas)
    '-----------------------------------------------------------
    Dim destRange As Range
    On Error Resume Next
    Set destRange = Application.InputBox( _
        "Select the top-left cell for the result:", _
        "Destination", Type:=8)
    On Error GoTo ErrorHandler
    If destRange Is Nothing Then Exit Sub
    Set destRange = destRange.Cells(1, 1)

    Dim rowsCount As Long, colsCount As Long
    rowsCount = UBound(outputArray, 1) - LBound(outputArray, 1) + 1
    colsCount = UBound(outputArray, 2) - LBound(outputArray, 2) + 1

    ' Paste formulas
    destRange.Resize(rowsCount, colsCount).Formula = outputArray

    ' Apply percentage format to data columns (exclude header and label column)
    For j = dataStartCol To UBound(dataArray, 2)
        Dim destCol As Long
        destCol = destRange.Column + j - 1
        Dim formatRange As Range
        Set formatRange = destRange.offset(1, j - 1).Resize(rowsCount - 1, 1)
        formatRange.NumberFormat = "0.00%"
    Next j

    MsgBox "Growth calculation completed!" & vbNewLine & _
           "Processed " & validRows.count & " data rows." & vbNewLine & _
           "Results pasted at " & destRange.Address, vbInformation
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

