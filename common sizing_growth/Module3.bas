Sub AutoCommonSize_SelectedRange()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '-----------------------------------------------------------
    ' 1. Validate selected range
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

    ' Read entire range into array
    Dim dataArray As Variant
    dataArray = rng.Value

    '-----------------------------------------------------------
    ' 2. Determine first data column (skip code column if present)
    '-----------------------------------------------------------
    Dim dataStartCol As Long
    Dim found As Boolean
    Dim col As Long

    ' Search for header with "/" (date) or "Qu?" (quarter)
    For col = 2 To UBound(dataArray, 2)
        Dim headerVal As String
        headerVal = dataArray(1, col)
        If Len(headerVal) > 0 Then
            If InStr(1, headerVal, "/", vbTextCompare) > 0 Or _
               InStr(1, headerVal, "Qu?", vbTextCompare) > 0 Then
                dataStartCol = col
                found = True
                Exit For
            End If
        End If
    Next col

    If Not found Then
        ' Check if second column is a code column
        Dim header2 As Variant
        header2 = dataArray(1, 2)
        Dim isCodeCol As Boolean
        isCodeCol = False
        If IsNumeric(header2) And Len(Trim(header2)) <= 3 Then
            isCodeCol = True
        ElseIf InStr(1, header2, "Ma", vbTextCompare) > 0 Or _
               InStr(1, header2, "Ma", vbTextCompare) > 0 Or _
               InStr(1, header2, "Code", vbTextCompare) > 0 Or _
               InStr(1, header2, "STT", vbTextCompare) > 0 Then
            isCodeCol = True
        End If

        If isCodeCol Then
            If UBound(dataArray, 2) >= 3 Then
                dataStartCol = 3
                found = True
            Else
                MsgBox "Not enough data columns after skipping code column.", vbExclamation
                GoTo CleanExit
            End If
        Else
            dataStartCol = 2
            found = True
        End If
    End If

    If Not found Or dataStartCol > UBound(dataArray, 2) Then
        MsgBox "Could not determine data columns.", vbExclamation
        GoTo CleanExit
    End If

    '-----------------------------------------------------------
    ' 3. Ask user which report type (BS or PL)
    '-----------------------------------------------------------
    Dim reportType As Long
    Dim msg As String
    msg = "Choose the report type for Common Size calculation:" & vbNewLine & _
          "1 - Income Statement (Revenue / Doanh thu)" & vbNewLine & _
          "2 - Balance Sheet (Total Assets / T?ng t�i s?n)"
    reportType = Application.InputBox(msg, "Select Report Type", 1, Type:=1)
    If reportType = 0 Then Exit Sub   ' cancelled

    '-----------------------------------------------------------
    ' 4. Find base row based on chosen type
    '-----------------------------------------------------------
    Dim dataStartRow As Integer
    dataStartRow = 2

    Dim baseRow As Long
    baseRow = 0

    Dim keywords As Variant
    Dim exclusions As Variant
    Dim baseType As String

    If reportType = 1 Then
        ' Income Statement (Revenue)
        keywords = Array("net revenue", "total revenue", "doanh thu thu?n", "revenue")
        exclusions = Array("deferred", "ch? phn b?", "cha thuc hien", "short-term", "long-term", _
                           "gross", "operating", "other", "gop", "hoat donng", "khc")
        baseType = "Revenue (Income Statement)"
    ElseIf reportType = 2 Then
        ' Balance Sheet (Assets)
        keywords = Array("total assets", "t?ng tai san", "assets")
        exclusions = Array("current", "non-current", "short-term", "long-term", _
                           "hi?n t?i", "d�i h?n", "other", "kh�c")
        baseType = "Total Assets (Balance Sheet)"
    Else
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, j As Long
    Dim cellText As String
    Dim normalizedText As String
    Dim lowerText As String
    Dim kw As Variant
    Dim excl As Variant
    Dim isExcluded As Boolean

    ' First, try exact match of priority keywords (no exclusion)
    For i = dataStartRow To UBound(dataArray, 1)
        cellText = dataArray(i, 1)
        If Len(cellText) > 0 Then
            normalizedText = StripAccents(LCase(Trim(cellText)))
            For Each kw In keywords
                If InStr(1, normalizedText, kw, vbTextCompare) > 0 Then
                    baseRow = i
                    GoTo foundBase
                End If
            Next kw
        End If
    Next i

    ' If not found, try broader (contains any keyword) with exclusion
    For i = dataStartRow To UBound(dataArray, 1)
        cellText = dataArray(i, 1)
        If Len(cellText) > 0 Then
            lowerText = LCase(Trim(cellText))
            normalizedText = StripAccents(lowerText)
            Dim containsKeyword As Boolean
            containsKeyword = False
            For Each kw In keywords
                If InStr(1, normalizedText, kw, vbTextCompare) > 0 Then
                    containsKeyword = True
                    Exit For
                End If
            Next kw
            If containsKeyword Then
                isExcluded = False
                For Each excl In exclusions
                    If InStr(1, lowerText, excl, vbTextCompare) > 0 Then
                        isExcluded = True
                        Exit For
                    End If
                Next excl
                If Not isExcluded Then
                    baseRow = i
                    GoTo foundBase
                End If
            End If
        End If
    Next i

foundBase:
    If baseRow = 0 Then
        Dim keywordManual As String
        keywordManual = InputBox("Could not find the base row automatically." & vbNewLine & _
                                 "Please enter a keyword (e.g., '" & keywords(0) & "'):", _
                                 "Manual Keyword", keywords(0))
        If keywordManual = "" Then Exit Sub
        For i = dataStartRow To UBound(dataArray, 1)
            cellText = dataArray(i, 1)
            If InStr(1, cellText, keywordManual, vbTextCompare) > 0 Then
                baseRow = i
                baseType = "Manual: " & keywordManual
                Exit For
            End If
        Next i
        If baseRow = 0 Then
            MsgBox "Keyword '" & keywordManual & "' not found.", vbExclamation
            Exit Sub
        End If
    End If

    ' Confirm with user
    Dim confirmMsg As String
    confirmMsg = "Base row detected: " & dataArray(baseRow, 1) & "  (" & baseType & ")" & vbNewLine & _
                 "Is this correct? Click Yes to proceed, No to enter manually."
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Base Row") = vbNo Then
        keywordManual = InputBox("Enter the correct keyword:", "Manual Keyword", keywords(0))
        If keywordManual = "" Then Exit Sub
        For i = dataStartRow To UBound(dataArray, 1)
            cellText = dataArray(i, 1)
            If InStr(1, cellText, keywordManual, vbTextCompare) > 0 Then
                baseRow = i
                baseType = "Manual: " & keywordManual
                Exit For
            End If
        Next i
        If baseRow = 0 Then
            MsgBox "Keyword not found.", vbExclamation
            Exit Sub
        End If
    End If

    '-----------------------------------------------------------
    ' 5. Calculate Common Size on array
    '-----------------------------------------------------------
    Dim outputArray As Variant
    outputArray = dataArray

    Dim baseVal As Double
    Dim cellVal As Variant

    For col = dataStartCol To UBound(dataArray, 2)
        baseVal = dataArray(baseRow, col)
        If Not IsNumeric(baseVal) Or baseVal = 0 Then GoTo NextCol
        For i = dataStartRow To UBound(dataArray, 1)
            cellVal = dataArray(i, col)
            If IsNumeric(cellVal) Then
                outputArray(i, col) = CDbl(cellVal) / baseVal
            Else
                ' keep original
                outputArray(i, col) = cellVal
            End If
        Next i
NextCol:
    Next col

    '-----------------------------------------------------------
    ' 6. Ask for destination and paste
    '-----------------------------------------------------------
    Dim destRange As Range
    On Error Resume Next
    Set destRange = Application.InputBox( _
        "Select the destination cell (top-left corner of the output range):", _
        "Choose Destination", Type:=8)
    On Error GoTo ErrorHandler
    If destRange Is Nothing Then Exit Sub

    Set destRange = destRange.Cells(1, 1)
    Dim rowsCount As Long, colsCount As Long
    rowsCount = UBound(outputArray, 1) - LBound(outputArray, 1) + 1
    colsCount = UBound(outputArray, 2) - LBound(outputArray, 2) + 1

    destRange.Resize(rowsCount, colsCount).Value = outputArray

    ' Apply percentage format to data columns (exclude header)
    For col = dataStartCol To UBound(dataArray, 2)
        Dim destCol As Long
        destCol = destRange.Column + col - 1
        Dim formatRange As Range
        Set formatRange = destRange.offset(1, col - 1).Resize(rowsCount - 1, 1)
        formatRange.NumberFormat = "0.00%"
    Next col

    MsgBox "Common Size calculation completed!" & vbNewLine & _
           "Base row: " & dataArray(baseRow, 1) & vbNewLine & _
           "Results pasted starting at " & destRange.Address, vbInformation
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description & vbNewLine & _
           "Occurred at row " & i & ", column " & col, vbCritical
    Resume CleanExit

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------
' Helper: Strip accents
'------------------------------------------------------------------
Function StripAccents(ByVal txt As String) As String
    Dim chars As String
    Dim i As Long
    Dim result As String
    result = txt
    chars = "��???�?????�?????" & _
            "��???�?????" & _
            "?�???" & _
            "?�???�?????�?????" & _
            "��???�?????" & _
            "?????" & _
            "�" & _
            "��???�?????�?????" & _
            "��???�?????" & _
            "?�???" & _
            "?�???�?????�?????" & _
            "��???�?????" & _
            "?????" & _
            "�"
    Dim noAccent As String
    noAccent = "aaaaaaaaaaaaaaaaa" & _
               "eeeeeeeeeee" & _
               "iiiii" & _
               "ooooooooooooooooooo" & _
               "uuuuuuuuuuu" & _
               "yyyyy" & _
               "d" & _
               "AAAAAAAAAAAAAAAAA" & _
               "EEEEEEEEEEE" & _
               "IIIII" & _
               "OOOOOOOOOOOOOOOOOOO" & _
               "UUUUUUUUUUU" & _
               "YYYYY" & _
               "D"
    For i = 1 To Len(chars)
        result = Replace(result, Mid(chars, i, 1), Mid(noAccent, i, 1))
    Next i
    StripAccents = result
End Function

