Option Explicit

' ======================================================================
' INTERCOMPANY ELIMINATION MACRO - DDF GROUP (v2 - English source only)
' ----------------------------------------------------------------------
' WHY THIS VERSION CHANGED:
'   The previous version stored Vietnamese company names directly as
'   string literals in the VBA source code. VBA source files are saved
'   using the Windows system codepage (not UTF-8), so on many machines
'   Vietnamese diacritics silently get corrupted into "?" characters
'   (e.g. "C? PH?N" instead of "C? PHAN"). That made every related-party
'   comparison fail.
'
'   Fix applied here:
'   1. NO Vietnamese diacritic characters are stored anywhere in this
'      source file. Everything is plain English/ASCII text.
'   2. A helper function (NormalizeName) strips diacritics from the
'      customer name that Excel reads from the worksheet AT RUNTIME
'      (that value is safe/correct because it comes from the cell,
'      not from VBA source code), using numeric Unicode code points
'      (ChrW(&Hxxxx)) instead of literal accented letters. Numeric
'      code points cannot be corrupted by codepage/font issues.
'   3. The related-party list is now just plain ASCII (e.g.
'      "CONG TY CO PHAN THUC PHAM DAN ON"), so it is 100% safe to store
'      in source code and 100% safe to compare against the normalized
'      runtime name.
'   4. A second, lighter-weight "keyword" pass automatically flags
'      NEW/unlisted customer names that look like they might be a
'      related party (e.g. contain "DAN ON", "DAN D", "AN DIEN",
'      "DANWAY", "DOF"), even if they don't exactly match the list.
'      These are reported separately for manual review - they are NOT
'      auto-eliminated, so the elimination total stays precise while
'      you still get automatic detection of new naming variants.
'
' Functionality is otherwise the same as before:
'   1. Scans every sheet in the workbook and identifies sales-data
'      sheets for each legal entity based on SHEET NAME.
'   2. On each sheet, auto-detects the customer Name column, the
'      revenue column(s) (Net Sales/Amount - summed if both exist),
'      the cost column (Total Cost/COGS), or derives cost from
'      Revenue - GP$ if no cost column exists.
'   3. Matches customer names against the related-party list.
'   4. Outputs results to sheet "IC_Elimination", plus a new
'      "Review" section listing possible-but-unconfirmed matches.
'
' HOW TO UPDATE THE RELATED PARTY LIST:
'   Add new entries to GetRelatedParties() using plain English/ASCII
'   text only (upper or lower case, punctuation is fine - it gets
'   normalized automatically). Do not type accented Vietnamese
'   characters directly into this file.
' ======================================================================

Const OUTPUT_SHEET_NAME As String = "IC_Elimination"

' Keywords used to identify which sheets contain sales data
Private Function GetSheetKeywords() As Variant
    GetSheetKeywords = Array("DOF", "ADF", "DOAU", "CAN")
End Function

' ----------------------------------------------------------------------
' Related-party list - PLAIN ENGLISH/ASCII ONLY.
' These are compared against the NORMALIZED customer name (see
' NormalizeName), which uppercases, strips accents, and collapses all
' punctuation/spacing. So "Dan-D Foods Ltd.", "DAN D FOODS LTD",
' "Dan- D  Foods,Ltd" all normalize to the same string and will match
' the single entry "DAN D FOODS LTD" below.
' ----------------------------------------------------------------------
Private Function GetRelatedParties() As Variant
    GetRelatedParties = Array( _
        "CONG TY CO PHAN THUC PHAM DAN ON", _
        "CONG TY CP THUC PHAM DAN ON CHI NHANH MY PHUOC 3", _
        "CONG TY CO PHAN THUC PHAM DAN ON CHI NHANH MY PHUOC 3", _
        "CONG TY CO PHAN THUC PHAM DAN ON CN MY PHUOC 3", _
        "CONG TY CO PHAN THUC PHAM AN DIEN", _
        "QUY XA HOI TU THIEN DAN ON", _
        "DAN D PAK CO LTD", _
        "DAN D FOODS HONG KONG LIMITED", _
        "DAN D FOODS HONGKONG LTD", _
        "DAN D FOODS LTD", _
        "DAN D PAK USA", _
        "DANWAY HOLDINGS LIMITED", _
        "DOF VIETNAM", _
        "DOF MY PHUOC 3", _
        "DAN D FOODS LTD CANADA", _
        "DAN ON FOODS CORPORATION CAD", _
        "DAN D FOODS HK LTD US" _
    )
End Function

' ----------------------------------------------------------------------
' Keyword list used ONLY for the "possible match / needs review" pass.
' If a customer name is not an exact match above, but its normalized
' text contains one of these tokens, it gets listed for manual review
' instead of being silently missed. Keep these short and distinctive
' so they don't catch unrelated third-party customers.
' ----------------------------------------------------------------------
Private Function GetReviewKeywords() As Variant
    GetReviewKeywords = Array( _
        "DAN ON", _
        "DAN D FOODS", _
        "DAN D PAK", _
        "AN DIEN", _
        "DANWAY", _
        "DOF VIETNAM", _
        "DOF MY PHUOC" _
    )
End Function

' --- Header candidates (English only; Vietnamese header variants are
'     handled by NormalizeName when matching, so we don't need to list
'     accented header text here) ---
Private Function GetNameHeaderCandidates() As Variant
    GetNameHeaderCandidates = Array("NAME", "CUSTOMER NAME", "TEN KHACH HANG", "KHACH HANG", "KHACH")
End Function

Private Function GetNetSalesCandidates() As Variant
    GetNetSalesCandidates = Array("NET SALES", "DOANH THU THUAN", "DOANH THU")
End Function

Private Function GetAmountCandidates() As Variant
    GetAmountCandidates = Array("AMOUNT", "DOANH SO", "THANH TIEN", "SALES")
End Function

Private Function GetCogsHeaderCandidates() As Variant
    GetCogsHeaderCandidates = Array("TOTAL COST", "GIA VON", "COGS", "COST OF GOODS")
End Function

Private Function GetGPHeaderCandidates() As Variant
    GetGPHeaderCandidates = Array("GP$", "GM", "LOI NHUAN GOP", "GROSS PROFIT")
End Function

' ======================================================================
' MAIN
' ======================================================================
Sub EliminateIntercompany()

    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim relatedParties As Variant
    Dim reviewKeywords As Variant
    Dim sheetKeywords As Variant
    Dim outRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Running Intercompany Elimination..."

    relatedParties = GetRelatedParties()
    reviewKeywords = GetReviewKeywords()
    sheetKeywords = GetSheetKeywords()

    ' --- Delete old output sheet if it exists, create a fresh one ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    wsOut.Name = OUTPUT_SHEET_NAME

    ' --- Title ---
    With wsOut
        .Range("A1").Value = "INTERCOMPANY ELIMINATION SUMMARY"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Auto-generated by macro EliminateIntercompany - run at: " & Format(Now, "dd/mm/yyyy hh:nn")
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Size = 9
    End With

    outRow = 4
    Dim summaryHeaderRow As Long
    summaryHeaderRow = outRow
    wsOut.Cells(outRow, 1).Value = "1. SUMMARY BY SHEET (SELLING ENTITY) - BEFORE / ELIMINATED / AFTER"
    wsOut.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = "Sheet"
    wsOut.Cells(outRow, 2).Value = "Intercompany Rows"
    wsOut.Cells(outRow, 3).Value = "Total Revenue (Before)"
    wsOut.Cells(outRow, 4).Value = "Total Cost (Before)"
    wsOut.Cells(outRow, 5).Value = "Revenue Eliminated"
    wsOut.Cells(outRow, 6).Value = "Cost Eliminated"
    wsOut.Cells(outRow, 7).Value = "Revenue (After)"
    wsOut.Cells(outRow, 8).Value = "Cost (After)"
    wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 8)).Font.Bold = True
    outRow = outRow + 1
    Dim summaryFirstRow As Long
    summaryFirstRow = outRow

    Dim detailStartRow As Long
    Dim matchCount As Long
    matchCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> OUTPUT_SHEET_NAME Then
            If SheetMatchesKeyword(ws.Name, sheetKeywords) Then matchCount = matchCount + 1
        End If
    Next ws

    detailStartRow = summaryFirstRow + matchCount + 4
    Dim curDetailRow As Long
    curDetailRow = detailStartRow
    Dim curSummaryRow As Long
    curSummaryRow = summaryFirstRow

    ' Collect review (possible-match) rows from every sheet here first,
    ' we will write them out after the detail sections
    Dim reviewDict As Object
    Set reviewDict = CreateObject("Scripting.Dictionary")   ' key = sheet|name, item = Array(rev, cogs, count)

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> OUTPUT_SHEET_NAME Then
            If SheetMatchesKeyword(ws.Name, sheetKeywords) Then
                Dim revTotal As Double, cogsTotal As Double, rowCount As Long
                Dim grandRevTotal As Double, grandCogsTotal As Double
                ProcessSheet ws, relatedParties, reviewKeywords, wsOut, curDetailRow, _
                    revTotal, cogsTotal, rowCount, grandRevTotal, grandCogsTotal, reviewDict

                wsOut.Cells(curSummaryRow, 1).Value = ws.Name
                wsOut.Cells(curSummaryRow, 2).Value = rowCount
                wsOut.Cells(curSummaryRow, 3).Value = grandRevTotal
                wsOut.Cells(curSummaryRow, 4).Value = grandCogsTotal
                wsOut.Cells(curSummaryRow, 5).Value = revTotal
                wsOut.Cells(curSummaryRow, 6).Value = cogsTotal
                wsOut.Cells(curSummaryRow, 7).Value = grandRevTotal - revTotal
                wsOut.Cells(curSummaryRow, 8).Value = grandCogsTotal - cogsTotal
                wsOut.Range(wsOut.Cells(curSummaryRow, 3), wsOut.Cells(curSummaryRow, 8)).NumberFormat = "#,##0;(#,##0);""-"""
                curSummaryRow = curSummaryRow + 1
            End If
        End If
    Next ws

    ' --- GRAND TOTAL row ---
    wsOut.Cells(curSummaryRow, 1).Value = "GRAND TOTAL (*)"
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
    wsOut.Cells(curSummaryRow, 1).Value = "(*) Only valid to sum if all sheets use the same currency. If a sheet is in a foreign currency (e.g. DOAU, CAN), convert before summing."
    wsOut.Cells(curSummaryRow, 1).Font.Italic = True
    wsOut.Cells(curSummaryRow, 1).Font.Size = 9

    ' --- Formatting for summary block ---
    With wsOut
        .Range(.Cells(summaryHeaderRow, 1), .Cells(summaryHeaderRow, 8)).Interior.Color = RGB(252, 228, 214)
        .Range(.Cells(summaryFirstRow - 1, 1), .Cells(summaryFirstRow - 1, 8)).Interior.Color = RGB(31, 78, 120)
        .Range(.Cells(summaryFirstRow - 1, 1), .Cells(summaryFirstRow - 1, 8)).Font.Color = RGB(255, 255, 255)
    End With

    ' --- Write the "Review - possible related party (not auto-eliminated)" section
    '     after all sheet detail blocks ---
    Dim reviewRow As Long
    reviewRow = curDetailRow + 1
    wsOut.Cells(reviewRow, 1).Value = "2. REVIEW - POSSIBLE RELATED PARTIES NOT IN THE EXACT LIST (NOT ELIMINATED)"
    wsOut.Cells(reviewRow, 1).Font.Bold = True
    reviewRow = reviewRow + 1
    wsOut.Cells(reviewRow, 1).Value = "These names were NOT eliminated because they don't exactly match GetRelatedParties(), but they contain a keyword (e.g. DAN ON, DAN D FOODS, AN DIEN, DANWAY, DOF) that suggests they might be a related party. Review and, if confirmed, add the exact normalized name to GetRelatedParties() and re-run."
    wsOut.Cells(reviewRow, 1).Font.Italic = True
    wsOut.Cells(reviewRow, 1).Font.Size = 9
    reviewRow = reviewRow + 1

    wsOut.Cells(reviewRow, 1).Value = "Sheet"
    wsOut.Cells(reviewRow, 2).Value = "Customer Name (as in data)"
    wsOut.Cells(reviewRow, 3).Value = "Revenue"
    wsOut.Cells(reviewRow, 4).Value = "Cost"
    wsOut.Cells(reviewRow, 5).Value = "Row Count"
    wsOut.Range(wsOut.Cells(reviewRow, 1), wsOut.Cells(reviewRow, 5)).Font.Bold = True
    wsOut.Range(wsOut.Cells(reviewRow, 1), wsOut.Cells(reviewRow, 5)).Interior.Color = RGB(255, 242, 204)
    reviewRow = reviewRow + 1

    If reviewDict.Count = 0 Then
        wsOut.Cells(reviewRow, 1).Value = "(No possible-match names found outside the exact related-party list)"
        wsOut.Cells(reviewRow, 1).Font.Italic = True
        reviewRow = reviewRow + 1
    Else
        Dim rKey As Variant
        Dim rVal As Variant
        Dim sheetName As String, custName As String
        Dim sepPos As Long
        For Each rKey In reviewDict.Keys
            rVal = reviewDict(rKey)
            sepPos = InStr(1, CStr(rKey), "|")
            sheetName = Left(CStr(rKey), sepPos - 1)
            custName = Mid(CStr(rKey), sepPos + 1)
            wsOut.Cells(reviewRow, 1).Value = sheetName
            wsOut.Cells(reviewRow, 2).Value = custName
            wsOut.Cells(reviewRow, 3).Value = rVal(0)
            wsOut.Cells(reviewRow, 4).Value = rVal(1)
            wsOut.Cells(reviewRow, 5).Value = rVal(2)
            wsOut.Range(wsOut.Cells(reviewRow, 3), wsOut.Cells(reviewRow, 4)).NumberFormat = "#,##0;(#,##0);""-"""
            reviewRow = reviewRow + 1
        Next rKey
    End If

    wsOut.Columns("A:H").AutoFit
    wsOut.Activate
    wsOut.Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Intercompany Elimination finished." & vbCrLf & _
           "See sheet: " & OUTPUT_SHEET_NAME & vbCrLf & _
           "Check section 2 (Review) for possible related parties that were not auto-eliminated.", vbInformation
End Sub

' ======================================================================
' Process one sheet
' ======================================================================
Private Sub ProcessSheet(ws As Worksheet, relatedParties As Variant, reviewKeywords As Variant, _
        wsOut As Worksheet, ByRef curRow As Long, _
        ByRef revTotal As Double, ByRef cogsTotal As Double, ByRef rowCount As Long, _
        ByRef grandRevTotal As Double, ByRef grandCogsTotal As Double, _
        ByRef reviewDict As Object)

    revTotal = 0: cogsTotal = 0: rowCount = 0
    grandRevTotal = 0: grandCogsTotal = 0

    ' --- Find the Name column ---
    Dim headerCell As Range
    Set headerCell = FindHeaderCell(ws, GetNameHeaderCandidates())
    If headerCell Is Nothing Then
        wsOut.Cells(curRow, 1).Value = "[" & ws.Name & "] Name/Customer Name column not found - sheet skipped"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 2
        Exit Sub
    End If

    Dim headerRow As Long, nameCol As Long
    headerRow = headerCell.Row
    nameCol = headerCell.Column

    ' --- Find revenue column(s): Net Sales and Amount (added together if both exist) ---
    Dim revNetSalesCol As Long, revAmountCol As Long
    revNetSalesCol = FindColInRow(ws, headerRow, GetNetSalesCandidates())
    revAmountCol = FindColInRow(ws, headerRow, GetAmountCandidates())

    If revNetSalesCol = 0 And revAmountCol = 0 Then
        wsOut.Cells(curRow, 1).Value = "[" & ws.Name & "] Revenue column (Net Sales/Amount) not found - sheet skipped"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 2
        Exit Sub
    End If

    ' --- Find the cost column ---
    Dim cogsCol As Long, gpCol As Long
    cogsCol = FindColInRow(ws, headerRow, GetCogsHeaderCandidates())
    gpCol = 0
    If cogsCol = 0 Then gpCol = FindColInRow(ws, headerRow, GetGPHeaderCandidates())

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    Dim maxCol As Long
    maxCol = nameCol
    If revNetSalesCol > maxCol Then maxCol = revNetSalesCol
    If revAmountCol > maxCol Then maxCol = revAmountCol
    If cogsCol > maxCol Then maxCol = cogsCol
    If gpCol > maxCol Then maxCol = gpCol

    Dim dataArr As Variant
    dataArr = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, maxCol)).Value2

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, nm As String, nmNorm As String
    Dim rev As Double, cogs As Double

    For i = 1 To UBound(dataArr, 1)
        nm = Trim(CStr(dataArr(i, nameCol)))

        rev = 0
        If revNetSalesCol > 0 Then
            If IsNumeric(dataArr(i, revNetSalesCol)) Then rev = rev + CDbl(dataArr(i, revNetSalesCol))
        End If
        If revAmountCol > 0 Then
            If IsNumeric(dataArr(i, revAmountCol)) Then rev = rev + CDbl(dataArr(i, revAmountCol))
        End If

        cogs = 0
        If cogsCol > 0 Then
            If IsNumeric(dataArr(i, cogsCol)) Then cogs = CDbl(dataArr(i, cogsCol))
        ElseIf gpCol > 0 Then
            Dim gp As Double: gp = 0
            If IsNumeric(dataArr(i, gpCol)) Then gp = CDbl(dataArr(i, gpCol))
            cogs = rev - gp
        End If

        grandRevTotal = grandRevTotal + rev
        grandCogsTotal = grandCogsTotal + cogs

        If Len(nm) > 0 Then
            nmNorm = NormalizeName(nm)

            If IsRelatedParty(nmNorm, relatedParties) Then
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

            ElseIf ContainsReviewKeyword(nmNorm, reviewKeywords) Then
                ' Not an exact related-party match, but looks suspicious -
                ' log it for manual review instead of silently dropping it
                Dim rKey As String
                rKey = ws.Name & "|" & nm
                If reviewDict.Exists(rKey) Then
                    Dim rArr As Variant
                    rArr = reviewDict(rKey)
                    rArr(0) = rArr(0) + rev
                    rArr(1) = rArr(1) + cogs
                    rArr(2) = rArr(2) + 1
                    reviewDict(rKey) = rArr
                Else
                    reviewDict.Add rKey, Array(rev, cogs, 1)
                End If
            End If
        End If
    Next i

    ' --- Write detail table for this sheet ---
    wsOut.Cells(curRow, 1).Value = "Detail - sheet """ & ws.Name & """"
    wsOut.Cells(curRow, 1).Font.Bold = True
    wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 4)).Interior.Color = RGB(252, 228, 214)
    curRow = curRow + 1

    wsOut.Cells(curRow, 1).Value = "Related Party"
    wsOut.Cells(curRow, 2).Value = "Revenue Eliminated"
    wsOut.Cells(curRow, 3).Value = "Cost Eliminated"
    wsOut.Cells(curRow, 4).Value = "Row Count"
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
        wsOut.Cells(curRow, 1).Value = "(No intercompany transactions matched on this sheet)"
        wsOut.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 1
    End If

    curRow = curRow + 1
End Sub

' ======================================================================
' Utility functions
' ======================================================================
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

' Exact match against the normalized related-party list
Private Function IsRelatedParty(normalizedName As String, relatedParties As Variant) As Boolean
    Dim p As Variant
    For Each p In relatedParties
        If normalizedName = UCase(Trim(CStr(p))) Then
            IsRelatedParty = True
            Exit Function
        End If
    Next p
    IsRelatedParty = False
End Function

' "Looks like it might be related" check, for the review section only
Private Function ContainsReviewKeyword(normalizedName As String, reviewKeywords As Variant) As Boolean
    Dim k As Variant
    For Each k In reviewKeywords
        If InStr(1, normalizedName, UCase(Trim(CStr(k))), vbTextCompare) > 0 Then
            ContainsReviewKeyword = True
            Exit Function
        End If
    Next k
    ContainsReviewKeyword = False
End Function

Private Function FindHeaderCell(ws As Worksheet, candidates As Variant) As Range
    Dim r As Long, c As Long
    Dim cell As Range
    Dim cand As Variant
    For r = 1 To 10
        For c = 1 To 60
            Set cell = ws.Cells(r, c)
            If Len(Trim(CStr(cell.Value))) > 0 Then
                For Each cand In candidates
                    If NormalizeName(CStr(cell.Value)) = UCase(cand) Then
                        Set FindHeaderCell = cell
                        Exit Function
                    End If
                Next cand
            End If
        Next c
    Next r
    Set FindHeaderCell = Nothing
End Function

Private Function FindColInRow(ws As Worksheet, headerRow As Long, candidates As Variant) As Long
    Dim cand As Variant
    Dim c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 60 Then lastCol = 60
    For Each cand In candidates
        For c = 1 To lastCol
            If NormalizeName(CStr(ws.Cells(headerRow, c).Value)) = UCase(cand) Then
                FindColInRow = c
                Exit Function
            End If
        Next c
    Next cand
    FindColInRow = 0
End Function

' ======================================================================
' NormalizeName: uppercase + strip Vietnamese diacritics + collapse
' punctuation/spacing, so that names like "Dan-D Foods, Ltd." and
' "DAN D FOODS LTD" compare as equal. Diacritics are stripped using
' numeric Unicode code points (ChrW), never literal accented letters,
' so this function is 100% safe from source-file codepage corruption.
' ======================================================================
Private Function NormalizeName(ByVal s As String) As String
    Dim vnCodes As Variant
    Dim asciiChars As String
    Dim i As Long, j As Long
    Dim ch As String
    Dim code As Long
    Dim found As Boolean
    Dim result As String

    ' Vietnamese accented character Unicode code points (lower+upper),
    ' in the same order as asciiChars below
    vnCodes = Array( _
        &HE0, &HE1, &H1EA3, &HE3, &H1EA1, &H103, &H1EB1, &H1EAF, &H1EB3, &H1EB5, &H1EB7, _
        &HE2, &H1EA7, &H1EA5, &H1EA9, &H1EAB, &H1EAD, _
        &HE8, &HE9, &H1EBB, &H1EBD, &H1EB9, &HEA, &H1EC1, &H1EBF, &H1EC3, &H1EC5, &H1EC7, _
        &HEC, &HED, &H1EC9, &H129, &H1ECB, _
        &HF2, &HF3, &H1ECF, &HF5, &H1ECD, &HF4, &H1ED3, &H1ED1, &H1ED5, &H1ED7, &H1ED9, _
        &H1A1, &H1EDD, &H1EDB, &H1EDF, &H1EE1, &H1EE3, _
        &HF9, &HFA, &H1EE7, &H169, &H1EE5, &H1B0, &H1EEB, &H1EE9, &H1EED, &H1EEF, &H1EF1, _
        &H1EF3, &HFD, &H1EF7, &H1EF9, &H1EF5, &H111, _
        &HC0, &HC1, &H1EA2, &HC3, &H1EA0, &H102, &H1EB0, &H1EAE, &H1EB2, &H1EB4, &H1EB6, _
        &HC2, &H1EA6, &H1EA4, &H1EA8, &H1EAA, &H1EAC, _
        &HC8, &HC9, &H1EBA, &H1EBC, &H1EB8, &HCA, &H1EC0, &H1EBE, &H1EC2, &H1EC4, &H1EC6, _
        &HCC, &HCD, &H1EC8, &H128, &H1ECA, _
        &HD2, &HD3, &H1ECE, &HD5, &H1ECC, &HD4, &H1ED2, &H1ED0, &H1ED4, &H1ED6, &H1ED8, _
        &H1A0, &H1EDC, &H1EDA, &H1EDE, &H1EE0, &H1EE2, _
        &HD9, &HDA, &H1EE6, &H168, &H1EE4, &H1AF, &H1EEA, &H1EE8, &H1EEC, &H1EEE, &H1EF0, _
        &H1EF2, &HDD, &H1EF6, &H1EF8, &H1EF4, &H110 _
    )
    asciiChars = "AAAAAAAAAAAAAAAAAEEEEEEEEEEEIIIIIOOOOOOOOOOOOOOOOOOOOOOOUUUUUUUUUUUUYYYYYD" & _
                 "AAAAAAAAAAAAAAAAAEEEEEEEEEEEIIIIIOOOOOOOOOOOOOOOOOOOOOOOUUUUUUUUUUUUYYYYYD"
    ' Note: the mapping is case-insensitive by design because we uppercase
    ' the whole string before stripping, so a single ASCII target letter
    ' per code point is sufficient (both the lower and upper Vietnamese
    ' variant of a letter map to the same uppercase ASCII letter).

    s = UCase(s)
    result = ""

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536   ' normalize signed AscW on some locales

        ' --- Combining diacritical marks (Unicode block U+0300-U+036F) ---
        ' Some source systems save Vietnamese text as "to hop" / decomposed
        ' Unicode: a plain base letter (e.g. "O") immediately followed by a
        ' separate combining accent mark, instead of a single precomposed
        ' character. Treating that mark like any other unknown symbol would
        ' turn it into a space and split the word in two (e.g. "MY" becomes
        ' "M Y"), which breaks matching. So: DROP combining marks entirely
        ' (append nothing) instead of turning them into a space. The base
        ' letter right before it is already kept as-is below.
        If code >= &H300 And code <= &H36F Then
            ' skip this character completely - do not add to result
        Else
            found = False
            For j = LBound(vnCodes) To UBound(vnCodes)
                If code = vnCodes(j) Then
                    result = result & Mid(asciiChars, j + 1, 1)
                    found = True
                    Exit For
                End If
            Next j

            If Not found Then
                If (code >= 65 And code <= 90) Or (code >= 48 And code <= 57) Then
                    result = result & ch          ' keep A-Z and 0-9 as is
                Else
                    result = result & " "          ' collapse any other punctuation/space to a space
                End If
            End If
        End If
    Next i

    ' Collapse multiple spaces into one and trim
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    NormalizeName = Trim(result)
End Function
