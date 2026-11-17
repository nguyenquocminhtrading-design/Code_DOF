Attribute VB_Name = "Module2"
Option Explicit
'=====================================================================
' MODULE: mod_BuildStandardFS  (updated per 001-DOF chuan - accurate columns/rows)
' MUC DICH: Tao sheet BCTC chuan hoa cho tung entity tu sheet du lieu
'           (DOF, ADF, CAN, DOAU, DW, TW, HK, DDFH)
'=====================================================================

'---------------------- 1. CAU TRUC DU LIEU ---------------------------
Public Type EntityConfig
    EntityCode      As String
    DataSheetName   As String
    CodeCol         As String
    FirstDataCol    As String
    numMonths       As Long
    BSFirstRow      As Long
    BSLastRow       As Long
    PLFirstRow      As Long
    PLLastRow       As Long
    FXRow           As String
End Type

Private Type TplRow
    RowNum  As Long
    Label   As String
    Code    As String
    Kind    As String   ' D, R, E, B
    Args    As String
    section As String   ' BS hoac PL
End Type

'---------------------- 2. TEMPLATE CHUAN (LABELS IN ENGLISH) -----------
Private Function GetTemplateLines() As String
    Dim s As String
    s = ""
    s = s & "3|ASSETS|280|E|{4}+{36}|BS" & vbCrLf
    s = s & "4|CURRENT ASSETS|100|E|{5}+{8}+{15}+{23}+{26}+{30}|BS" & vbCrLf
    s = s & "5|Cash and cash equivalents|110|E|{6}+{7}|BS" & vbCrLf
    s = s & "6|Cash|111|D||BS" & vbCrLf
    s = s & "7|Cash equivalents|112|D||BS" & vbCrLf
    s = s & "8|Short-term financial investments|120|R|9:14|BS" & vbCrLf
    s = s & "9|Trading securities|121|D||BS" & vbCrLf
    s = s & "10|Provision for devaluation of trading securities|122|D||BS" & vbCrLf
    s = s & "11|Held-to-maturity investments (short-term)|123|D||BS" & vbCrLf
    s = s & "12|Provision for short-term held-to-maturity investments|124|D||BS" & vbCrLf
    s = s & "13|Other short-term investments|125|D||BS" & vbCrLf
    s = s & "14|Provision for losses on other short-term investments|126|D||BS" & vbCrLf
    s = s & "15|Short-term receivables|130|R|16:22|BS" & vbCrLf
    s = s & "16|Short-term trade receivables|131|D||BS" & vbCrLf
    s = s & "17|Short-term prepayments to suppliers|132|D||BS" & vbCrLf
    s = s & "18|Short-term inter-company receivables|133|D||BS" & vbCrLf
    s = s & "19|Receivables from construction contracts|134|D||BS" & vbCrLf
    s = s & "20|Other short-term receivables|135|D||BS" & vbCrLf
    s = s & "21|Allowance for doubtful short-term receivables|136|D||BS" & vbCrLf
    s = s & "22|Assets shortage awaiting treatment|137|D||BS" & vbCrLf
    s = s & "23|Inventories|140|R|24:25|BS" & vbCrLf
    s = s & "24|Inventories|141|D||BS" & vbCrLf
    s = s & "25|Allowance for inventory write-down|142|D||BS" & vbCrLf
    s = s & "26|Short-term biological assets|150|R|27:29|BS" & vbCrLf
    s = s & "27|Short-term livestock for one-time harvest|151|D||BS" & vbCrLf
    s = s & "28|Short-term seasonal crops|152|D||BS" & vbCrLf
    s = s & "29|Allowance for short-term biological assets|153|D||BS" & vbCrLf
    s = s & "30|Other current assets|160|R|31:35|BS" & vbCrLf
    s = s & "31|Short-term prepaid expenses|161|D||BS" & vbCrLf
    s = s & "32|Deductible VAT|162|D||BS" & vbCrLf
    s = s & "33|Taxes and other receivables from State|163|D||BS" & vbCrLf
    s = s & "34|Trading government bonds|164|D||BS" & vbCrLf
    s = s & "35|Other current assets (other)|165|D||BS" & vbCrLf
    s = s & "36|NON-CURRENT ASSETS|200|E|{37}+{44}+{54}+{63}+{66}+{69}+{76}|BS" & vbCrLf
    s = s & "37|Long-term receivables|210|R|38:43|BS" & vbCrLf
    s = s & "38|Long-term trade receivables|211|D||BS" & vbCrLf
    s = s & "39|Long-term prepayments to suppliers|212|D||BS" & vbCrLf
    s = s & "40|Working capital in affiliates|213|D||BS" & vbCrLf
    s = s & "41|Long-term inter-company receivables|214|D||BS" & vbCrLf
    s = s & "42|Other long-term receivables|215|D||BS" & vbCrLf
    s = s & "43|Allowance for doubtful long-term receivables|216|D||BS" & vbCrLf
    s = s & "44|Fixed assets|220|E|{45}+{51}+{48}|BS" & vbCrLf
    s = s & "45|Tangible fixed assets|221|E|{46}+{47}|BS" & vbCrLf
    s = s & "46|Historical cost (tangible)|222|D||BS" & vbCrLf
    s = s & "47|Accumulated depreciation (tangible)|223|D||BS" & vbCrLf
    s = s & "48|Finance leased assets|224|E|{49}+{50}|BS" & vbCrLf
    s = s & "49|Historical cost (finance lease)|225|D||BS" & vbCrLf
    s = s & "50|Accumulated depreciation (finance lease)|226|D||BS" & vbCrLf
    s = s & "51|Intangible fixed assets|227|E|{52}+{53}|BS" & vbCrLf
    s = s & "52|Initial cost (intangible)|228|D||BS" & vbCrLf
    s = s & "53|Accumulated amortization (intangible)|229|D||BS" & vbCrLf
    s = s & "54|Long-term biological assets|230|R|55:62|BS" & vbCrLf
    s = s & "55|Livestock for periodic production|231|D||BS" & vbCrLf
    s = s & "56|Immature livestock for periodic production|232|D||BS" & vbCrLf
    s = s & "57|Mature livestock for periodic production|233|D||BS" & vbCrLf
    s = s & "58|Historical cost (biological)|234|D||BS" & vbCrLf
    s = s & "59|Accumulated depreciation (biological)|235|D||BS" & vbCrLf
    s = s & "60|Long-term livestock for one-time harvest|236|D||BS" & vbCrLf
    s = s & "61|Long-term seasonal crops|237|D||BS" & vbCrLf
    s = s & "62|Allowance for long-term biological assets|238|D||BS" & vbCrLf
    s = s & "63|Investment property|240|R|64:65|BS" & vbCrLf
    s = s & "64|Historical cost (investment property)|241|D||BS" & vbCrLf
    s = s & "65|Accumulated depreciation (investment property)|242|D||BS" & vbCrLf
    s = s & "66|Long-term assets in progress|250|R|67:68|BS" & vbCrLf
    s = s & "67|Long-term work in process|251|D||BS" & vbCrLf
    s = s & "68|Construction in progress|252|D||BS" & vbCrLf
    s = s & "69|Long-term financial investments|260|R|70:75|BS" & vbCrLf
    s = s & "70|Investments in subsidiaries|261|D||BS" & vbCrLf
    s = s & "71|Investments in joint ventures and associates|262|D||BS" & vbCrLf
    s = s & "72|Investments in other entities|263|D||BS" & vbCrLf
    s = s & "73|Provision for losses on long-term investments|264|D||BS" & vbCrLf
    s = s & "74|Held-to-maturity investments (long-term)|265|D||BS" & vbCrLf
    s = s & "75|Provision for long-term held-to-maturity investments|266|D||BS" & vbCrLf
    s = s & "76|Other non-current assets|270|R|77:80|BS" & vbCrLf
    s = s & "77|Long-term prepaid expenses|271|D||BS" & vbCrLf
    s = s & "78|Deferred income tax assets|272|D||BS" & vbCrLf
    s = s & "79|Long-term components and spare parts|273|D||BS" & vbCrLf
    s = s & "80|Other non-current assets (other)|274|D||BS" & vbCrLf
    s = s & "81|LIABILITIES AND OWNER'S EQUITY|440|E|{82}+{114}|BS" & vbCrLf
    s = s & "82|LIABILITIES|300|E|{83}+{99}|BS" & vbCrLf
    s = s & "83|Current liabilities|310|R|84:98|BS" & vbCrLf
    s = s & "84|Short-term trade payables|311|D||BS" & vbCrLf
    s = s & "85|Short-term advances from customers|312|D||BS" & vbCrLf
    s = s & "86|Dividends payable|313|D||BS" & vbCrLf
    s = s & "87|Taxes and other payables to State|314|D||BS" & vbCrLf
    s = s & "88|Payables to employees|315|D||BS" & vbCrLf
    s = s & "89|Short-term accrued expenses|316|D||BS" & vbCrLf
    s = s & "90|Short-term inter-company payables|317|D||BS" & vbCrLf
    s = s & "91|Payables from construction contracts|318|D||BS" & vbCrLf
    s = s & "92|Short-term deferred revenue|319|D||BS" & vbCrLf
    s = s & "93|Other short-term payables|320|D||BS" & vbCrLf
    s = s & "94|Short-term borrowings and finance leases|321|D||BS" & vbCrLf
    s = s & "95|Short-term provisions|322|D||BS" & vbCrLf
    s = s & "96|Bonus and welfare funds|323|D||BS" & vbCrLf
    s = s & "97|Price stabilization fund|324|D||BS" & vbCrLf
    s = s & "98|Trading government bonds (short-term)|325|D||BS" & vbCrLf
    s = s & "99|Long-term liabilities|330|R|100:113|BS" & vbCrLf
    s = s & "100|Long-term trade payables|331|D||BS" & vbCrLf
    s = s & "101|Long-term advances from customers|332|D||BS" & vbCrLf
    s = s & "102|Long-term taxes and other payables to State|333|D||BS" & vbCrLf
    s = s & "103|Long-term accrued expenses|334|D||BS" & vbCrLf
    s = s & "104|Inter-company payables for working capital|335|D||BS" & vbCrLf
    s = s & "105|Long-term inter-company payables|336|D||BS" & vbCrLf
    s = s & "106|Long-term deferred revenue|337|D||BS" & vbCrLf
    s = s & "107|Other long-term payables|338|D||BS" & vbCrLf
    s = s & "108|Long-term borrowings and finance leases|339|D||BS" & vbCrLf
    s = s & "109|Convertible bonds|340|D||BS" & vbCrLf
    s = s & "110|Preferred shares|341|D||BS" & vbCrLf
    s = s & "111|Deferred income tax liabilities|342|D||BS" & vbCrLf
    s = s & "112|Long-term provisions|343|D||BS" & vbCrLf
    s = s & "113|Science and technology development fund|344|D||BS" & vbCrLf
    s = s & "114|OWNER'S EQUITY|400|E|{115}|BS" & vbCrLf
    s = s & "115|Owner's equity|410|E|{116}+{127}+SUM({119}:{126})|BS" & vbCrLf
    s = s & "116|Share capital|411|R|117:118|BS" & vbCrLf
    s = s & "117|Ordinary shares|411a|D||BS" & vbCrLf
    s = s & "118|Preferred shares (owner)|411b|D||BS" & vbCrLf
    s = s & "119|Share premium|412|D||BS" & vbCrLf
    s = s & "120|Bond conversion options|413|D||BS" & vbCrLf
    s = s & "121|Other owners' equity|414|D||BS" & vbCrLf
    s = s & "122|Treasury shares|415|D||BS" & vbCrLf
    s = s & "123|Revaluation surplus|416|D||BS" & vbCrLf
    s = s & "124|Foreign exchange differences|417|D||BS" & vbCrLf
    s = s & "125|Investment and development fund|418|D||BS" & vbCrLf
    s = s & "126|Other funds|419|D||BS" & vbCrLf
    s = s & "127|Retained earnings|420|R|128:129|BS" & vbCrLf
    s = s & "128|Retained earnings brought forward|420a|D||BS" & vbCrLf
    s = s & "129|Retained earnings current period|420b|D||BS" & vbCrLf
    s = s & "130|Check (should be 0)||E|{81}-{3}|BS" & vbCrLf
    ' ------------------ PROFIT AND LOSS ------------------
    s = s & "133|Revenue|1|D||PL" & vbCrLf
    s = s & "134|Sales deductions|2|D||PL" & vbCrLf
    s = s & "135|Net revenue||E|{133}-{134}|PL" & vbCrLf
    s = s & "136|Cost of goods sold|11|D||PL" & vbCrLf
    s = s & "137|Gross profit||E|{135}-{136}|PL" & vbCrLf
    s = s & "138|Finance income|22|D||PL" & vbCrLf
    s = s & "139|Finance costs|23|D||PL" & vbCrLf
    s = s & "140|In which: interest expenses|24|D||PL" & vbCrLf
    s = s & "141|Selling expenses|25|D||PL" & vbCrLf
    s = s & "142|General and administrative expenses|26|D||PL" & vbCrLf
    s = s & "143|Operating profit||E|{137}+{138}-{139}-{141}-{142}|PL" & vbCrLf
    s = s & "144|Other income|31|D||PL" & vbCrLf
    s = s & "145|Other expenses|32|D||PL" & vbCrLf
    s = s & "146|Other profit/loss||E|{144}-{145}|PL" & vbCrLf
    s = s & "147|Profit before tax||E|{143}+{146}|PL" & vbCrLf
    s = s & "148|Current income tax expense|51|D||PL" & vbCrLf
    s = s & "149|Deferred income tax expense|52|D||PL" & vbCrLf
    s = s & "150|Profit after tax||E|{147}-{148}|PL" & vbCrLf
    GetTemplateLines = s
End Function

Private Function ParseTemplate() As TplRow()
    Dim lines() As String, parts() As String
    Dim i As Long, n As Long
    Dim result() As TplRow

    lines = Split(GetTemplateLines(), vbCrLf)
    ReDim result(0 To UBound(lines))
    n = -1
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            n = n + 1
            result(n).RowNum = CLng(parts(0))
            result(n).Label = parts(1)
            result(n).Code = parts(2)
            result(n).Kind = parts(3)
            result(n).Args = parts(4)
            result(n).section = parts(5)
        End If
    Next i
    ReDim Preserve result(0 To n)
    ParseTemplate = result
End Function

'---------------------- 3. CONFIG TUNG ENTITY --------------------------
Public Function GetEntityConfigs() As EntityConfig()
    Dim cfgs(0 To 7) As EntityConfig

    ' ---- 001-DOF ----
    cfgs(0).EntityCode = "001-DOF"
    cfgs(0).DataSheetName = "DOF"
    cfgs(0).CodeCol = "D"
    cfgs(0).FirstDataCol = "F"
    cfgs(0).numMonths = 12
    cfgs(0).BSFirstRow = 4
    cfgs(0).BSLastRow = 133
    cfgs(0).PLFirstRow = 140
    cfgs(0).PLLastRow = 157
    cfgs(0).FXRow = ""

    ' ---- 002-ADF ----
    cfgs(1).EntityCode = "002-ADF"
    cfgs(1).DataSheetName = "ADF"
    cfgs(1).CodeCol = "D"
    cfgs(1).FirstDataCol = "E"
    cfgs(1).numMonths = 12
    cfgs(1).BSFirstRow = 6
    cfgs(1).BSLastRow = 134
    cfgs(1).PLFirstRow = 140
    cfgs(1).PLLastRow = 157
    cfgs(1).FXRow = ""

    ' ---- 003-CAN ----
    cfgs(2).EntityCode = "003-CAN"
    cfgs(2).DataSheetName = "CAN"
    cfgs(2).CodeCol = "O"
    cfgs(2).FirstDataCol = "C"
    cfgs(2).numMonths = 12
    cfgs(2).BSFirstRow = 7
    cfgs(2).BSLastRow = 54
    cfgs(2).PLFirstRow = 61
    cfgs(2).PLLastRow = 138
    cfgs(2).FXRow = "3"

    ' ---- 004-DOAU ----
    cfgs(3).EntityCode = "004-DOAU"
    cfgs(3).DataSheetName = "DOAU"
    cfgs(3).CodeCol = "S"
    cfgs(3).FirstDataCol = "G"
    cfgs(3).numMonths = 12
    cfgs(3).BSFirstRow = 6
    cfgs(3).BSLastRow = 87
    cfgs(3).PLFirstRow = 106
    cfgs(3).PLLastRow = 236
    cfgs(3).FXRow = "1"

    ' ---- 005-DW ----
    cfgs(4).EntityCode = "005-DW"
    cfgs(4).DataSheetName = "DW"
    cfgs(4).CodeCol = "O"
    cfgs(4).FirstDataCol = "B"
    cfgs(4).numMonths = 12
    cfgs(4).BSFirstRow = 4
    cfgs(4).BSLastRow = 34
    cfgs(4).PLFirstRow = 38
    cfgs(4).PLLastRow = 67
    cfgs(4).FXRow = "1"

    ' ---- 006-TW ----
    cfgs(5).EntityCode = "006-TW"
    cfgs(5).DataSheetName = "TW"
    cfgs(5).CodeCol = "O"
    cfgs(5).FirstDataCol = "B"
    cfgs(5).numMonths = 12
    cfgs(5).BSFirstRow = 4
    cfgs(5).BSLastRow = 101
    cfgs(5).PLFirstRow = 108
    cfgs(5).PLLastRow = 192
    cfgs(5).FXRow = ""

    ' ---- 007-HK ----
    cfgs(6).EntityCode = "007-HK"
    cfgs(6).DataSheetName = "HK"
    cfgs(6).CodeCol = "O"
    cfgs(6).FirstDataCol = "B"
    cfgs(6).numMonths = 12
    cfgs(6).BSFirstRow = 7
    cfgs(6).BSLastRow = 75
    cfgs(6).PLFirstRow = 81
    cfgs(6).PLLastRow = 176
    cfgs(6).FXRow = "2"

    ' ---- 008-DDFH ----
    cfgs(7).EntityCode = "008-DDFH"
    cfgs(7).DataSheetName = "DDFH"
    cfgs(7).CodeCol = "P"
    cfgs(7).FirstDataCol = "C"
    cfgs(7).numMonths = 12
    cfgs(7).BSFirstRow = 8
    cfgs(7).BSLastRow = 49
    cfgs(7).PLFirstRow = 54
    cfgs(7).PLLastRow = 107
    cfgs(7).FXRow = "3"

    GetEntityConfigs = cfgs
End Function

'---------------------- 4. HAM TIM DONG HEADER -------------------------
Private Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5

    Dim keywords As Variant
    keywords = Array("code", "ma so", "ma so", "items", "chi tieu", "description", "stt", "no.")

    For r = 1 To 5
        For c = 1 To lastCol
            Dim cellVal As String
            If IsError(ws.Cells(r, c).Value) Then
                cellVal = ""
            Else
                cellVal = LCase(Trim(CStr(ws.Cells(r, c).Value)))
            End If
            If cellVal <> "" Then
                Dim kw As Variant
                For Each kw In keywords
                    If InStr(cellVal, kw) > 0 Then
                        FindHeaderRow = r
                        Exit Function
                    End If
                Next kw
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

'---------------------- 5. HAM TIEN ICH -------------------------------
Private Function ColOffset(ByVal startCol As String, ByVal offset As Long) As String
    Dim idx As Long
    idx = Range(startCol & "1").Column + offset
    ColOffset = Split(Cells(1, idx).Address(True, False), "$")(0)
End Function

Private Function ExpandExpr(ByVal expr As String, ByVal colLetter As String) As String
    Dim result As String, i As Long, buf As String, ch As String
    Dim inToken As Boolean
    result = ""
    For i = 1 To Len(expr)
        ch = Mid(expr, i, 1)
        If ch = "{" Then
            inToken = True
            buf = ""
        ElseIf ch = "}" Then
            inToken = False
            result = result & colLetter & buf
        ElseIf inToken Then
            buf = buf & ch
        Else
            result = result & ch
        End If
    Next i
    ExpandExpr = result
End Function

Private Function BuildFormula(ByRef t As TplRow, ByRef cfg As EntityConfig, ByVal colLetter As String, ByVal monthIdx As Long) As String
    Dim f As String
    Dim dataCol As String
    Dim dataRangeFrom As String, dataRangeTo As String
    Dim codeRangeFrom As String, codeRangeTo As String
    Dim rr() As String, ll() As String, k As Long, parts As String
    Dim dataRangeStr As String, codeRangeStr As String

    Select Case t.Kind
        Case "B"
            BuildFormula = ""
        Case "R"
            rr = Split(t.Args, ":")
            BuildFormula = "=IFERROR(SUM(" & colLetter & rr(0) & ":" & colLetter & rr(1) & "),"""")"
        Case "L"
            ll = Split(t.Args, ",")
            parts = ""
            For k = LBound(ll) To UBound(ll)
                If parts <> "" Then parts = parts & ","
                parts = parts & colLetter & Trim(ll(k))
            Next k
            BuildFormula = "=IFERROR(SUM(" & parts & "),"""")"
        Case "E"
            BuildFormula = "=IFERROR(" & ExpandExpr(t.Args, colLetter) & ","""")"
        Case "D"
            dataCol = ColOffset(cfg.FirstDataCol, monthIdx - 1)

            If t.section = "PL" And cfg.PLFirstRow > 0 Then
                dataRangeFrom = dataCol & cfg.PLFirstRow
                dataRangeTo = dataCol & cfg.PLLastRow
                codeRangeFrom = cfg.CodeCol & cfg.PLFirstRow
                codeRangeTo = cfg.CodeCol & cfg.PLLastRow
            ElseIf t.section = "BS" And cfg.BSFirstRow > 0 Then
                dataRangeFrom = dataCol & cfg.BSFirstRow
                dataRangeTo = dataCol & cfg.BSLastRow
                codeRangeFrom = cfg.CodeCol & cfg.BSFirstRow
                codeRangeTo = cfg.CodeCol & cfg.BSLastRow
            Else
                dataRangeFrom = dataCol & ":" & dataCol
                dataRangeTo = ""
                codeRangeFrom = cfg.CodeCol & ":" & cfg.CodeCol
                codeRangeTo = ""
            End If

            If dataRangeTo = "" Then
                dataRangeStr = "'" & cfg.DataSheetName & "'!" & dataRangeFrom
                codeRangeStr = "'" & cfg.DataSheetName & "'!$" & codeRangeFrom
            Else
                dataRangeStr = "'" & cfg.DataSheetName & "'!" & dataRangeFrom & ":" & dataRangeTo
                codeRangeStr = "'" & cfg.DataSheetName & "'!$" & codeRangeFrom & ":$" & codeRangeTo
            End If

            f = "SUMIFS(" & dataRangeStr & "," & codeRangeStr & ",$B" & t.RowNum & ")"

            If cfg.FXRow <> "" Then
                f = f & "*'" & cfg.DataSheetName & "'!" & dataCol & "$" & cfg.FXRow
            End If

            BuildFormula = "=IFERROR(" & f & ","""")"
    End Select
End Function

'---------------------- 6. TAO SHEET CHO 1 ENTITY ----------------------
Public Sub BuildEntitySheet(ByRef cfg As EntityConfig)
    Dim ws As Worksheet
    Dim tpl() As TplRow
    Dim i As Long, m As Long, colIdx As Long, colLetter As String
    Dim frm As String
    Const ACC_FORMAT As String = "#,##0;(#,##0);""-"""

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(cfg.EntityCode).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = cfg.EntityCode

    ' ----- Header BS -----
    ws.Range("A1").Value = "BALANCE SHEET"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Chi tieu"
    ws.Range("B2").Value = "Ma so"
    ws.Range("A2:B2").Font.Bold = True

    ' ----- Header PL -----
    ws.Range("A131").Value = "PROFIT AND LOSS"
    ws.Range("A131").Font.Bold = True
    ws.Range("A132").Value = "Chi tieu"
    ws.Range("B132").Value = "Ma so"
    ws.Range("A132:B132").Font.Bold = True

    ' ----- Lay ngay thang tu sheet nguon (header row) -----
    Dim srcWs As Worksheet
    Set srcWs = ThisWorkbook.Sheets(cfg.DataSheetName)
    Dim headerRow As Long
    headerRow = FindHeaderRow(srcWs)
    
    Dim startColIdx As Long
    startColIdx = Range(cfg.FirstDataCol & "1").Column
    
    For m = 1 To cfg.numMonths
        colIdx = 3 + (m - 1)   ' cot C = thang 1
        colLetter = Split(Cells(1, colIdx).Address(True, False), "$")(0)
        
        If headerRow > 0 Then
            Dim dateVal As Variant
            dateVal = srcWs.Cells(headerRow, startColIdx + m - 1).Value
            If IsDate(dateVal) Then
                ws.Cells(2, colIdx).Value = dateVal
                ws.Cells(132, colIdx).Value = dateVal
            Else
                ws.Cells(2, colIdx).Value = DateSerial(Year(Date), m, 1)
                ws.Cells(132, colIdx).Value = DateSerial(Year(Date), m, 1)
            End If
        Else
            ws.Cells(2, colIdx).Value = DateSerial(Year(Date), m, 1)
            ws.Cells(132, colIdx).Value = DateSerial(Year(Date), m, 1)
        End If
        
        ws.Cells(2, colIdx).NumberFormat = "dd/mm/yyyy"
        ws.Cells(132, colIdx).NumberFormat = "dd/mm/yyyy"
    Next m

    ' ----- To mau header (dong 2 va 132) voi RGB(31,78,152) -----
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 2 + cfg.numMonths))
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(132, 1), ws.Cells(132, 2 + cfg.numMonths))
        .Interior.Color = RGB(31, 78, 152)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ----- Ghi template -----
    tpl = ParseTemplate()
    For i = LBound(tpl) To UBound(tpl)
        ws.Cells(tpl(i).RowNum, 1).Value = tpl(i).Label
        If tpl(i).Code <> "" Then ws.Cells(tpl(i).RowNum, 2).Value = tpl(i).Code

        If tpl(i).Kind = "E" Then
            ws.Cells(tpl(i).RowNum, 1).Font.Bold = True
        End If

        If tpl(i).Kind <> "B" Then
            For m = 1 To cfg.numMonths
                colIdx = 3 + (m - 1)
                colLetter = Split(Cells(1, colIdx).Address(True, False), "$")(0)
                frm = BuildFormula(tpl(i), cfg, colLetter, m)
                If frm <> "" Then
                    ws.Cells(tpl(i).RowNum, colIdx).Formula = frm
                    ws.Cells(tpl(i).RowNum, colIdx).NumberFormat = ACC_FORMAT
                End If
            Next m
        End If
    Next i

    ws.Columns("A:B").AutoFit
End Sub

'---------------------- 7. CHAY CHO TAT CA ENTITY -----------------------
Public Sub BuildAllEntitySheets()
    Dim cfgs() As EntityConfig
    Dim i As Long

    cfgs = GetEntityConfigs()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = LBound(cfgs) To UBound(cfgs)
        If SheetExists(ThisWorkbook, cfgs(i).DataSheetName) Then
            BuildEntitySheet cfgs(i)
        Else
            Debug.Print "Bo qua sheet khong ton tai: " & cfgs(i).DataSheetName
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox "Da tao xong " & (UBound(cfgs) - LBound(cfgs) + 1) & " sheet chuan hoa." & vbCrLf & _
           "Kiem tra dong 130 (Check) = 0 la dung.", vbInformation
End Sub

'---------------------- UTILITY --------------------------
Private Function SheetExists(ByVal wb As Workbook, ByVal shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'=====================================================================
' 8. GOP TAT CA ENTITY THANH 2 SHEET HOP NHAT (CONSOLIDATION)
'    - Dong "con" (Kind = D): = tong cua chinh o do tren TAT CA sheet
'      entity (001-DOF, 002-ADF, ... da duoc BuildEntitySheet tao ra)
'    - Dong "me" (Kind = R / L / E, vd Tai san ngan han, Tong tai san...):
'      KHONG cong truc tiep dong me cua tung entity, ma tinh lai bang
'      cach SUM/expression tu cac dong con NGAY TRONG SHEET HOP NHAT
'      (dung lai template Kind/Args - dam bao dong me = tong dong con
'      sau khi da gop, giong het logic tung sheet entity dang dung).
'=====================================================================

Public Sub BuildConsolidatedSheets()
    Dim cfgs() As EntityConfig
    cfgs = GetEntityConfigs()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    BuildConsolSheet cfgs, "BS", "Consol - BS", 1, 130
    BuildConsolSheet cfgs, "PL", "Consol - PL", 131, 150

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox "Da gop xong 2 sheet hop nhat: 'Consol - BS' va 'Consol - PL'." & vbCrLf & _
           "Kiem tra dong 130 (Check) trong sheet 'Consol - BS' = 0 la dung.", vbInformation
End Sub

Private Sub BuildConsolSheet(ByRef cfgs() As EntityConfig, ByVal section As String, _
                              ByVal sheetName As String, ByVal titleRow As Long, ByVal sectionLastRow As Long)
    Dim ws As Worksheet
    Dim tpl() As TplRow
    Dim i As Long, m As Long, colIdx As Long, colLetter As String
    Dim j As Long, k As Long
    Dim frm As String, sumParts As String
    Const ACC_FORMAT As String = "#,##0;(#,##0);""-"""

    ' ----- Xac dinh danh sach cac sheet entity DA DUOC TAO (ton tai trong workbook) -----
    Dim entNames() As String
    Dim entMonths() As Long
    Dim nEnt As Long
    ReDim entNames(LBound(cfgs) To UBound(cfgs))
    ReDim entMonths(LBound(cfgs) To UBound(cfgs))
    nEnt = -1
    For k = LBound(cfgs) To UBound(cfgs)
        If SheetExists(ThisWorkbook, cfgs(k).EntityCode) Then
            nEnt = nEnt + 1
            entNames(nEnt) = cfgs(k).EntityCode
            entMonths(nEnt) = cfgs(k).numMonths
        End If
    Next k

    If nEnt < 0 Then
        MsgBox "Chua co sheet entity nao duoc tao (chay BuildAllEntitySheets truoc). Bo qua " & sheetName, vbExclamation
        Exit Sub
    End If
    ReDim Preserve entNames(0 To nEnt)
    ReDim Preserve entMonths(0 To nEnt)

    Dim numMonths As Long
    numMonths = entMonths(0)
    For j = 0 To nEnt
        If entMonths(j) > numMonths Then numMonths = entMonths(j)
    Next j

    ' ----- Tao/lam moi sheet hop nhat -----
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = sheetName

    ' ----- Tieu de + dong header (cung vi tri dong voi sheet entity: 1/2 cho BS, 131/132 cho PL) -----
    If section = "BS" Then
        ws.Cells(titleRow, 1).Value = "BALANCE SHEET - CONSOLIDATED"
    Else
        ws.Cells(titleRow, 1).Value = "PROFIT AND LOSS - CONSOLIDATED"
    End If
    ws.Cells(titleRow, 1).Font.Bold = True

    Dim headerRow As Long
    headerRow = titleRow + 1
    ws.Cells(headerRow, 1).Value = "Chi tieu"
    ws.Cells(headerRow, 2).Value = "Ma so"
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 2)).Font.Bold = True

    ' ----- Ngay thang: lay tu sheet entity dau tien dang co (cung dong header voi sheet entity) -----
    Dim srcWs As Worksheet
    Set srcWs = ThisWorkbook.Worksheets(entNames(0))
    For m = 1 To numMonths
        colIdx = 3 + (m - 1)
        ws.Cells(headerRow, colIdx).Value = srcWs.Cells(headerRow, colIdx).Value
        ws.Cells(headerRow, colIdx).NumberFormat = "dd/mm/yyyy"
    Next m

    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 2 + numMonths))
        .Interior.Color = RGB(31, 78, 152)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ----- Ghi tung dong theo template, chi lay dung Section (BS hoac PL) -----
    tpl = ParseTemplate()
    Dim dummyCfg As EntityConfig   ' chi de truyen tham so cho BuildFormula khi Kind <> D (khong dung toi)

    For i = LBound(tpl) To UBound(tpl)
        If tpl(i).section = section Then
            ws.Cells(tpl(i).RowNum, 1).Value = tpl(i).Label
            If tpl(i).Code <> "" Then ws.Cells(tpl(i).RowNum, 2).Value = tpl(i).Code
            If tpl(i).Kind = "E" Then ws.Cells(tpl(i).RowNum, 1).Font.Bold = True

            If tpl(i).Kind = "D" Then
                ' DONG CON: cong dung o do (cung so dong, cung cot thang) tu TAT CA sheet entity dang co
                For m = 1 To numMonths
                    colIdx = 3 + (m - 1)
                    colLetter = Split(Cells(1, colIdx).Address(True, False), "$")(0)
                    sumParts = ""
                    For j = 0 To nEnt
                        If sumParts <> "" Then sumParts = sumParts & "+"
                        sumParts = sumParts & "'" & entNames(j) & "'!" & colLetter & tpl(i).RowNum
                    Next j
                    ws.Cells(tpl(i).RowNum, colIdx).Formula = "=IFERROR(" & sumParts & ","""")"
                    ws.Cells(tpl(i).RowNum, colIdx).NumberFormat = ACC_FORMAT
                Next m

            ElseIf tpl(i).Kind = "R" Or tpl(i).Kind = "L" Or tpl(i).Kind = "E" Then
                ' DONG ME: KHONG cong dong me cua tung entity, ma tinh lai = SUM/bieu thuc cua
                ' cac dong con NGAY TRONG SHEET HOP NHAT NAY (dung lai BuildFormula co san,
                ' Kind R/L/E khong dung den cfg nen truyen dummyCfg la duoc)
                For m = 1 To numMonths
                    colIdx = 3 + (m - 1)
                    colLetter = Split(Cells(1, colIdx).Address(True, False), "$")(0)
                    frm = BuildFormula(tpl(i), dummyCfg, colLetter, m)
                    If frm <> "" Then
                        ws.Cells(tpl(i).RowNum, colIdx).Formula = frm
                        ws.Cells(tpl(i).RowNum, colIdx).NumberFormat = ACC_FORMAT
                    End If
                Next m
            End If
            ' Kind = "B": dong trong, khong lam gi (giong BuildEntitySheet)
        End If
    Next i

    ws.Columns("A:B").AutoFit
End Sub


