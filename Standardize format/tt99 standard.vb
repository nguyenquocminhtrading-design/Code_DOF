Option Explicit
'=====================================================================
' MODULE: mod_BuildStandardFS  (updated per Circular 99 - TT99 format)
' MUC DICH: Tao sheet BCTC chuan hoa theo TT99 cho tung entity
'           (DOF, ADF, CAN, DOAU, DW, TW, HK, DDFH)
'=====================================================================

'---------------------- 1. CAU TRUC DU LIEU ---------------------------
Public Type EntityConfig
    EntityCode      As String
    DataSheetName   As String
    CodeCol         As String
    FirstDataCol    As String
    NumMonths       As Long
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
    Section As String   ' BS hoac PL
End Type

'---------------------- 2. TEMPLATE CHUAN (TT99) ---------------------------
Private Function GetTemplateLines() As String
    Dim s As String
    s = ""
    ' ========== BALANCE SHEET (TT99) ==========
    s = s & "3|SHORT-TERM ASSETS|100|E|{110}+{120}+{130}+{140}+{150}+{160}|BS" & vbCrLf
    s = s & "4|I. Cash and cash equivalents|110|R|111:112|BS" & vbCrLf
    s = s & "5|1. Money|111|D||BS" & vbCrLf
    s = s & "6|2. Cash equivalents|112|D||BS" & vbCrLf
    s = s & "7|II. Short-term financial investment|120|R|121:126|BS" & vbCrLf
    s = s & "8|1. Trading securities|121|D||BS" & vbCrLf
    s = s & "9|2. Provision for depreciation of trading securities (*)|122|D||BS" & vbCrLf
    s = s & "10|3. Investments held to short-term maturity|123|D||BS" & vbCrLf
    s = s & "11|4. Provision for investment held to short-term maturity (*)|124|D||BS" & vbCrLf
    s = s & "12|5. Other short-term investments|125|D||BS" & vbCrLf
    s = s & "13|6. Provision for losses of other short-term investments (*)|126|D||BS" & vbCrLf
    s = s & "14|III. Short-term receivables|130|R|131:137|BS" & vbCrLf
    s = s & "15|1. Short-term receivables of customers|131|D||BS" & vbCrLf
    s = s & "16|2. Short-term seller upfront|132|D||BS" & vbCrLf
    s = s & "17|3. Short-term internal receivables|133|D||BS" & vbCrLf
    s = s & "18|4. Receivables according to the construction contract schedule|134|D||BS" & vbCrLf
    s = s & "19|5. Other short-term receivables|135|D||BS" & vbCrLf
    s = s & "20|6. Provision for short-term bad debts (*)|136|D||BS" & vbCrLf
    s = s & "21|7. Pending shortage of assets|137|D||BS" & vbCrLf
    s = s & "22|IV. Inventory|140|R|141:142|BS" & vbCrLf
    s = s & "23|1. Inventory|141|D||BS" & vbCrLf
    s = s & "24|2. Provision for inventory discounts (*)|142|D||BS" & vbCrLf
    s = s & "25|V. Short-Term Biological Assets|150|R|151:153|BS" & vbCrLf
    s = s & "26|1. Pets take short-term one-time products|151|D||BS" & vbCrLf
    s = s & "27|2. Seasonal crops or short-term one-off products|152|D||BS" & vbCrLf
    s = s & "28|3. Provision for short-term biological asset losses (*)|153|D||BS" & vbCrLf
    s = s & "29|VI. Other short-term assets|160|R|161:165|BS" & vbCrLf
    s = s & "30|1. Short-term allocation pending costs|161|D||BS" & vbCrLf
    s = s & "31|2. Deductible VAT|162|D||BS" & vbCrLf
    s = s & "32|3. Taxes and other amounts receivable by the State|163|D||BS" & vbCrLf
    s = s & "33|4. Repurchase and sale of government bonds|164|D||BS" & vbCrLf
    s = s & "34|5. Other short-term assets|165|D||BS" & vbCrLf

    s = s & "35|LONG-TERM ASSETS|200|E|{210}+{220}+{230}+{240}+{250}+{260}+{270}|BS" & vbCrLf
    s = s & "36|I. Long-term receivables|210|R|211:216|BS" & vbCrLf
    s = s & "37|1. Long-term receivables of customers|211|D||BS" & vbCrLf
    s = s & "38|2. Pay long-term sellers upfront|212|D||BS" & vbCrLf
    s = s & "39|3. Business capital in affiliated units|213|D||BS" & vbCrLf
    s = s & "40|4. Long-term internal receivables|214|D||BS" & vbCrLf
    s = s & "41|5. Other long-term receivables|215|D||BS" & vbCrLf
    s = s & "42|6. Provision for long-term bad debts (*)|216|D||BS" & vbCrLf

    s = s & "43|II. Fixed assets|220|E|{221}+{224}+{227}|BS" & vbCrLf
    s = s & "44|1. Tangible fixed assets|221|E|{222}+{223}|BS" & vbCrLf
    s = s & "45|- Historical cost|222|D||BS" & vbCrLf
    s = s & "46|- Cumulative wear value (*)|223|D||BS" & vbCrLf
    s = s & "47|2. Fixed assets lease finance|224|E|{225}+{226}|BS" & vbCrLf
    s = s & "48|- Historical cost|225|D||BS" & vbCrLf
    s = s & "49|- Cumulative wear value (*)|226|D||BS" & vbCrLf
    s = s & "50|3. Intangible fixed assets|227|E|{228}+{229}|BS" & vbCrLf
    s = s & "51|- Historical cost|228|D||BS" & vbCrLf
    s = s & "52|- Cumulative wear value (*)|229|D||BS" & vbCrLf

    s = s & "53|III. Long-term biological assets|230|E|{231}+{236}+{237}+{238}|BS" & vbCrLf
    s = s & "54|1. Pet Animals for Periodic Products|231|D||BS" & vbCrLf
    s = s & "55|a) Livestock for periodic products that have not reached the adult stage|232|D||BS" & vbCrLf
    s = s & "56|b) Livestock for periodic products up to the adult stage|233|D||BS" & vbCrLf
    s = s & "57|- Historical cost|234|D||BS" & vbCrLf
    s = s & "58|- Accumulated depreciation value (*)|235|D||BS" & vbCrLf
    s = s & "59|2. Pets take long-term one-time products|236|D||BS" & vbCrLf
    s = s & "60|3. Seasonal crops or long-term one-time products|237|D||BS" & vbCrLf
    s = s & "61|4. Provision for long-term biological asset losses (*)|238|D||BS" & vbCrLf

    s = s & "62|IV. Investment real estate|240|E|{241}+{242}|BS" & vbCrLf
    s = s & "63|- Historical cost|241|D||BS" & vbCrLf
    s = s & "64|- Cumulative wear value (*)|242|D||BS" & vbCrLf

    s = s & "65|V. Long-term unfinished assets|250|E|{251}+{252}|BS" & vbCrLf
    s = s & "66|1. Long-term unfinished production and business expenses|251|D||BS" & vbCrLf
    s = s & "67|2. Unfinished capital construction costs|252|D||BS" & vbCrLf

    s = s & "68|VI. Long-term financial investment|260|R|261:266|BS" & vbCrLf
    s = s & "69|1. Invest in subsidiaries|261|D||BS" & vbCrLf
    s = s & "70|2. Investment in joint venture or associate companies|262|D||BS" & vbCrLf
    s = s & "71|3. Investment in capital contribution to other units|263|D||BS" & vbCrLf
    s = s & "72|4. Provision for long-term investment losses in other units (*)|264|D||BS" & vbCrLf
    s = s & "73|5. Investments held to long-term maturity|265|D||BS" & vbCrLf
    s = s & "74|6. Provision for investment held to long-term maturity (*)|266|D||BS" & vbCrLf

    s = s & "75|VII. Other long-term assets|270|R|271:274|BS" & vbCrLf
    s = s & "76|1. Long-term allocation waiting costs|271|D||BS" & vbCrLf
    s = s & "77|2. Deferred income tax assets|272|D||BS" & vbCrLf
    s = s & "78|3. Long-term equipment, supplies, spare parts|273|D||BS" & vbCrLf
    s = s & "79|4. Other long-term assets|274|D||BS" & vbCrLf

    s = s & "80|TOTAL ASSETS|280|E|{100}+{200}|BS" & vbCrLf

    s = s & "81|LIABILITIES|300|E|{310}+{330}|BS" & vbCrLf
    s = s & "82|I. Short-term debt|310|R|311:325|BS" & vbCrLf
    s = s & "83|1. Payable to short-term sellers|311|D||BS" & vbCrLf
    s = s & "84|2. Buyer pays in advance for a short term|312|D||BS" & vbCrLf
    s = s & "85|3. Dividends and profits must be paid|313|D||BS" & vbCrLf
    s = s & "86|4. Taxes and short-term payables to the State|314|D||BS" & vbCrLf
    s = s & "87|5. Payables to employees|315|D||BS" & vbCrLf
    s = s & "88|6. Short-term expenses|316|D||BS" & vbCrLf
    s = s & "89|7. Short-term internal payables|317|D||BS" & vbCrLf
    s = s & "90|8. Payable according to the schedule of short-term construction contracts|318|D||BS" & vbCrLf
    s = s & "91|9. Revenue pending short-term allocation|319|D||BS" & vbCrLf
    s = s & "92|10. Short-term financial loans and leases|320|D||BS" & vbCrLf
    s = s & "93|11. Other short-term payables|321|D||BS" & vbCrLf
    s = s & "94|12. Provision for short-term payables|322|D||BS" & vbCrLf
    s = s & "95|13. Reward and welfare funds|323|D||BS" & vbCrLf
    s = s & "96|14. Price Stabilization Fund|324|D||BS" & vbCrLf
    s = s & "97|15. Repurchase and sale of government bonds|325|D||BS" & vbCrLf

    s = s & "98|II. Long-term debt|330|R|331:344|BS" & vbCrLf
    s = s & "99|1. Payable to long-term sellers|331|D||BS" & vbCrLf
    s = s & "100|2. Buyers pay in advance for the long term|332|D||BS" & vbCrLf
    s = s & "101|3. Taxes and long-term payables to the State|333|D||BS" & vbCrLf
    s = s & "102|4. Long-term expenses|334|D||BS" & vbCrLf
    s = s & "103|5. Internal payables for business capital|335|D||BS" & vbCrLf
    s = s & "104|6. Long-term internal payables|336|D||BS" & vbCrLf
    s = s & "105|7. Revenue pending long-term attribution|337|D||BS" & vbCrLf
    s = s & "106|8. Other long-term payables|338|D||BS" & vbCrLf
    s = s & "107|9. Long-term financial loans and leases|339|D||BS" & vbCrLf
    s = s & "108|10. Convertible Bonds|340|D||BS" & vbCrLf
    s = s & "109|11. Preferred Stocks|341|D||BS" & vbCrLf
    s = s & "110|12. Deferred income tax payable|342|D||BS" & vbCrLf
    s = s & "111|13. Long-term payable provisions|343|D||BS" & vbCrLf
    s = s & "112|14. The Scientific and Technological Development Fund|344|D||BS" & vbCrLf

    s = s & "113|EQUITY|400|E|{410}+{430}|BS" & vbCrLf
    s = s & "114|I. Equity|410|E|{411}+{412}+{413}+{414}+{415}+{416}+{417}+{418}+{419}+{420}|BS" & vbCrLf
    s = s & "115|1. Owner's contributed capital|411|E|{116}+{117}|BS" & vbCrLf
    s = s & "116|- Voting common shares|411a|D||BS" & vbCrLf
    s = s & "117|- Preferred stock|411b|D||BS" & vbCrLf
    s = s & "118|2. Surplus of equity capital|412|D||BS" & vbCrLf
    s = s & "119|3. Bond Conversion Options|413|D||BS" & vbCrLf
    s = s & "120|4. Other capital of the owner|414|D||BS" & vbCrLf
    s = s & "121|5. Own redemption shares (*)|415|D||BS" & vbCrLf
    s = s & "122|6. Asset revaluation difference|416|D||BS" & vbCrLf
    s = s & "123|7. Exchange Rate Differences|417|D||BS" & vbCrLf
    s = s & "124|8. Development investment funds|418|D||BS" & vbCrLf
    s = s & "125|9. Other funds belonging to the owner's equity|419|D||BS" & vbCrLf
    s = s & "126|10. Undistributed after-tax profit|420|E|{127}+{128}|BS" & vbCrLf
    s = s & "127|- Accumulated undistributed profit by the end of the previous period|420a|D||BS" & vbCrLf
    s = s & "128|- Undistributed profit for this period|420b|D||BS" & vbCrLf
    s = s & "129|II. Other funding sources and funds|430|E|{130}+{131}|BS" & vbCrLf
    s = s & "130|1. Funding sources|431|D||BS" & vbCrLf
    s = s & "131|2. Funding sources that have formed fixed assets|432|D||BS" & vbCrLf
    s = s & "132|TOTAL CAPITAL|440|E|{300}+{400}|BS" & vbCrLf
    s = s & "133|Check (should be 0)||E|{80}-{132}|BS" & vbCrLf

    ' ========== PROFIT AND LOSS (TT99) ==========
    s = s & "140|1. Revenue from sales and provision of services|1|D||PL" & vbCrLf
    s = s & "141|2. Revenue deductions|2|D||PL" & vbCrLf
    s = s & "142|3. Net revenue|10|E|{140}-{141}|PL" & vbCrLf
    s = s & "143|4. Cost of goods sold|11|D||PL" & vbCrLf
    s = s & "144|5. Gross profit|20|E|{142}-{143}|PL" & vbCrLf
    s = s & "145|6. Profit/loss of sale and liquidation of investment real estate|21|D||PL" & vbCrLf
    s = s & "146|7. Revenue from financial activities|22|D||PL" & vbCrLf
    s = s & "147|8. Financing costs|23|D||PL" & vbCrLf
    s = s & "148|- In which: Borrowing costs|24|D||PL" & vbCrLf
    s = s & "149|9. Cost of sales|25|D||PL" & vbCrLf
    s = s & "150|10. Business management expenses|26|D||PL" & vbCrLf
    s = s & "151|11. Net profit from business activities|30|E|{144}+{145}+{146}-{147}-{149}-{150}|PL" & vbCrLf
    s = s & "152|12. Other income|31|D||PL" & vbCrLf
    s = s & "153|13. Other expenses|32|D||PL" & vbCrLf
    s = s & "154|14. Other Profits|40|E|{152}-{153}|PL" & vbCrLf
    s = s & "155|15. Total accounting profit before tax|50|E|{151}+{154}|PL" & vbCrLf
    s = s & "156|16. Current CIT expenses|51|D||PL" & vbCrLf
    s = s & "157|17. Deferred CIT expenses|52|D||PL" & vbCrLf
    s = s & "158|18. Profit after corporate income tax|60|E|{155}-{156}-{157}|PL" & vbCrLf
    s = s & "159|19. Basic earnings per share (*)|71|D||PL" & vbCrLf
    s = s & "160|20. Declining interest on stocks (*)|72|D||PL" & vbCrLf

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
            result(n).Section = parts(5)
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
    cfgs(0).NumMonths = 12
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
    cfgs(1).NumMonths = 12
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
    cfgs(2).NumMonths = 12
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
    cfgs(3).NumMonths = 12
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
    cfgs(4).NumMonths = 12
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
    cfgs(5).NumMonths = 12
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
    cfgs(6).NumMonths = 12
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
    cfgs(7).NumMonths = 12
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
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5

    Dim keywords As Variant
    keywords = Array("code", "ma so", "mã s?", "items", "ch? tiêu", "description", "stt", "no.")

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

' ===== FIXED ByRef -> ByVal HERE =====
Private Function BuildFormula(ByVal t As TplRow, ByVal cfg As EntityConfig, ByVal colLetter As String, ByVal monthIdx As Long) As String
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

            If t.Section = "PL" And cfg.PLFirstRow > 0 Then
                dataRangeFrom = dataCol & cfg.PLFirstRow
                dataRangeTo = dataCol & cfg.PLLastRow
                codeRangeFrom = cfg.CodeCol & cfg.PLFirstRow
                codeRangeTo = cfg.CodeCol & cfg.PLLastRow
            ElseIf t.Section = "BS" And cfg.BSFirstRow > 0 Then
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
' ===== FIXED ByRef -> ByVal HERE =====
Public Sub BuildEntitySheet(ByVal cfg As EntityConfig)
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
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
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

    ' ----- L?y ngày tháng t? sheet ngu?n (header row) -----
    Dim srcWs As Worksheet
    Set srcWs = ThisWorkbook.Sheets(cfg.DataSheetName)
    Dim headerRow As Long
    headerRow = FindHeaderRow(srcWs)
    
    Dim startColIdx As Long
    startColIdx = Range(cfg.FirstDataCol & "1").Column
    
    For m = 1 To cfg.NumMonths
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

    ' ----- Tô màu header (d?ng 2 và 132) v?i RGB(31,78,152) -----
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 2 + cfg.NumMonths))
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(132, 1), ws.Cells(132, 2 + cfg.NumMonths))
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
            For m = 1 To cfg.NumMonths
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

    MsgBox "Da tao xong " & (UBound(cfgs) - LBound(cfgs) + 1) & " sheet chuan hoa theo TT99." & vbCrLf & _
           "Kiem tra dong 133 (Check) = 0 la dung.", vbInformation
End Sub

'---------------------- UTILITY --------------------------
Private Function SheetExists(ByVal wb As Workbook, ByVal shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function