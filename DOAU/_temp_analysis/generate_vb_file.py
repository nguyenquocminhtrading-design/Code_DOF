import openpyxl
import re

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

true_headers = {
    # BS
    4, 5, 6, 12, 16, 17, 25, 39, 44, 45, 46, 48, 53, 55, 60, 66, 79,
    # PL
    90, 91, 92, 105, 106, 112, 119, 136, 158, 159, 161, 175, 186, 205, 220, 221, 225
}

true_formulas = {
    11: "=ROUND(SUM(D6:D10),5)",
    15: "=ROUND(SUM(D12:D14),5)",
    20: "=ROUND(SUM(D17:D19),5)",
    23: "=ROUND(D16+SUM(D20:D22),5)",
    24: "=ROUND(D5+D11+D15+D23,5)",
    38: "=ROUND(SUM(D25:D37),5)",
    42: "=ROUND(SUM(D39:D41),5)",
    43: "=ROUND(D4+D24+D38+D42,5)",
    52: "=ROUND(SUM(D48:D51),5)",
    59: "=ROUND(SUM(D56:D58),5)",
    63: "=ROUND(SUM(D60:D62),5)",
    64: "=ROUND(SUM(D53:D54)+D59+D63,5)",
    65: "=ROUND(D47+D52+D64,5)",
    77: "=ROUND(SUM(D66:D76),5)",
    78: "=ROUND(+D65+D77,5)",
    84: "=ROUND(SUM(D79:D83),5)",
    85: "=ROUND(D46+D78+D84,5)",
    86: "=D85-D43",
    97: "=ROUND(SUM(D92:D96),5)",
    103: "=ROUND(D91+SUM(D97:D102),5)",
    104: "=ROUND(D90+D103,5)",
    111: "=ROUND(SUM(D106:D110),5)",
    118: "=ROUND(SUM(D112:D117),5)",
    140: "=ROUND(SUM(D136:D139),5)",
    155: "=ROUND(SUM(D119:D135)+SUM(D140:D154),5)",
    156: "=ROUND(D105+D111+D118+D155,5)",
    157: "=ROUND(D104-D156,5)",
    166: "=ROUND(SUM(D161:D165),5)",
    180: "=ROUND(SUM(D175:D179),5)",
    194: "=ROUND(SUM(D186:D193),5)",
    204: "=ROUND(SUM(D159:D160)+SUM(D166:D174)+SUM(D180:D185)+SUM(D194:D203),5)",
    216: "=ROUND(SUM(D205:D215),5)",
    218: "=ROUND(D158+D204+SUM(D216:D217),5)",
    219: "=ROUND(D157-D218,5)",
    224: "=ROUND(SUM(D221:D223),5)",
    229: "=ROUND(SUM(D225:D228),5)",
    230: "=ROUND(D224-D229,5)",
    231: "=ROUND(D219+D230,5)",
    234: "=SUM(D231:D233)"
}

rows_spec = []

for r in range(1, 235):
    txt = ""
    lvl = 0
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v and str(v).strip():
            txt = str(v).strip()
            lvl = c
            break
            
    row_type = ""
    formula_spec = ""
    
    if r == 1:
        row_type = "TITLE"
        txt = "FINANCIAL STATEMENT SUMMARY"
    elif r == 2:
        row_type = "TITLE"
        txt = "BALANCE SHEET"
    elif r == 3:
        row_type = "HEADER_MONTH"
        txt = "Month"
    elif r in [87, 88, 232]:
        row_type = "BLANK"
        txt = ""
    elif r == 89:
        row_type = "TITLE"
        txt = "PROFIT AND LOSS"
    elif r in true_formulas:
        row_type = "FORMULA"
        formula_spec = true_formulas[r]
    elif r in true_headers:
        row_type = "HEADER"
    else:
        if r < 89:
            row_type = "BS_SUMIFS"
        else:
            row_type = "PL_SUMIFS"
            
    txt_escaped = txt.replace('"', '""')
    rows_spec.append((r, lvl, txt_escaped, row_type, formula_spec))

print(f"Generated {len(rows_spec)} row specifications.")

vba_lines = []
vba_lines.append("Option Explicit")
vba_lines.append("")
vba_lines.append("' =====================================================================")
vba_lines.append("' DOAU_v2.bas - Creates a 'Summary' sheet matching 100% row-to-row with")
vba_lines.append("' '007-DOAU data' (Rows 1 to 234).")
vba_lines.append("'")
vba_lines.append("' Key features:")
vba_lines.append("' 1. Exact Row-to-Row Alignment:")
vba_lines.append("'    Row r in Summary = Row r in 007-DOAU data.")
vba_lines.append("'    Allows 1-click Copy-Paste Values of D4:D234 directly into 007-DOAU data.")
vba_lines.append("' 2. Dynamic SUMIFS Criteria:")
vba_lines.append("'    Uses column B cell reference (e.g. B7, B8...) instead of hardcoded strings.")
vba_lines.append("' 3. Standardized Subtotal Formulas:")
vba_lines.append("'    Matches the exact ROUND/SUM accounting formulas of 007-DOAU data.")
vba_lines.append("' =====================================================================")
vba_lines.append("Sub CreateSummaryDOAU()")
vba_lines.append("    Dim wb As Workbook")
vba_lines.append("    Dim wsNew As Worksheet, wsBS As Worksheet, wsPL As Worksheet")
vba_lines.append("    Dim i As Long, r As Long")
vba_lines.append("    Dim lastRowBS As Long, lastRowPL As Long")
vba_lines.append("    Dim inputMonth As String")
vba_lines.append("    Dim colAfterBS As String, colAfterPL As String")
vba_lines.append("    Dim helperLevelBS As Long, helperNameBS As Long")
vba_lines.append("    Dim helperLevelPL As Long, helperNamePL As Long")
vba_lines.append("    Dim colNameBS As String, colNamePL As String")
vba_lines.append("    Dim rowsMap() As Variant")
vba_lines.append("    Dim lvl As Long, accountName As String, rowType As String, formulaSpec As String")
vba_lines.append("")
vba_lines.append("    Set wb = ThisWorkbook")
vba_lines.append("")
vba_lines.append("    inputMonth = InputBox(\"Enter the month this file covers (e.g. Sep 2026):\", \"Month Selection\", \"Sep 2026\")")
vba_lines.append("    If inputMonth = \"\" Then Exit Sub")
vba_lines.append("")
vba_lines.append("    On Error Resume Next")
vba_lines.append("    Set wsBS = wb.Sheets(\"BS\")")
vba_lines.append("    Set wsPL = wb.Sheets(\"PL\")")
vba_lines.append("    On Error GoTo 0")
vba_lines.append("    If wsBS Is Nothing Then MsgBox \"Sheet 'BS' not found.\", vbCritical: Exit Sub")
vba_lines.append("    If wsPL Is Nothing Then MsgBox \"Sheet 'PL' not found.\", vbCritical: Exit Sub")
vba_lines.append("")
vba_lines.append("    colAfterBS = FindColumnByHeader(wsBS, 1, \"After\")")
vba_lines.append("    colAfterPL = FindColumnByHeader(wsPL, 1, \"After\")")
vba_lines.append("    If colAfterBS = \"\" Then MsgBox \"Could not find the 'After' column in BS row 1.\", vbCritical: Exit Sub")
vba_lines.append("    If colAfterPL = \"\" Then MsgBox \"Could not find the 'After' column in PL row 1.\", vbCritical: Exit Sub")
vba_lines.append("")
vba_lines.append("    lastRowBS = wsBS.Cells(wsBS.Rows.Count, \"A\").End(xlUp).Row")
vba_lines.append("    lastRowPL = wsPL.Cells(wsPL.Rows.Count, \"A\").End(xlUp).Row")
vba_lines.append("    Dim tmpLast As Long")
vba_lines.append("    For i = 2 To 7")
vba_lines.append("        tmpLast = wsBS.Cells(wsBS.Rows.Count, i).End(xlUp).Row")
vba_lines.append("        If tmpLast > lastRowBS Then lastRowBS = tmpLast")
vba_lines.append("        tmpLast = wsPL.Cells(wsPL.Rows.Count, i).End(xlUp).Row")
vba_lines.append("        If tmpLast > lastRowPL Then lastRowPL = tmpLast")
vba_lines.append("    Next i")
vba_lines.append("")
vba_lines.append("    helperLevelBS = AddHelperColumns(wsBS, lastRowBS)")
vba_lines.append("    helperLevelPL = AddHelperColumns(wsPL, lastRowPL)")
vba_lines.append("    helperNameBS = helperLevelBS + 1")
vba_lines.append("    helperNamePL = helperLevelPL + 1")
vba_lines.append("")
vba_lines.append("    colNameBS = Split(wsBS.Cells(1, helperNameBS).Address, \"$\")(1)")
vba_lines.append("    colNamePL = Split(wsPL.Cells(1, helperNamePL).Address, \"$\")(1)")
vba_lines.append("")
vba_lines.append("    ' Recreate the Summary sheet")
vba_lines.append("    On Error Resume Next")
vba_lines.append("    Application.DisplayAlerts = False")
vba_lines.append("    wb.Sheets(\"Summary\").Delete")
vba_lines.append("    Application.DisplayAlerts = True")
vba_lines.append("    On Error GoTo 0")
vba_lines.append("")
vba_lines.append("    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))")
vba_lines.append("    wsNew.Name = \"Summary\"")
vba_lines.append("")
vba_lines.append("    ' Initialize row map (1 to 234)")
vba_lines.append("    InitRowsMap rowsMap")
vba_lines.append("")
vba_lines.append("    Application.ScreenUpdating = False")
vba_lines.append("")
vba_lines.append("    For r = 1 To 234")
vba_lines.append("        lvl = rowsMap(r)(0)")
vba_lines.append("        accountName = rowsMap(r)(1)")
vba_lines.append("        rowType = rowsMap(r)(2)")
vba_lines.append("        formulaSpec = rowsMap(r)(3)")
vba_lines.append("")
vba_lines.append("        Select Case rowType")
vba_lines.append("            Case \"TITLE\"")
vba_lines.append("                wsNew.Cells(r, \"A\").Value = accountName")
vba_lines.append("                wsNew.Cells(r, \"A\").Font.Bold = True")
vba_lines.append("            Case \"HEADER_MONTH\"")
vba_lines.append("                wsNew.Cells(r, \"D\").Value = inputMonth")
vba_lines.append("                wsNew.Cells(r, \"D\").Font.Bold = True")
vba_lines.append("                wsNew.Cells(r, \"D\").HorizontalAlignment = xlRight")
vba_lines.append("            Case \"HEADER\"")
vba_lines.append("                wsNew.Cells(r, \"B\").Value = accountName")
vba_lines.append("                wsNew.Cells(r, \"B\").IndentLevel = lvl - 1")
vba_lines.append("                If lvl <= 2 Then wsNew.Cells(r, \"B\").Font.Bold = True")
vba_lines.append("            Case \"BS_SUMIFS\"")
vba_lines.append("                wsNew.Cells(r, \"B\").Value = accountName")
vba_lines.append("                wsNew.Cells(r, \"B\").IndentLevel = lvl - 1")
vba_lines.append("                wsNew.Cells(r, \"D\").Formula = _")
vba_lines.append("                    \"=SUMIFS(BS!$\" & colAfterBS & \"$3:$\" & colAfterBS & \"$\" & lastRowBS & \",\" & _")
vba_lines.append("                    \"BS!$\" & colNameBS & \"$3:$\" & colNameBS & \"$\" & lastRowBS & \",B\" & r & \")\"")
vba_lines.append("            Case \"PL_SUMIFS\"")
vba_lines.append("                wsNew.Cells(r, \"B\").Value = accountName")
vba_lines.append("                wsNew.Cells(r, \"B\").IndentLevel = lvl - 1")
vba_lines.append("                wsNew.Cells(r, \"D\").Formula = _")
vba_lines.append("                    \"=SUMIFS(PL!$\" & colAfterPL & \"$3:$\" & colAfterPL & \"$\" & lastRowPL & \",\" & _")
vba_lines.append("                    \"PL!$\" & colNamePL & \"$3:$\" & colNamePL & \"$\" & lastRowPL & \",B\" & r & \")\"")
vba_lines.append("            Case \"FORMULA\"")
vba_lines.append("                wsNew.Cells(r, \"B\").Value = accountName")
vba_lines.append("                wsNew.Cells(r, \"B\").IndentLevel = lvl - 1")
vba_lines.append("                If lvl <= 2 Then wsNew.Cells(r, \"B\").Font.Bold = True")
vba_lines.append("                wsNew.Cells(r, \"D\").Formula = formulaSpec")
vba_lines.append("            Case \"BLANK\"")
vba_lines.append("                ' Leave blank to keep exact row alignment")
vba_lines.append("        End Select")
vba_lines.append("    Next r")
vba_lines.append("")
vba_lines.append("    ' Formatting")
vba_lines.append("    With wsNew")
vba_lines.append("        .Columns(\"A:B\").AutoFit")
vba_lines.append("        .Columns(\"D\").ColumnWidth = 18")
vba_lines.append("        .Range(\"D4:D234\").NumberFormat = \"#,##0.00;(#,##0.00)\"")
vba_lines.append("    End With")
vba_lines.append("")
vba_lines.append("    Application.ScreenUpdating = True")
vba_lines.append("")
vba_lines.append("    MsgBox \"Sheet 'Summary' created successfully matching 007-DOAU data rows 1:1!\" & vbCrLf & _")
vba_lines.append("           \"You can now copy Summary!D4:D234 directly into 007-DOAU data.\", vbInformation")
vba_lines.append("End Sub")
vba_lines.append("")

vba_lines.append("' =====================================================================")
vba_lines.append("' Helper: Initialize row specifications (1 to 234)")
vba_lines.append("' Format: Array(Level, Name, RowType, FormulaSpec)")
vba_lines.append("' =====================================================================")
vba_lines.append("Private Sub InitRowsMap(ByRef m() As Variant)")
vba_lines.append("    ReDim m(1 To 234)")
vba_lines.append("    InitRowsMapPart1 m")
vba_lines.append("    InitRowsMapPart2 m")
vba_lines.append("    InitRowsMapPart3 m")
vba_lines.append("End Sub")
vba_lines.append("")

vba_lines.append("Private Sub InitRowsMapPart1(ByRef m() As Variant)")
for r, lvl, txt, rt, fs in rows_spec[:88]:
    vba_lines.append(f'    m({r}) = Array({lvl}, "{txt}", "{rt}", "{fs}")')
vba_lines.append("End Sub")
vba_lines.append("")

vba_lines.append("Private Sub InitRowsMapPart2(ByRef m() As Variant)")
for r, lvl, txt, rt, fs in rows_spec[88:165]:
    vba_lines.append(f'    m({r}) = Array({lvl}, "{txt}", "{rt}", "{fs}")')
vba_lines.append("End Sub")
vba_lines.append("")

vba_lines.append("Private Sub InitRowsMapPart3(ByRef m() As Variant)")
for r, lvl, txt, rt, fs in rows_spec[165:]:
    vba_lines.append(f'    m({r}) = Array({lvl}, "{txt}", "{rt}", "{fs}")')
vba_lines.append("End Sub")
vba_lines.append("")

vba_lines.append("' =====================================================================")
vba_lines.append("' Helper: Find column letter by header text")
vba_lines.append("' =====================================================================")
vba_lines.append("Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerText As String) As String")
vba_lines.append("    Dim c As Range")
vba_lines.append("    Dim lastCol As Long")
vba_lines.append("    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column")
vba_lines.append("    If lastCol < 20 Then lastCol = 20")
vba_lines.append("    For Each c In ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))")
vba_lines.append("        If InStr(1, CStr(c.Value), headerText, vbTextCompare) > 0 Then")
vba_lines.append("            FindColumnByHeader = Split(c.Address, \"$\")(1)")
vba_lines.append("            Exit Function")
vba_lines.append("        End If")
vba_lines.append("    Next c")
vba_lines.append("    FindColumnByHeader = \"\"")
vba_lines.append("End Function")
vba_lines.append("")

vba_lines.append("' =====================================================================")
vba_lines.append("' Helper: Add helper columns (Level, Name Cleaned) to BS and PL")
vba_lines.append("' =====================================================================")
vba_lines.append("Function AddHelperColumns(ws As Worksheet, lastRow As Long) As Long")
vba_lines.append("    Dim colLevel As Long, colName As Long")
vba_lines.append("    Dim r As Long")
vba_lines.append("    Dim q As String")
vba_lines.append("    q = Chr(34) & Chr(34)")
vba_lines.append("")
vba_lines.append("    If ws.Cells(1, \"P\").Value = \"Level\" And ws.Cells(1, \"Q\").Value = \"Name Cleaned\" Then")
vba_lines.append("        AddHelperColumns = ws.Range(\"P1\").Column")
vba_lines.append("        Exit Function")
vba_lines.append("    End If")
vba_lines.append("")
vba_lines.append("    colLevel = 16")
vba_lines.append("    colName = 17")
vba_lines.append("")
vba_lines.append("    ws.Cells(1, colLevel).Value = \"Level\"")
vba_lines.append("    ws.Cells(1, colName).Value = \"Name Cleaned\"")
vba_lines.append("    ws.Range(ws.Cells(1, colLevel), ws.Cells(1, colName)).Font.Bold = True")
vba_lines.append("")
vba_lines.append("    Dim formulaLevel As String, formulaName As String")
vba_lines.append("    For r = 3 To lastRow")
vba_lines.append("        formulaLevel = \"=IF(LEN(A\" & r & \")>0,1,IF(LEN(B\" & r & \")>0,2,IF(LEN(C\" & r & \")>0,3,\" & _")
vba_lines.append("                       \"IF(LEN(D\" & r & \")>0,4,IF(LEN(E\" & r & \")>0,5,IF(LEN(F\" & r & \")>0,6,\" & _")
vba_lines.append("                       \"IF(LEN(G\" & r & \")>0,7,\" & q & \")))))))\"")
vba_lines.append("        ws.Cells(r, colLevel).Formula = formulaLevel")
vba_lines.append("")
vba_lines.append("        formulaName = \"=TRIM(PROPER(CLEAN(IF(LEN(A\" & r & \")>0,A\" & r & \",IF(LEN(B\" & r & \")>0,B\" & r & _")
vba_lines.append("                      \",IF(LEN(C\" & r & \")>0,C\" & r & \",IF(LEN(D\" & r & \")>0,D\" & r & _")
vba_lines.append("                      \",IF(LEN(E\" & r & \")>0,E\" & r & \",IF(LEN(F\" & r & \")>0,F\" & r & \",G\" & r & \")))))))))\"")
vba_lines.append("        ws.Cells(r, colName).Formula = formulaName")
vba_lines.append("    Next r")
vba_lines.append("")
vba_lines.append("    AddHelperColumns = colLevel")
vba_lines.append("End Function")
vba_lines.append("")

output_code = "\r\n".join(vba_lines)
with open("DOAU_v2.bas", "w", encoding="utf-8") as f:
    f.write(output_code)

print(f"Successfully generated clean DOAU_v2.bas ({len(output_code)} bytes, {len(vba_lines)} lines).")
