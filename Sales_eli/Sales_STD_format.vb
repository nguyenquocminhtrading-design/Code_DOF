Option Explicit

'============================================================
' MACRO CHINH: Quet tat ca sheet trong workbook hien tai
' (moi sheet la 1 file da duoc paste vao), tu nhan dien format
' + tu dong do dong header (khong can biet truoc header o dong nao)
' va gom ve sheet "Standardized" theo cau truc Format 4 + Currency + OEM
'============================================================
Sub StandardizeSalesReports()
    Dim wsOut As Worksheet
    Dim ws As Worksheet
    Dim outputRow As Long
    Dim rowsImported As Long
    Dim summary As String
    Dim skipped As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wsOut = PrepareOutputSheet()
    outputRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1
    If outputRow < 2 Then outputRow = 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Standardized" Then
            rowsImported = ImportSheet(ws, wsOut, outputRow)
            If rowsImported > 0 Then
                outputRow = outputRow + rowsImported
                summary = summary & ws.Name & ": " & rowsImported & " dong" & vbCrLf
            Else
                skipped = skipped & ws.Name & " (khong nhan dien duoc format hoac khong co du lieu)" & vbCrLf
            End If
        End If
    Next ws

    wsOut.Columns.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Dim msg As String
    msg = "HOAN TAT CHUAN HOA DU LIEU" & vbCrLf & vbCrLf
    If Len(summary) > 0 Then msg = msg & "Da import:" & vbCrLf & summary
    If Len(skipped) > 0 Then msg = msg & vbCrLf & "Bo qua:" & vbCrLf & skipped
    MsgBox msg, vbInformation
End Sub

'============================================================
' Tao / lam sach sheet dich "Standardized"
' Cau truc = giong Format 4, them cot Currency sau Entity,
' va them cot OEM sau Currency
'============================================================
Function PrepareOutputSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sales_standardized")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Sales_Standardized"
    Else
        ws.Cells.Clear
    End If

    Dim headers As Variant
    headers = Array("No.", "Name", "Address", "Country", "Product Group", "Item No.", _
                     "Description", "Category", "Invoice No.", "Location", "Invoiced Qty.", _
                     "Base UOM", "Sales Amount", "Discount Amount", "Net Sales", "Unit Price", _
                     "Unit Cost", "Total Cost", "GM", "Plan", "Weight", "Total Weight", _
                     "Entity", "Currency", "OEM", "Month", "Group", _
                     "Detected Format", "Source Sheet", "Header Row")

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(217, 217, 217)

    Set PrepareOutputSheet = ws
End Function

'============================================================
' BANG TRA CUU CURRENCY MAC DINH THEO ENTITY
' (dung khi source KHONG co san cot Currency - Format 1, 2, 4)
' >>> CHINH SUA O DAY neu sai <<<
'============================================================
Function GetDefaultCurrency(ByVal entity As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    d("DOAU") = "USD"   ' Dan-D Australia - gia dinh AUD, chinh lai neu sai
    d("CAN") = "CAD"    ' Canada - Amount LCY = CAD
    d("DO") = "VND"        ' Dominican - chua chac chan, de trong, ban tu dien vao
    d("MP3") = "VND"
    d("ADF") = "VND"
    If d.Exists(Trim(entity)) Then
        GetDefaultCurrency = d(Trim(entity))
    Else
        GetDefaultCurrency = ""
    End If
End Function

'============================================================
' NHAN DIEN CURRENCY THEO TEN SHEET
' Neu ten sheet co chua "CAD" -> tra "CAD"
' Neu ten sheet co chua "USD" -> tra "USD"
' Neu khong co ca hai -> tra "" (se fallback sang logic khac)
' Uu tien CAD/USD xuat hien ro trong ten, vd: "Sales(CAD)", "Sales(USD)", "CAN_USD"...
'============================================================
Function GetCurrencyFromSheetName(ByVal sheetName As String) As String
    Dim s As String
    s = UCase(sheetName)
    If InStr(s, "USD") > 0 Then
        GetCurrencyFromSheetName = "USD"
    ElseIf InStr(s, "CAD") > 0 Then
        GetCurrencyFromSheetName = "CAD"
    Else
        GetCurrencyFromSheetName = ""
    End If
End Function

'============================================================
' TU DONG DO DONG HEADER: quet tu dong 1 den dong 15,
' thu nhan dien format o tung dong, dong nao khop thi lay lam header.
' Tra ve so dong header (0 neu khong tim duoc), dong thoi tra ve
' headerMap va fmt qua tham chieu ByRef.
'============================================================
Function FindHeaderRow(ws As Worksheet, ByRef headerMap As Object, ByRef fmt As Integer) As Long
    Dim scanLimit As Long
    On Error Resume Next
    scanLimit = ws.UsedRange.Rows.Count
    On Error GoTo 0
    If scanLimit > 15 Then scanLimit = 15
    If scanLimit < 1 Then
        FindHeaderRow = 0
        fmt = 0
        Exit Function
    End If

    Dim r As Long, c As Long, lastColScan As Long
    Dim testMap As Object, testFmt As Integer, key As String

    For r = 1 To scanLimit
        Set testMap = CreateObject("Scripting.Dictionary")
        lastColScan = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastColScan >= 2 Then
            For c = 1 To lastColScan
                key = NormalizeHeader(ws.Cells(r, c).Value)
                If Len(key) > 0 And Not testMap.Exists(key) Then testMap.Add key, c
            Next c
            testFmt = DetectFormat(testMap)
            If testFmt > 0 Then
                Set headerMap = testMap
                fmt = testFmt
                FindHeaderRow = r
                Exit Function
            End If
        End If
    Next r

    FindHeaderRow = 0
    fmt = 0
End Function

'============================================================
' Doc 1 sheet nguon, tu do header, nhan dien format, map va ghi ra wsOut
'============================================================
Function ImportSheet(wsSrc As Worksheet, wsOut As Worksheet, ByVal startRow As Long) As Long
    Dim headerMap As Object
    Dim fmt As Integer
    Dim headerRow As Long

    headerRow = FindHeaderRow(wsSrc, headerMap, fmt)
    If headerRow = 0 Or fmt = 0 Then
        ImportSheet = 0
        Exit Function
    End If

    Dim lastRow As Long
    On Error Resume Next
    lastRow = wsSrc.UsedRange.Row + wsSrc.UsedRange.Rows.Count - 1
    On Error GoTo 0
    If lastRow <= headerRow Then
        ImportSheet = 0
        Exit Function
    End If

    ' Cot nguon can doc, tuy format
    Dim colNo As Long, colName As Long, colAddress As Long, colCountry As Long
    Dim colProductGroup As Long, colItemNo As Long, colDesc As Long, colCategory As Long
    Dim colInvoiceNo As Long, colLocation As Long, colQty As Long, colUOM As Long
    Dim colSalesAmount As Long, colDiscount As Long, colNetSales As Long, colUnitPrice As Long
    Dim colUnitCost As Long, colTotalCost As Long, colGM As Long, colPlan As Long
    Dim colWeight As Long, colTotalWeight As Long, colEntity As Long, colCurrency As Long
    Dim colMonth As Long, colGroup As Long, colDate As Long, colGP As Long
    Dim colOEM As Long

    Select Case fmt
        Case 1 ' Format 1: QuickBooks/DOAU style
            colEntity = GetCol(headerMap, Array("entity"))
            colDate = GetCol(headerMap, Array("date"))
            colName = GetCol(headerMap, Array("name"))
            colItemNo = GetCol(headerMap, Array("itemcode"))
            colDesc = GetCol(headerMap, Array("itemdescription"))
            colInvoiceNo = GetCol(headerMap, Array("num"))
            colQty = GetCol(headerMap, Array("qty"))
            colUOM = GetCol(headerMap, Array("um"))
            colSalesAmount = GetCol(headerMap, Array("amount"))
            colUnitPrice = GetCol(headerMap, Array("salesprice"))
            ' --- B? SUNG: t?m c?t Total Cost và Unit Cost cho DOAU ---
            colTotalCost = GetCol(headerMap, Array("totalcost", "total", "total cost"))
            colUnitCost = GetCol(headerMap, Array("unitcost", "unit cost", "cost"))
            ' Các c?t khác không có
            colAddress = 0: colCountry = 0: colProductGroup = 0: colCategory = 0
            colLocation = 0: colDiscount = 0: colGM = 0: colPlan = 0: colWeight = 0: colTotalWeight = 0
            colCurrency = 0: colMonth = 0: colGroup = 0: colGP = 0: colNo = 0

        Case 2 ' Format 2: CAN sales - Amount LCY, khong co Currency (thuong header o dong 3)
            colEntity = GetCol(headerMap, Array("entity"))
            colMonth = GetCol(headerMap, Array("month"))
            colName = GetCol(headerMap, Array("customername"))
            colLocation = GetCol(headerMap, Array("prov"))
            colItemNo = GetCol(headerMap, Array("itemno"))
            colDesc = GetCol(headerMap, Array("description"))
            colUOM = GetCol(headerMap, Array("packsize"))
            colQty = GetCol(headerMap, Array("invqty"))
            colSalesAmount = GetCol(headerMap, Array("sumofamountlcy", "amountlcy"))
            colGP = GetCol(headerMap, Array("sumofgpdollar"))
            colGM = GetCol(headerMap, Array("gppct"))
            colNo = 0: colAddress = 0: colCountry = 0: colProductGroup = 0: colCategory = 0
            colInvoiceNo = 0: colDiscount = 0: colUnitPrice = 0: colUnitCost = 0: colTotalCost = 0
            colPlan = 0: colWeight = 0: colTotalWeight = 0: colCurrency = 0: colGroup = 0

        Case 3 ' Format 3: CAN sales - Amount (CAD/USD) + Currency that (thuong header o dong 5 cho USD)
            colEntity = GetCol(headerMap, Array("entity"))
            colMonth = GetCol(headerMap, Array("month"))
            colName = GetCol(headerMap, Array("customername"))
            colLocation = GetCol(headerMap, Array("prov"))
            colItemNo = GetCol(headerMap, Array("itemno"))
            colDesc = GetCol(headerMap, Array("description"))
            colUOM = GetCol(headerMap, Array("packsize"))
            colQty = GetCol(headerMap, Array("invqty"))
            colCurrency = GetCol(headerMap, Array("currency"))
            colSalesAmount = FindColContains(headerMap, "amount")
            colGP = FindColContains(headerMap, "gpdollar")
            colGM = GetCol(headerMap, Array("gppct"))
            colCategory = GetCol(headerMap, Array("itemcatdescr"))
            colGroup = GetCol(headerMap, Array("itemcatdescr")) ' dung chung gia tri voi Category
            colNo = 0: colAddress = 0: colCountry = 0: colProductGroup = 0
            colInvoiceNo = 0: colDiscount = 0: colUnitPrice = 0: colUnitCost = 0: colTotalCost = 0
            colPlan = 0: colWeight = 0: colTotalWeight = 0

        Case 4 ' Format 4: DO chi tiet - map 1:1
            colNo = GetCol(headerMap, Array("no"))
            colName = GetCol(headerMap, Array("name"))
            colAddress = GetCol(headerMap, Array("address"))
            colCountry = GetCol(headerMap, Array("country"))
            colProductGroup = GetCol(headerMap, Array("productgroup"))
            colItemNo = GetCol(headerMap, Array("itemno"))
            colDesc = GetCol(headerMap, Array("description"))
            colCategory = GetCol(headerMap, Array("category"))
            colInvoiceNo = GetCol(headerMap, Array("invoiceno"))
            colLocation = GetCol(headerMap, Array("location"))
            colQty = GetCol(headerMap, Array("invoicedqty"))
            colUOM = GetCol(headerMap, Array("baseuom"))
            colSalesAmount = GetCol(headerMap, Array("salesamount"))
            colDiscount = GetCol(headerMap, Array("discountamount"))
            colNetSales = GetCol(headerMap, Array("netsales"))
            colUnitPrice = GetCol(headerMap, Array("unitprice"))
            colUnitCost = GetCol(headerMap, Array("unitcost"))
            colTotalCost = GetCol(headerMap, Array("totalcost"))
            colGM = GetCol(headerMap, Array("gm"))
            colPlan = GetCol(headerMap, Array("plan"))
            colWeight = GetCol(headerMap, Array("weight"))
            colTotalWeight = GetCol(headerMap, Array("totalweight"))
            colEntity = GetCol(headerMap, Array("entity"))
            colMonth = GetCol(headerMap, Array("month"))
            colGroup = GetCol(headerMap, Array("group"))
            colCurrency = 0: colGP = 0
    End Select

    ' T?m c?t OEM chung cho m?i format (n?u có trong header ngu?n)
    colOEM = GetCol(headerMap, Array("oem", "oemcode", "oem_name", "oem no", "oemcode"))

    Dim r As Long, outR As Long
    outR = startRow
    Dim lastMonth As String, lastName As String
    lastMonth = "": lastName = ""

    Dim vNo As Variant, vName As String, vAddress As String, vCountry As String
    Dim vProductGroup As String, vItemNo As String, vDesc As String, vCategory As String
    Dim vInvoiceNo As Variant, vLocation As String, vQty As Variant, vUOM As String
    Dim vSalesAmount As Variant, vDiscount As Variant, vNetSales As Variant, vUnitPrice As Variant
    Dim vUnitCost As Variant, vTotalCost As Variant, vGM As Variant, vPlan As Variant
    Dim vWeight As Variant, vTotalWeight As Variant, vEntity As String, vCurrency As String
    Dim vMonth As String, vGroup As String, vGP As Variant
    Dim vOEM As Variant

    For r = headerRow + 1 To lastRow
        If Application.WorksheetFunction.CountA(wsSrc.Rows(r)) = 0 Then GoTo NextRow

        vEntity = CStr(GetVal(wsSrc, r, colEntity))

        If fmt = 1 Then
            vMonth = DateToYYYYMM(GetVal(wsSrc, r, colDate))
        Else
            vMonth = CStr(GetVal(wsSrc, r, colMonth))
            If Len(Trim(vMonth)) = 0 Then
                vMonth = lastMonth
            Else
                lastMonth = vMonth
            End If
        End If

        vName = CStr(GetVal(wsSrc, r, colName))
        If Len(Trim(vName)) = 0 Then
            vName = lastName
        Else
            lastName = vName
        End If

        vNo = GetVal(wsSrc, r, colNo)
        vAddress = CStr(GetVal(wsSrc, r, colAddress))
        vCountry = CStr(GetVal(wsSrc, r, colCountry))
        vProductGroup = CStr(GetVal(wsSrc, r, colProductGroup))
        vItemNo = CStr(GetVal(wsSrc, r, colItemNo))
        vDesc = CStr(GetVal(wsSrc, r, colDesc))
        vCategory = CStr(GetVal(wsSrc, r, colCategory))
        vInvoiceNo = GetVal(wsSrc, r, colInvoiceNo)
        vLocation = CStr(GetVal(wsSrc, r, colLocation))
        vQty = GetVal(wsSrc, r, colQty)
        vUOM = CStr(GetVal(wsSrc, r, colUOM))
        vSalesAmount = GetVal(wsSrc, r, colSalesAmount)
        vDiscount = GetVal(wsSrc, r, colDiscount)
        vUnitPrice = GetVal(wsSrc, r, colUnitPrice)
        vUnitCost = GetVal(wsSrc, r, colUnitCost)
        vTotalCost = GetVal(wsSrc, r, colTotalCost)   ' Ð? có giá tr? cho Format 1
        vGM = GetVal(wsSrc, r, colGM)
        vPlan = GetVal(wsSrc, r, colPlan)
        vWeight = GetVal(wsSrc, r, colWeight)
        vTotalWeight = GetVal(wsSrc, r, colTotalWeight)
        vGroup = CStr(GetVal(wsSrc, r, colGroup))
        vGP = GetVal(wsSrc, r, colGP)

        ' L?y giá tr? OEM n?u có c?t
        vOEM = GetVal(wsSrc, r, colOEM)

        ' Net Sales: neu source co san thi lay, khong thi = Sales Amount
        If fmt = 4 Then
            vNetSales = GetVal(wsSrc, r, colNetSales)
        Else
            vNetSales = vSalesAmount
        End If

        ' Total Cost: neu chua co san nhung co Sales Amount va GP$ thi tu tinh
        If (fmt = 2 Or fmt = 3) And IsNumeric(vSalesAmount) And IsNumeric(vGP) Then
            vTotalCost = CDbl(vSalesAmount) - CDbl(vGP)
        End If

        ' Currency: uu tien 1) ten sheet co CAD/USD, 2) cot Currency that (Format 3), 3) bang mac dinh theo Entity
        Dim vCurrencyFromSheet As String
        vCurrencyFromSheet = GetCurrencyFromSheetName(wsSrc.Name)

        If Len(vCurrencyFromSheet) > 0 Then
            vCurrency = vCurrencyFromSheet
        ElseIf fmt = 3 And Len(Trim(CStr(vCurrency))) > 0 Then
            ' giu nguyen vCurrency da doc duoc tu cot Currency
        Else
            vCurrency = GetDefaultCurrency(vEntity)
        End If

        ' Bo qua dong khong co du lieu Qty va Sales Amount (dong trang/subtotal)
        If Len(Trim(CStr(vSalesAmount))) = 0 And Len(Trim(CStr(vQty))) = 0 Then GoTo NextRow

        With wsOut
            .Cells(outR, 1).Value = vNo
            .Cells(outR, 2).Value = vName
            .Cells(outR, 3).Value = vAddress
            .Cells(outR, 4).Value = vCountry
            .Cells(outR, 5).Value = vProductGroup
            .Cells(outR, 6).Value = vItemNo
            .Cells(outR, 7).Value = vDesc
            .Cells(outR, 8).Value = vCategory
            .Cells(outR, 9).Value = vInvoiceNo
            .Cells(outR, 10).Value = vLocation
            .Cells(outR, 11).Value = vQty
            .Cells(outR, 12).Value = vUOM
            .Cells(outR, 13).Value = vSalesAmount
            .Cells(outR, 14).Value = vDiscount
            .Cells(outR, 15).Value = vNetSales
            .Cells(outR, 16).Value = vUnitPrice
            .Cells(outR, 17).Value = vUnitCost
            .Cells(outR, 18).Value = vTotalCost      ' Total Cost - c?t R (c?t 18)
            .Cells(outR, 19).Value = vGM
            .Cells(outR, 20).Value = vPlan
            .Cells(outR, 21).Value = vWeight
            .Cells(outR, 22).Value = vTotalWeight
            .Cells(outR, 23).Value = vEntity
            .Cells(outR, 24).Value = vCurrency
            .Cells(outR, 25).Value = vOEM
            .Cells(outR, 26).Value = vMonth
            .Cells(outR, 27).Value = vGroup
            .Cells(outR, 28).Value = "Format " & fmt
            .Cells(outR, 29).Value = wsSrc.Name
            .Cells(outR, 30).Value = headerRow
        End With

        outR = outR + 1
NextRow:
    Next r

    ImportSheet = outR - startRow
End Function

'============================================================
' Nhan dien format dua tren header da chuan hoa
'============================================================
Function DetectFormat(headerMap As Object) As Integer
    Dim score1 As Integer, score2 As Integer, score3 As Integer, score4 As Integer

    score1 = 0
    If headerMap.Exists("itemcode") Then score1 = score1 + 2
    If headerMap.Exists("itemdescription") Then score1 = score1 + 2
    If headerMap.Exists("balance") Then score1 = score1 + 1
    If headerMap.Exists("memo") Then score1 = score1 + 1
    If headerMap.Exists("num") Then score1 = score1 + 1

    score2 = 0
    If headerMap.Exists("itemno") Then score2 = score2 + 1
    If headerMap.Exists("sumofamountlcy") Or headerMap.Exists("amountlcy") Then score2 = score2 + 3
    If headerMap.Exists("sumofgpdollar") Then score2 = score2 + 1
    If headerMap.Exists("prov") Then score2 = score2 + 1
    If headerMap.Exists("currency") Then score2 = score2 - 5 ' Format 2 KHONG co cot Currency

    score3 = 0
    If headerMap.Exists("itemno") Then score3 = score3 + 1
    If headerMap.Exists("currency") Then score3 = score3 + 3
    If headerMap.Exists("itemcatdescr") Then score3 = score3 + 1
    If headerMap.Exists("prov") Then score3 = score3 + 1

    score4 = 0
    If headerMap.Exists("invoicedqty") Then score4 = score4 + 2
    If headerMap.Exists("baseuom") Then score4 = score4 + 2
    If headerMap.Exists("totalcost") Then score4 = score4 + 1
    If headerMap.Exists("productgroup") Then score4 = score4 + 1
    If headerMap.Exists("gm") Then score4 = score4 + 1

    Dim best As Integer, bestScore As Integer
    best = 0: bestScore = 2 ' nguong toi thieu de tranh nhan nham dong khac
    If score1 > bestScore Then best = 1: bestScore = score1
    If score2 > bestScore Then best = 2: bestScore = score2
    If score3 > bestScore Then best = 3: bestScore = score3
    If score4 > bestScore Then best = 4: bestScore = score4

    DetectFormat = best
End Function

Function NormalizeHeader(ByVal s As Variant) As String
    Dim txt As String, i As Long, ch As String, res As String
    txt = LCase(Trim(CStr(s)))
    txt = Replace(txt, "$", "dollar")
    txt = Replace(txt, "%", "pct")
    res = ""
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch Like "[a-z0-9]" Then res = res & ch
    Next i
    NormalizeHeader = res
End Function

Function GetCol(headerMap As Object, candidates As Variant) As Long
    Dim v As Variant
    For Each v In candidates
        If headerMap.Exists(v) Then
            GetCol = headerMap(v)
            Exit Function
        End If
    Next v
    GetCol = 0
End Function

Function FindColContains(headerMap As Object, ByVal sub_ As String) As Long
    Dim k As Variant
    For Each k In headerMap.Keys
        If InStr(1, CStr(k), sub_, vbTextCompare) > 0 Then
            FindColContains = headerMap(k)
            Exit Function
        End If
    Next k
    FindColContains = 0
End Function

Function GetVal(ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then
        GetVal = ""
    Else
        GetVal = ws.Cells(r, c).Value
    End If
End Function

Function DateToYYYYMM(ByVal raw As Variant) As String
    Dim d As Date
    On Error GoTo Fallback
    If IsDate(raw) Then
        d = CDate(raw)
    Else
        Dim s As String, parts() As String
        s = CStr(raw)
        parts = Split(s, "/")
        If UBound(parts) = 2 Then
            d = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
        Else
            GoTo Fallback
        End If
    End If
    DateToYYYYMM = Format(d, "yyyy-mm")
    Exit Function
Fallback:
    DateToYYYYMM = CStr(raw)
End Function



