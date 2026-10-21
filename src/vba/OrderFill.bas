Attribute VB_Name = "OrderFill"
Option Explicit

'==================================================
' 起單回填
' 把上手回覆的起單信件貼到「起單回填」工作表，
' 按「填入」回填 Term!A2:T2 與 Term!O5:O9；按「清空」清掉貼上區。
'
' 找值方式：依表頭名稱（不依欄位位置），支援三種排法
'   1. 表格：一列表頭、下一列數值（大部分上手）
'   2. 左右一組：左格標題、右格數值（例：HSBC 的日期區）
'   3. 同一格「標題: 數值」（例：Initial price: Closing price）
'==================================================

Private Const SHEET_NAME As String = "起單回填"
Private Const PASTE_TOP As Long = 7          ' 貼上區從第 7 列開始
Private Const CELL_ISSUER_PICK As String = "C3"
Private Const CELL_ISSUER_RESULT As String = "C4"

Private Const PRI_TABLE As Long = 1
Private Const PRI_PAIR As Long = 2
Private Const PRI_INLINE As Long = 3

Private mHeaderMap As Object   ' 正規化表頭 → 欄位代號
Private mVal As Object         ' 欄位代號 → 原始值
Private mPri As Object         ' 欄位代號 → 來源優先序

'==================================================
' 按鈕：清空
'==================================================
Public Sub OrderFill_Clear()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    With ws.Range(ws.Cells(PASTE_TOP, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count))
        .UnMerge
        .Clear
    End With
    ws.Range(CELL_ISSUER_PICK & "," & CELL_ISSUER_RESULT).ClearContents

    ' 從信件貼上的圖片一併刪除（按鈕在貼上區上方，不受影響）
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.TopLeftCell.Row >= PASTE_TOP Then shp.Delete
    Next i

    ws.Activate
    ws.Cells(PASTE_TOP, 1).Select
End Sub

'==================================================
' 按鈕：填入
'==================================================
Public Sub OrderFill_Fill()
    Dim report As String
    OrderFill_Run False, report
End Sub

'==================================================
' 主流程；silent = True 時不跳訊息（測試用），結果放在 report
'==================================================
Public Function OrderFill_Run(ByVal silent As Boolean, ByRef report As String) As Boolean
    Dim ws As Worksheet, wsT As Worksheet
    Dim data As Variant
    Dim issuer As String, issuerHow As String
    Dim warn As String
    Dim prevEvents As Boolean

    Dim product As String, tdcc As String, ccy As String
    Dim guar As Variant, strike As Variant, koBarrier As Variant
    Dim coupon As Variant, notePrice As Variant, tenor As Variant
    Dim ki As Variant, pgRate As Variant
    Dim koType As String, barrierType As String, initPrice As String
    Dim ul(1 To 5) As String
    Dim tradeDate As Date, issueDate As Date, finalDate As Date, maturityDate As Date
    Dim hasTrade As Boolean, hasIssue As Boolean, hasFinal As Boolean, hasMaturity As Boolean
    Dim i As Long

    report = ""
    prevEvents = Application.EnableEvents
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set wsT = ThisWorkbook.Worksheets("Term")

    '--- 讀貼上區 ---
    data = ReadPasteArea(ws)
    If IsEmpty(data) Then
        report = "貼上區沒有資料，請先把上手的起單信件內容貼在第 " & PASTE_TOP & " 列以下。"
        If Not silent Then MsgBox report, vbExclamation, "起單回填"
        Exit Function
    End If

    CollectFields data

    '--- TDCC（信裡沒有就留空）---
    tdcc = NormTDCC(GetVal("TDCC"))

    '--- 判斷上手 ---
    issuer = Trim$(CStr(ws.Range(CELL_ISSUER_PICK).Value))
    If issuer <> "" Then
        issuerHow = "下拉選單"
    Else
        If tdcc <> "" Then issuer = IssuerFromTDCC(tdcc)
        If issuer <> "" Then
            issuerHow = "TDCC 前 3 碼"
        Else
            issuer = IssuerFromSignature(data)
            If issuer <> "" Then issuerHow = "表頭特徵"
        End If
    End If

    If issuer = "" Then
        ws.Range(CELL_ISSUER_RESULT).Value = "判斷不出上手"
        report = "判斷不出是哪一家上手，Term 沒有變更。" & vbCrLf & _
                 "請在 " & CELL_ISSUER_PICK & " 的下拉選單選擇上手後，再按一次「填入」。"
        If Not silent Then MsgBox report, vbExclamation, "起單回填"
        Exit Function
    End If
    ws.Range(CELL_ISSUER_RESULT).Value = issuer & "（" & issuerHow & "）"

    '--- 整理各欄位 ---
    product = NormProduct(GetVal("PRODUCT"))
    If product = "" Then
        product = "FCN"
        warn = AddWarn(warn, "信中找不到 Product，A2 預設填 FCN")
    End If

    ccy = UCase$(Trim$(CStr(GetVal("CURRENCY"))))

    For i = 1 To 5
        ul(i) = UCase$(Trim$(CStr(GetVal("UL" & i))))
    Next i

    guar = NumField("GUAR", 0)
    ' Citi 的 Non Callable Periods 比我們的 Guaranteed Period 少 1 個月
    If issuer = "Citi" And Not IsEmpty(guar) Then guar = guar + 1

    strike = NumField("STRIKE", 3)
    koBarrier = NumField("KOBARRIER", 3)
    coupon = NumField("COUPON", 1)
    ki = NumField("KI", 3)
    pgRate = NumField("PGRATE", 0)
    tenor = NumField("TENOR", 0)

    If issuer = "Citi" Then
        ' Citi 只給 Upfront（手續費），Note Price = 100 - Upfront
        notePrice = NumField("PRICE", 0)
        If Not IsEmpty(notePrice) Then notePrice = Round(100 - notePrice, 4)
    Else
        notePrice = NumField("PRICE", 3)
    End If

    koType = NormKOType(GetVal("KOTYPE"), GetVal("KOMEMORY"), GetVal("MEMAUTOCALL"), GetVal("DAILYKO"))
    barrierType = NormBarrierType(GetVal("BARRIERTYPE"))
    If barrierType = "NONE" Then ki = Empty

    initPrice = NormInitPrice(GetVal("INITPRICE"))

    hasTrade = DateField("TRADE", tradeDate)
    hasIssue = DateField("ISSUE", issueDate)
    hasFinal = DateField("FINAL", finalDate)
    hasMaturity = DateField("MATURITY", maturityDate)

    '--- 提醒：找不到的欄位 ---
    If tdcc = "" Then warn = AddWarn(warn, "信中找不到 TDCC，B2 留空")
    If ccy = "" Then warn = AddWarn(warn, "信中找不到 Currency")
    If IsEmpty(guar) Then warn = AddWarn(warn, "信中找不到 Guaranteed Period / Non-call")
    If ul(1) = "" Then warn = AddWarn(warn, "信中找不到標的")
    If IsEmpty(strike) Then warn = AddWarn(warn, "信中找不到 Strike")
    If koType = "" Then warn = AddWarn(warn, "信中找不到 KO Type")
    If IsEmpty(koBarrier) Then warn = AddWarn(warn, "信中找不到 KO Barrier")
    If IsEmpty(coupon) Then warn = AddWarn(warn, "信中找不到 Coupon")
    If IsEmpty(notePrice) Then warn = AddWarn(warn, "信中找不到 Note Price")
    If IsEmpty(tenor) Then warn = AddWarn(warn, "信中找不到 Tenor")
    If barrierType = "" Then warn = AddWarn(warn, "信中找不到 Barrier Type")
    If barrierType <> "" And barrierType <> "NONE" And IsEmpty(ki) Then warn = AddWarn(warn, "信中找不到 KI")
    If IsEmpty(pgRate) Then warn = AddWarn(warn, "信中找不到 PG rate")
    If Not hasTrade Then
        tradeDate = Date
        warn = AddWarn(warn, "信中找不到 Trade date，O5 填今天")
    End If
    If Not hasIssue Then warn = AddWarn(warn, "信中找不到 Issue date，O6 由系統推算")
    If Not hasFinal Then warn = AddWarn(warn, "信中找不到 Final valuation date，O7 由系統推算")
    If Not hasMaturity Then warn = AddWarn(warn, "信中找不到 Maturity date，O8 由系統推算")
    If initPrice = "" Then warn = AddWarn(warn, "信中找不到期初價格，O9 留空")

    '--- 寫入 Term ---
    On Error GoTo Fail
    Application.EnableEvents = False

    wsT.Range("A2:T2").ClearContents
    wsT.Range("O5:O9").ClearContents

    wsT.Range("A2").Value = product
    If tdcc <> "" Then wsT.Range("B2").Value = "'" & tdcc
    wsT.Range("C2").Value = issuer
    wsT.Range("D2").Value = ccy
    wsT.Range("E2").Value = guar
    For i = 1 To 5
        If ul(i) <> "" Then wsT.Cells(2, 5 + i).Value = ul(i)   ' F2:J2
    Next i
    wsT.Range("K2").Value = strike
    wsT.Range("L2").Value = koType
    wsT.Range("M2").Value = koBarrier
    wsT.Range("N2").Value = coupon
    wsT.Range("O2").Value = notePrice
    wsT.Range("P2").Value = tenor
    wsT.Range("Q2").Value = "M"
    wsT.Range("R2").Value = barrierType
    wsT.Range("S2").Value = ki
    wsT.Range("T2").Value = pgRate

    wsT.Range("O5").Value = Format$(tradeDate, "yyyy/mm/dd")
    wsT.Range("O9").Value = initPrice

    ' 讓 Term 既有的事件更新「產品名稱」的幣別與期間，並依 O5 推算日期；
    ' 信裡有的日期接著再覆蓋回去
    wsT.Activate
    Application.EnableEvents = True
    wsT.Range("D2").Value = wsT.Range("D2").Value
    wsT.Range("P2").Value = wsT.Range("P2").Value
    Application.EnableEvents = False

    If hasIssue Then wsT.Range("O6").Value = Format$(issueDate, "yyyy/mm/dd")
    If hasFinal Then wsT.Range("O7").Value = Format$(finalDate, "yyyy/mm/dd")
    If hasMaturity Then wsT.Range("O8").Value = Format$(maturityDate, "yyyy/mm/dd")

    Application.EnableEvents = prevEvents
    wsT.Range("A2").Select

    report = "已回填 Term。" & vbCrLf & "上手：" & issuer & "（" & issuerHow & "）"
    If warn <> "" Then report = report & vbCrLf & vbCrLf & "請留意：" & vbCrLf & warn
    If Not silent Then MsgBox report, IIf(warn = "", vbInformation, vbExclamation), "起單回填"
    OrderFill_Run = True
    Exit Function

Fail:
    Application.EnableEvents = prevEvents
    report = "回填時發生錯誤：" & Err.Description
    If Not silent Then MsgBox report, vbCritical, "起單回填"
End Function

'==================================================
' 建立「起單回填」工作表與按鈕（匯入模組後執行一次）
'==================================================
Public Sub OrderFill_Setup()
    OrderFill_SetupCore False
End Sub

' silent = True 時不跳訊息（自動化用）
Public Sub OrderFill_SetupCore(ByVal silent As Boolean)
    Dim ws As Worksheet
    Dim wsV As Worksheet
    Dim lastV As Long
    Dim btn As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Term"))
        ws.Name = SHEET_NAME
    Else
        If Not silent Then
            If MsgBox("「" & SHEET_NAME & "」已存在，要重新建立按鈕與版面嗎？（貼上區內容會保留）", _
                      vbYesNo + vbQuestion, "起單回填") = vbNo Then Exit Sub
        End If
        Do While ws.Buttons.Count > 0
            ws.Buttons(1).Delete
        Loop
    End If

    ws.Range("A1").Value = "上手起單信件回填 Term"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "上手（自動判斷失敗時才選）"
    ws.Range("A4").Value = "判斷結果"
    ws.Range("A6").Value = "↓ 請把上手的起單信件內容貼在第 " & PASTE_TOP & " 列以下（建議貼在 A" & PASTE_TOP & "）"
    ws.Range("A6").Font.Color = RGB(192, 0, 0)
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 22

    Set wsV = ThisWorkbook.Worksheets("Variable")
    lastV = wsV.Cells(wsV.Rows.Count, "B").End(xlUp).Row
    With ws.Range(CELL_ISSUER_PICK).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=Variable!$B$2:$B$" & lastV
    End With
    ws.Range(CELL_ISSUER_PICK).Interior.Color = RGB(255, 242, 204)

    Set btn = ws.Buttons.Add(ws.Range("E1").Left, ws.Range("E1").Top, 90, 30)
    btn.Caption = "清空"
    btn.OnAction = "OrderFill_Clear"

    Set btn = ws.Buttons.Add(ws.Range("E1").Left + 100, ws.Range("E1").Top, 90, 30)
    btn.Caption = "填入"
    btn.OnAction = "OrderFill_Fill"

    ws.Activate
    ws.Cells(PASTE_TOP, 1).Select
    If Not silent Then MsgBox "「" & SHEET_NAME & "」已建立。", vbInformation, "起單回填"
End Sub

'==================================================
' 讀取貼上區為 2D 陣列；沒有資料回傳 Empty
'==================================================
Private Function ReadPasteArea(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long, lastCol As Long
    Dim ur As Range
    Dim arr As Variant

    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1
    If lastRow < PASTE_TOP Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(PASTE_TOP, 1), ws.Cells(lastRow, lastCol))) = 0 Then Exit Function

    If lastRow = PASTE_TOP And lastCol = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = ws.Cells(PASTE_TOP, 1).Value
    Else
        arr = ws.Range(ws.Cells(PASTE_TOP, 1), ws.Cells(lastRow, lastCol)).Value
    End If
    ReadPasteArea = arr
End Function

'==================================================
' 掃描貼上區，把找得到的欄位放進 mVal / mPri
'==================================================
Private Sub CollectFields(ByRef data As Variant)
    Dim nR As Long, nC As Long
    Dim r As Long, c As Long, r2 As Long
    Dim isHeader() As Boolean
    Dim cnt As Long, valueRow As Long
    Dim key As String, s As String, p As Long

    BuildHeaderMap
    Set mVal = CreateObject("Scripting.Dictionary")
    Set mPri = CreateObject("Scripting.Dictionary")

    nR = UBound(data, 1)
    nC = UBound(data, 2)
    ReDim isHeader(1 To nR)

    '--- 1. 表格：一列有 3 個以上認得的表頭，就把下一列當數值 ---
    For r = 1 To nR
        cnt = 0
        For c = 1 To nC
            If mHeaderMap.Exists(CellKey(data(r, c))) Then cnt = cnt + 1
        Next c

        If cnt >= 3 Then
            isHeader(r) = True
            valueRow = 0
            For r2 = r + 1 To IIf(r + 3 > nR, nR, r + 3)
                For c = 1 To nC
                    If mHeaderMap.Exists(CellKey(data(r, c))) Then
                        If Not IsBlankVal(data(r2, c)) Then valueRow = r2
                    End If
                    If valueRow > 0 Then Exit For
                Next c
                If valueRow > 0 Then Exit For
            Next r2

            If valueRow > 0 Then
                For c = 1 To nC
                    key = CellKey(data(r, c))
                    If mHeaderMap.Exists(key) Then RecordField mHeaderMap(key), data(valueRow, c), PRI_TABLE
                Next c
            End If
        End If
    Next r

    '--- 2. 左右一組、3.「標題: 數值」 ---
    For r = 1 To nR
        If Not isHeader(r) Then
            For c = 1 To nC
                If VarType(data(r, c)) = vbString Then
                    s = Trim$(data(r, c))
                    key = NormKey(s)
                    If mHeaderMap.Exists(key) And c < nC Then
                        If Not IsBlankVal(data(r, c + 1)) Then RecordField mHeaderMap(key), data(r, c + 1), PRI_PAIR
                    End If

                    p = InStr(s, ":")
                    If p > 1 And p < 40 Then
                        key = NormKey(Left$(s, p - 1))
                        If mHeaderMap.Exists(key) Then RecordField mHeaderMap(key), Trim$(Mid$(s, p + 1)), PRI_INLINE
                    End If
                End If
            Next c
        End If
    Next r
End Sub

' 依欄位型別檢查後記錄；同一欄位保留優先序較高（數字較小）、較早出現的值
Private Sub RecordField(ByVal field As String, ByVal v As Variant, ByVal pri As Long)
    Dim d As Date, n As Double, pct As Boolean

    If IsBlankVal(v) Then Exit Sub

    Select Case field
        Case "TRADE", "ISSUE", "FINAL", "MATURITY"
            If Not ParseDateVal(v, d) Then Exit Sub
        Case "GUAR", "STRIKE", "KOBARRIER", "COUPON", "PRICE", "TENOR", "KI", "PGRATE"
            If Not ParseNumber(v, n, pct) Then Exit Sub
    End Select

    If mPri.Exists(field) Then
        If mPri(field) <= pri Then Exit Sub
    End If
    mVal(field) = v
    mPri(field) = pri
End Sub

Private Function GetVal(ByVal field As String) As Variant
    If mVal.Exists(field) Then GetVal = mVal(field) Else GetVal = ""
End Function

' 數值欄位：pctLimit > 0 時，沒有 % 符號且 0 < 值 <= pctLimit 視為小數表示（0.8 → 80）
Private Function NumField(ByVal field As String, ByVal pctLimit As Double) As Variant
    Dim n As Double, pct As Boolean

    NumField = Empty
    If Not mVal.Exists(field) Then Exit Function
    If Not ParseNumber(mVal(field), n, pct) Then Exit Function

    If pctLimit > 0 And Not pct Then
        If field = "COUPON" Then
            If n > 0 And n < pctLimit Then n = n * 100
        ElseIf n > 0 And n <= pctLimit Then
            n = n * 100
        End If
    End If
    NumField = Round(n, 4)
End Function

Private Function DateField(ByVal field As String, ByRef d As Date) As Boolean
    If mVal.Exists(field) Then DateField = ParseDateVal(mVal(field), d)
End Function

'==================================================
' 表頭對照表：正規化後的表頭 → 欄位代號
'==================================================
Private Sub BuildHeaderMap()
    Dim i As Long

    If Not mHeaderMap Is Nothing Then Exit Sub
    Set mHeaderMap = CreateObject("Scripting.Dictionary")

    AddSyn "PRODUCT", "product"
    AddSyn "CURRENCY", "currency", "curr"
    AddSyn "GUAR", "guaranteedperiodsm", "guaranteedperiods", "noncallperiodsm", _
                   "noncallableperiods", "noncallm"
    For i = 1 To 5
        AddSyn "UL" & i, "bbgcode" & i, "stock" & i, "underlying" & i
    Next i
    AddSyn "STRIKE", "strike", "putstrike"
    AddSyn "KOTYPE", "kotype"
    AddSyn "KOBARRIER", "kobarrier", "ko", "autocall"
    AddSyn "COUPON", "couponpa", "coupon"
    AddSyn "PRICE", "upfrontnoteprice", "noteprice", "cost", "reoffer", "offerprice", "upfront"
    AddSyn "TENOR", "tenorm", "tenor", "noofperiods"
    AddSyn "BARRIERTYPE", "barriertype", "putki?", "kitype"
    AddSyn "KI", "kibarrier", "putki", "ki"
    AddSyn "KOMEMORY", "komemory"
    AddSyn "MEMAUTOCALL", "memoryautocall"
    AddSyn "DAILYKO", "dailyko"
    AddSyn "TRADE", "tradedate", "strikedate", "initialvaluationdate"
    AddSyn "ISSUE", "issuedate", "issuancedate", "issuedatetcalendardays"
    AddSyn "FINAL", "finalvaluationdate", "finalfixingdate", "redemptionvaluationdate", "finalvaluation"
    AddSyn "MATURITY", "maturitydate", "redemptiondate", "finalmaturitydate"
    AddSyn "TDCC", "tdcc", "tdcccode"
    AddSyn "INITPRICE", "initialprice"
    AddSyn "PGRATE", "pgrate"
End Sub

Private Sub AddSyn(ByVal field As String, ParamArray keys() As Variant)
    Dim k As Variant
    For Each k In keys
        mHeaderMap(CStr(k)) = field
    Next k
End Sub

' 正規化：轉小寫，只留英數字與「?」（用來分辨 MS 的 "Put KI?" 與 "Put KI (%)"）
Private Function NormKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "?" Then out = out & ch
    Next i
    NormKey = out
End Function

Private Function CellKey(ByVal v As Variant) As String
    If VarType(v) = vbString Then CellKey = NormKey(v)
End Function

Private Function IsBlankVal(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        IsBlankVal = True
    Else
        IsBlankVal = (Trim$(Replace(CStr(v), ChrW(160), " ")) = "" Or Trim$(CStr(v)) = "-")
    End If
End Function

'==================================================
' 判斷上手
'==================================================
Private Function IssuerFromTDCC(ByVal tdcc As String) As String
    Dim wsV As Worksheet
    Dim lastRow As Long, r As Long
    Dim a As String

    Set wsV = ThisWorkbook.Worksheets("Variable")
    lastRow = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        a = Trim$(CStr(wsV.Cells(r, "A").Value))
        If IsNumeric(a) And Len(a) < 3 Then a = Format$(CLng(a), "000")
        If a = Left$(tdcc, 3) Then
            IssuerFromTDCC = Trim$(CStr(wsV.Cells(r, "B").Value))
            Exit Function
        End If
    Next r
End Function

' 依各家起單信件的表頭特徵判斷；順序有意義（較獨特的特徵先比）
Private Function IssuerFromSignature(ByRef data As Variant) As String
    Dim keys As Object
    Dim allText As String
    Dim r As Long, c As Long

    Set keys = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            If VarType(data(r, c)) = vbString Then
                keys(NormKey(data(r, c))) = True
                allText = allText & "|" & UCase$(data(r, c))
            End If
        Next c
    Next r

    If keys.Exists("gsquoteid") Or InStr(allText, "HKGEQV") > 0 Then
        IssuerFromSignature = "GS"
    ElseIf keys.Exists("msid") Then
        IssuerFromSignature = "MS"
    ElseIf keys.Exists("nomuraid") Then
        IssuerFromSignature = "Nomura"
    ElseIf InStr(allText, "UBS ID") > 0 Then
        IssuerFromSignature = "UBS"
    ElseIf keys.Exists("priceid") Or keys.Exists("refspot1") Or keys.Exists("swapindex") Then
        IssuerFromSignature = "Citi"
    ElseIf keys.Exists("datescheduleconfig") Or keys.Exists("reoffer") Or keys.Exists("wrapper") Then
        IssuerFromSignature = "CA"
    ElseIf keys.Exists("settlementfrequency") Or keys.Exists("offerprice") Or InStr(allText, "AUTOPRICER") > 0 Then
        IssuerFromSignature = "SG"
    ElseIf keys.Exists("issuerprodref") Or keys.Exists("issuedatelag") Then
        IssuerFromSignature = "DBS"
    ElseIf keys.Exists("clientref") Or keys.Exists("redemptionvaluationdate") Or InStr(allText, "ESS_TW_") > 0 Then
        IssuerFromSignature = "BNP"
    ElseIf keys.Exists("systemremarks") Then
        IssuerFromSignature = "JPM"
    ElseIf keys.Exists("message") Then
        IssuerFromSignature = "Natixis"
    ElseIf keys.Exists("comments") And keys.Exists("id") Then
        IssuerFromSignature = "HSBC"
    ElseIf keys.Exists("quoteid") And keys.Exists("notional") Then
        IssuerFromSignature = "BARC"
    ElseIf keys.Exists("id") And keys.Exists("remarks") Then
        IssuerFromSignature = "BBVA"
    End If
End Function

'==================================================
' 數值與文字正規化
'==================================================
' 支援：數字、" 20.00 "、"96.20%→96.15%"（取最後值）、"1m"
Private Function ParseNumber(ByVal v As Variant, ByRef n As Double, ByRef hadPct As Boolean) As Boolean
    Dim s As String, p As Long

    hadPct = False
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            n = CDbl(v)
            ParseNumber = True
        Case vbString
            s = Trim$(Replace(v, ChrW(160), " "))
            p = InStrRev(s, ChrW(&H2192))   ' →
            If p > 0 Then s = Trim$(Mid$(s, p + 1))
            hadPct = (InStr(s, "%") > 0)
            s = Replace(Replace(Replace(s, "%", ""), ",", ""), " ", "")
            If LCase$(Right$(s, 1)) = "m" Then s = Left$(s, Len(s) - 1)
            If s = "" Or s Like "*[!0-9.-]*" Then Exit Function
            n = Val(s)
            ParseNumber = True
    End Select
End Function

' 支援：日期值、yyyy-mm-dd、yyyy/mm/dd、dd/mm/yyyy（HSBC），後面可帶時間
Private Function ParseDateVal(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim s As String, p() As String
    Dim y As Long, m As Long, dd As Long

    If VarType(v) = vbDate Then
        d = DateSerial(Year(v), Month(v), Day(v))
        ParseDateVal = True
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function

    s = Trim$(v)
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    s = Replace(Replace(s, "-", "/"), ".", "/")
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function

    If Len(p(0)) = 4 Then
        y = CLng(p(0)): m = CLng(p(1)): dd = CLng(p(2))
    ElseIf Len(p(2)) = 4 Then
        dd = CLng(p(0)): m = CLng(p(1)): y = CLng(p(2))
    Else
        Exit Function
    End If
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function

    d = DateSerial(y, m, dd)
    ParseDateVal = (Month(d) = m And Day(d) = dd)
End Function

' TDCC：只留數字；Excel 吃掉開頭 0 時補回 12 碼；不是 12 碼就當作沒有
Private Function NormTDCC(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, out As String

    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbDecimal Then
        s = Format$(v, "0")
    Else
        s = CStr(v)
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    If Len(out) = 11 Then out = "0" & out
    If Len(out) = 12 Then NormTDCC = out
End Function

' Product 對應到 Variable!D 的寫法；Citi 的 FCA 即 FCN
Private Function NormProduct(ByVal v As Variant) As String
    Dim s As String
    Dim wsV As Worksheet
    Dim lastRow As Long, r As Long

    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    If UCase$(s) = "FCA" Then s = "FCN"

    Set wsV = ThisWorkbook.Worksheets("Variable")
    lastRow = wsV.Cells(wsV.Rows.Count, "D").End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(wsV.Cells(r, "D").Value)), s, vbTextCompare) = 0 Then
            NormProduct = Trim$(CStr(wsV.Cells(r, "D").Value))
            Exit Function
        End If
    Next r
    NormProduct = s
End Function

' KO Type 對應到 Variable!C：Daily Memory / Daily / Period End / Period End Memory
' MS：KO Type = OBSERVATION 搭配 KO Memory；Citi：沒有 KO Type，由 Memory Autocall / Daily KO 推得
Private Function NormKOType(ByVal koType As Variant, ByVal koMemory As Variant, _
                            ByVal memAutocall As Variant, ByVal dailyKO As Variant) As String
    Dim s As String
    Dim isDaily As Boolean, isMemory As Boolean

    s = LCase$(Trim$(CStr(koType)))
    If s = "" Then
        If IsBlankVal(memAutocall) And IsBlankVal(dailyKO) Then Exit Function
        isDaily = IsYes(dailyKO)
        isMemory = IsYes(memAutocall)
    Else
        isDaily = (InStr(s, "daily") > 0)
        isMemory = (InStr(s, "memory") > 0) Or IsYes(koMemory)
        If Not isDaily And InStr(s, "period") = 0 And InStr(s, "observation") = 0 And Not isMemory Then
            NormKOType = Trim$(CStr(koType))
            Exit Function
        End If
    End If

    If isDaily Then
        NormKOType = IIf(isMemory, "Daily Memory", "Daily")
    Else
        NormKOType = IIf(isMemory, "Period End Memory", "Period End")
    End If
End Function

Private Function IsYes(ByVal v As Variant) As Boolean
    Dim s As String
    If VarType(v) = vbBoolean Then
        IsYes = v
    ElseIf Not IsError(v) Then
        s = UCase$(Trim$(CStr(v)))
        IsYes = (s = "TRUE" Or s = "YES" Or s = "Y")
    End If
End Function

' Barrier Type 對應到 AKI / MKI / EKI / NONE
Private Function NormBarrierType(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    Select Case s
        Case "AKI", "MKI", "EKI", "NONE"
            NormBarrierType = s
        Case "EUROPEAN"
            NormBarrierType = "EKI"
        Case "AMERICAN"
            NormBarrierType = "AKI"
        Case "N/A", "NA", "NO", "NIL"
            NormBarrierType = "NONE"
        Case ""
            NormBarrierType = ""
        Case Else
            NormBarrierType = s
    End Select
End Function

Private Function NormInitPrice(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(CStr(v))
    If InStr(s, "open") > 0 Then
        NormInitPrice = "開盤價"
    ElseIf InStr(s, "clos") > 0 Then
        NormInitPrice = "收盤價"
    End If
End Function

Private Function AddWarn(ByVal baseText As String, ByVal newText As String) As String
    If baseText = "" Then
        AddWarn = "- " & newText
    Else
        AddWarn = baseText & vbCrLf & "- " & newText
    End If
End Function
