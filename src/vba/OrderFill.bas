Attribute VB_Name = "OrderFill"
Option Explicit

'==================================================
' 起單回填
' 把上手回覆的起單信件貼到「起單回填」工作表，
' 按「填入」回填 Term!A2:T2、Term!V2 與 Term!O5:O9；按「清空」清掉貼上區。
' 保本型商品（目前只有 Range Accrual）要先在 E5 選商品才填。
'
' 找值方式：依表頭名稱（不依欄位位置），支援三種排法
'   1. 表格：一列表頭、下一列數值（大部分上手）
'   2. 左右一組：左格標題、右格數值（例：HSBC 的日期區）
'   3. 同一格「標題: 數值」（例：Initial price: Closing price；GS 的 [ Issue Date: 20Oct26 ]）
'
' 對照規則與 structured-product-workflow 讀 (Order) 信的規則一致
' （該 repo commit f0bcce0 之前的 src/sp_workflow/order_formats.py、order_template.py）：
' 不猜值；讀不到、認不得、兩處不一致的欄位留空，並在摘要提醒。
'==================================================

Private Const SHEET_NAME As String = "起單回填"
Private Const PASTE_TOP As Long = 7          ' 貼上區從第 7 列開始
Private Const CELL_ISSUER_PICK As String = "C3"
Private Const CELL_ISSUER_RESULT As String = "C4"
Private Const CELL_TDCC_INPUT As String = "C5"
Private Const CELL_PROTECTED_PICK As String = "E5"
' E5 下拉選單的保本型商品（逗號分隔）；新增保本型商品時加在這裡，並在 OrderFill_Run 補上它的規則
Private Const PROTECTED_PRODUCTS As String = "Range Accrual"

Private Const PRI_TABLE As Long = 1
Private Const PRI_PAIR As Long = 2
Private Const PRI_INLINE As Long = 3

Private mHeaderMap As Object   ' 正規化表頭 → 欄位代號
Private mVal As Object         ' 欄位代號 → 值（改價已取最後的值）
Private mKey As Object         ' 欄位代號 → 比較用的值
Private mSrc As Object         ' 欄位代號 → 讀到的值與位置
Private mConflict As Object    ' 欄位代號 → 兩處以上不一致時的各處值
Private mFlag As Object        ' 欄位代號 → 已經提醒過（格式不認得、不一致、認不得）
Private mNotes As String       ' 摘要的提醒

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
    ws.Range(CELL_ISSUER_PICK & "," & CELL_ISSUER_RESULT & "," & CELL_TDCC_INPUT & "," & CELL_PROTECTED_PICK).ClearContents

    ' 從信件貼上的圖片一併刪除（按鈕在貼上區上方，不受影響）
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        ' 表單控制項不動：選過 C3 後 Excel 會加一個隱藏的下拉框，它沒有 TopLeftCell，讀取會出現 1004
        If shp.Type <> msoFormControl Then
            If shp.TopLeftCell.Row >= PASTE_TOP Then shp.Delete
        End If
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
    Dim prevEvents As Boolean

    Dim product As String, tdcc As String, mailTdcc As String, manualTdcc As String, ccy As String
    Dim guar As Variant, strike As Variant, koBarrier As Variant
    Dim coupon As Variant, notePrice As Variant, upfront As Variant
    Dim tenor As Variant, periods As Variant
    Dim ki As Variant, pgRate As Variant, accrualBarrier As Variant, expressMult As Variant, benKO As Variant
    Dim floorCpn As Variant
    Dim koType As String, barrierType As String, initPrice As String, fixedCpn As String, protPick As String
    Dim guarField As String, strikeField As String, koField As String, couponField As String
    Dim isSuperBull As Boolean, isBEN As Boolean, isRA As Boolean
    Dim ul(1 To 5) As String
    Dim tradeDate As Date, issueDate As Date, finalDate As Date, maturityDate As Date
    Dim hasTrade As Boolean, hasIssue As Boolean, hasFinal As Boolean, hasMaturity As Boolean
    Dim i As Long

    report = ""
    prevEvents = Application.EnableEvents
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set wsT = ThisWorkbook.Worksheets("Term")

    '--- 手填 TDCC（C5，文字格式）：有值時以它為準；不是剛好 12 碼數字就整個不填 ---
    If Not IsBlankVal(ws.Range(CELL_TDCC_INPUT).Value) Then
        manualTdcc = Trim$(CStr(ws.Range(CELL_TDCC_INPUT).Value))
        If Not (manualTdcc Like "############") Then
            report = CELL_TDCC_INPUT & " 的 TDCC 要是 12 碼數字，Term 沒有變更。"
            If Not silent Then MsgBox report, vbExclamation, "起單回填"
            Exit Function
        End If
    End If

    '--- 讀貼上區 ---
    data = ReadPasteArea(ws)
    If IsEmpty(data) Then
        report = "貼上區沒有資料，請先把上手的起單信件內容貼在第 " & PASTE_TOP & " 列以下。"
        If Not silent Then MsgBox report, vbExclamation, "起單回填"
        Exit Function
    End If

    CollectFields data

    '--- TDCC：C5 → 信裡的 TDCC ---
    mailTdcc = NormTDCC(GetVal("TDCC"))
    If manualTdcc <> "" Then
        tdcc = manualTdcc
        If mailTdcc <> "" And mailTdcc <> manualTdcc Then
            AddNote CELL_TDCC_INPUT & " 的 TDCC 和信裡的 " & mailTdcc & " 不同，B2 以 " & CELL_TDCC_INPUT & " 為準"
        End If
    Else
        tdcc = mailTdcc
    End If

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

    ' MS 的 DACN 寫成 Range Accrual：有 Strike（Put Strike）就是 DACN
    If issuer = "MS" And product = "Range Accrual" And HasField("STRIKE") Then
        product = "DACN"
        AddNote "MS 的 Product 是 Range Accrual，但有 Strike，判定為 DACN，A2 請確認"
    End If

    ' 保本型商品：E5 有選時 A2 = 選的商品；沒選但信看起來是 Range Accrual 時不填
    protPick = Trim$(CStr(ws.Range(CELL_PROTECTED_PICK).Value))
    If protPick <> "" Then
        If InStr("," & PROTECTED_PRODUCTS & ",", "," & protPick & ",") = 0 Then
            report = CELL_PROTECTED_PICK & " 的保本型商品「" & protPick & "」認不得，Term 沒有變更。" & vbCrLf & _
                     "請從下拉選單選擇。"
        ElseIf IsFlagged("PRODUCT") Then
            report = CELL_PROTECTED_PICK & " 選的是 " & protPick & "，但信裡的 Product（" & GetVal("PRODUCT") & _
                     "）認不得，Term 沒有變更。" & vbCrLf & "請確認商品；不是保本型商品時請清空 " & CELL_PROTECTED_PICK & "。"
        ElseIf product <> "" And product <> protPick Then
            report = CELL_PROTECTED_PICK & " 選的是 " & protPick & "，但這封信的商品是 " & product & _
                     "，Term 沒有變更。" & vbCrLf & "請確認商品；不是保本型商品時請清空 " & CELL_PROTECTED_PICK & "。"
        ElseIf protPick = "Range Accrual" And Not LooksLikeRA(product) Then
            ' 例：E5 沒清掉，又貼了一封沒有 Product 的 SG FCN
            report = CELL_PROTECTED_PICK & " 選的是 Range Accrual，但信裡沒有 Range Accrual 的欄位（Product、" & _
                     "Coupon frequency=Callable frequency、DRA coupon、RA Rate），Term 沒有變更。" & vbCrLf & _
                     "請確認商品；不是保本型商品時請清空 " & CELL_PROTECTED_PICK & "。"
        End If
        If report = "" Then product = protPick
    ElseIf LooksLikeRA(product) Then
        report = "這封信看起來是保本型商品（Range Accrual），Term 沒有變更。" & vbCrLf & _
                 "請在 " & CELL_PROTECTED_PICK & " 的「保本型商品」下拉選單選擇商品後，再按一次「填入」。"
    End If
    If report <> "" Then
        If Not silent Then MsgBox report, vbExclamation, "起單回填"
        Exit Function
    End If
    isRA = (product = "Range Accrual")

    ccy = UCase$(Trim$(CStr(GetVal("CURRENCY"))))

    For i = 1 To 5
        ul(i) = UCase$(Trim$(CStr(GetVal("UL" & i))))
    Next i
    If Not IsBlankVal(GetVal("UL6")) Then
        AddNote "標的超過 5 檔（第 6 檔：" & GetVal("UL6") & "），Term 只放得下 5 檔"
    End If

    ' Super Bull 的 E／K／M／N 讀各自的表頭：HSBC、CA 是 Strike、Autocall Barrier；
    ' 其他上手（BNP）的 K、M 都是 Super Bull Level
    guarField = "GUAR"
    strikeField = "STRIKE"
    koField = "KOBARRIER"
    couponField = "COUPON"
    isSuperBull = (product = "Super Bull")
    If isSuperBull Then
        guarField = "ACFREQ"
        couponField = "FLATCPN"
        If issuer = "HSBC" Or issuer = "CA" Then
            koField = "ACBARRIER"
        Else
            strikeField = "SBLEVEL"
            koField = "SBLEVEL"
        End If
    End If

    ' BEN 的 N 是 Bonus Coupon（flat）；M 是 Coupon Barrier，信裡沒有（HSBC、CA、JPM、Natixis）時 = Strike
    isBEN = (product = "Regular BEN" Or product = "Digital BEN")
    If isBEN Then
        couponField = "FLATCPN"
        If mSrc.Exists("CPNBARRIER") Or IsFlagged("CPNBARRIER") Then koField = "CPNBARRIER" Else koField = strikeField
    End If

    If Not isBEN Then   ' BEN 的 E2 = Tenor，在下面另外算
        guar = NumField(guarField, 0)
        ' Citi 的 Non Callable Periods 比我們的 Guaranteed Period 少 1 個月（0 → 1）
        If issuer = "Citi" And Not isSuperBull And Not IsEmpty(guar) Then guar = guar + 1
    End If

    strike = NumField(strikeField, 3)
    If koField = strikeField Then
        koBarrier = strike
    Else
        koBarrier = NumField(koField, 3)
    End If
    If isBEN And koField = "CPNBARRIER" And Not IsEmpty(koBarrier) And Not IsEmpty(strike) Then
        If koBarrier <> strike Then
            AddNote "Coupon Barrier（" & koBarrier & "）和 Strike（" & strike & "）不同，M2 填 Coupon Barrier，請確認"
        End If
    End If
    coupon = NumField(couponField, 1)
    ki = NumField("KI", 3)
    pgRate = NumField("PGRATE", 0)

    ' 信中沒有 Product 時不預設；只有 SG 的 AutoCall 是單一數字、有 Put Strike 時，
    ' 依 Fixed Coupons 推定：All Periods（每期固定配息）→ FCN、First Period／First Two Periods（只有前一、兩期固定）→ DACN
    If product = "" And Not IsFlagged("PRODUCT") Then
        fixedCpn = LCase$(CollapseSpaces(CStr(GetVal("FIXEDCPN"))))
        If issuer = "SG" And (fixedCpn = "all periods" Or fixedCpn = "first period" Or fixedCpn = "first two periods") _
           And Not IsEmpty(koBarrier) And Not IsEmpty(strike) Then
            If fixedCpn = "all periods" Then product = "FCN" Else product = "DACN"
            AddNote "信中沒有 Product，依 SG 條件（Fixed Coupons 是 " & GetVal("FIXEDCPN") & "、AutoCall 是單一數字、" & _
                    "有 Put Strike）推定為 " & product & "，A2 請確認"
        Else
            WarnMissing "PRODUCT"
        End If
    End If

    ' Accrual Barrier（MS 的 DACN）沒有對應欄位；和 Strike 不同時提醒，K2 仍填 Strike
    accrualBarrier = NumField("ACCRUALBARRIER", 0)
    If Not IsEmpty(accrualBarrier) And Not IsEmpty(strike) Then
        If accrualBarrier <> strike Then
            AddNote "Accrual Barrier（" & accrualBarrier & "）和 Strike（" & strike & "）不同，K2 填 Strike"
        End If
    End If

    ' FCN_Express 從第一期就能 KO：Guaranteed Period 固定 0；倍數欄位沒有對應欄位，不是 1 時提醒
    If product = "FCN_Express" Then
        If Not IsEmpty(guar) Then
            If guar <> 0 Then AddNote "FCN_Express 的 Guaranteed Period 固定 0，信裡是 " & guar & "，E2 填 0"
        ElseIf IsFlagged("GUAR") Then
            AddNote "FCN_Express 的 Guaranteed Period 固定 0，E2 已填 0（不是留空）"
        End If
        guar = 0
        expressMult = NumField("EXPRESSMULT", 0)
        If Not IsEmpty(expressMult) Then
            If expressMult <> 1 Then
                AddNote FieldLabel("EXPRESSMULT") & " 是 " & expressMult & "，不是 1；Term 沒有這個欄位，請確認條件"
            End If
        End If
    End If

    tenor = NumField("TENOR", 0)
    ' 寫成年（BARC 的 RA 是 5Y）：Range Accrual 在下面另外讀，其他產品的 P2 是月，留空
    If IsEmpty(tenor) And mVal.Exists("TENOR") And Not isRA Then
        FlagField "TENOR", "Tenor 是 " & ShowVal(GetVal("TENOR")) & "（年），P2 要填月數，留空"
    End If
    If IsEmpty(tenor) And Not IsFlagged("TENOR") Then
        ' SG 沒有 Tenor：Settlement Frequency 是 Monthly 時 Tenor = No. of Periods
        periods = NumField("PERIODS", 0)
        If Not IsEmpty(periods) Then
            If LCase$(CollapseSpaces(CStr(GetVal("SETTLEFREQ")))) = "monthly" Then
                tenor = periods
            Else
                FlagField "TENOR", "No. of Periods 是 " & periods & "，但 Settlement Frequency 不是 Monthly（" & _
                                   ShowBlank(GetVal("SETTLEFREQ")) & "），P2 留空"
            End If
        End If
    End If

    ' BEN 依 Term!A14：Guaranteed Period = Tenor；KO Barrier 有值（BEN with KO，樣本沒有）時不猜，E2 留空
    If isBEN Then
        benKO = NumField("KOBARRIER", 3)
        If Not IsEmpty(benKO) Or IsFlagged("KOBARRIER") Then
            guar = Empty
            AddNote "信裡有 KO Barrier（" & ShowBlank(GetVal("KOBARRIER")) & "），可能是 BEN with KO，E2 留空，請依 TS 填入"
        Else
            guar = tenor
        End If
    End If

    notePrice = NumField("PRICE", 3)
    If IsEmpty(notePrice) And Not IsFlagged("PRICE") Then
        ' Citi 只給 Upfront（UF），Note Price = 100 - Upfront
        upfront = NumField("UPFRONT", 0)
        If Not IsEmpty(upfront) Then notePrice = Round(100 - upfront, 4)
    End If

    ' Super Bull、BEN 的 KO Type 固定 Period End
    If isSuperBull Or isBEN Then
        koType = "Period End"
    Else
        koType = NormKOType(GetVal("KOTYPE"), GetVal("KOMEMORY"), GetVal("MEMAUTOCALL"), GetVal("DAILYKO"))
    End If

    ' Super Bull 沒有 KI：R2、S2 留空
    If isSuperBull Then
        ki = Empty
    Else
        barrierType = NormBarrierType(GetVal("BARRIERTYPE"))
        If barrierType = "" And Not IsFlagged("BARRIERTYPE") Then
            If Not IsEmpty(ki) Then
                FlagField "BARRIERTYPE", "有 KI（" & ki & "）但沒有 Barrier Type，R2 留空"
            ElseIf ul(1) <> "" And Not IsFlagged("KI") Then
                barrierType = "NONE"   ' 例：SG 的 KI Type、KI 都是 N/A
            End If
        End If
        If barrierType = "NONE" Then
            If Not IsEmpty(ki) Then
                If ki <> 0 Then AddNote "Barrier Type 是 NONE，但 KI 有值（" & ki & "），S2 留空"
            End If
            ki = Empty
        End If
    End If

    initPrice = NormInitPrice(GetVal("INITPRICE"))

    ' Range Accrual 只填 A～F、O～Q、T、V：E2 是 Non-call、F2 是 CMS／CMT 10Y、P2 是年、V2 是 Floor
    If isRA Then
        guar = RANonCall()
        ul(1) = RAUnderlying(GetVal("RAUL"))
        For i = 2 To 5
            ul(i) = ""
        Next i
        strike = Empty
        koType = ""
        koBarrier = Empty
        coupon = Empty
        barrierType = ""
        ki = Empty
        initPrice = ""
        tenor = RATenor()
        ' 玉山格式只有 UF：Note Price = 100 - UF
        If IsEmpty(notePrice) And Not IsFlagged("PRICE") Then
            upfront = NumField("RAUF", 0)
            If Not IsEmpty(upfront) Then notePrice = Round(100 - upfront, 4)
        End If
        floorCpn = RAFloor(GetVal("RACOUPON"))
    End If

    hasTrade = DateField("TRADE", tradeDate)
    hasIssue = DateField("ISSUE", issueDate)
    hasFinal = DateField("FINAL", finalDate)
    hasMaturity = DateField("MATURITY", maturityDate)

    '--- 提醒：找不到的欄位（已提醒過格式或不一致的不重複）---
    If tdcc = "" Then WarnMissing "TDCC"
    If ccy = "" Then WarnMissing "CURRENCY"
    If IsEmpty(guar) And Not isBEN And Not IsFlagged("NCY") Then   ' BEN 的 E2 來自 Tenor，已另外提醒
        If isRA Then WarnMissing "CALLFEATURE" Else WarnMissing guarField
    End If
    If ul(1) = "" Then WarnMissing "UL1"
    If Not isRA Then   ' Range Accrual 的 K～N、R、S、O9 不填
        If IsEmpty(strike) Then WarnMissing strikeField
        If koType = "" Then WarnMissing "KOTYPE"
        If IsEmpty(koBarrier) And koField <> strikeField Then WarnMissing koField
        If IsEmpty(coupon) Then WarnMissing couponField
    End If
    If IsEmpty(notePrice) And Not IsFlagged("UPFRONT") And Not IsFlagged("RAUF") Then WarnMissing "PRICE"
    If IsEmpty(tenor) And Not IsFlagged("TENORY") Then WarnMissing "TENOR"
    If Not isRA Then
        If barrierType = "" And Not isSuperBull Then WarnMissing "BARRIERTYPE"
        If barrierType <> "" And barrierType <> "NONE" And IsEmpty(ki) Then WarnMissing "KI"
    End If
    If IsEmpty(pgRate) Then WarnMissing "PGRATE"
    If Not hasTrade Then
        tradeDate = Date
        WarnMissing "TRADE"
    End If
    If Not hasIssue Then WarnMissing "ISSUE"
    If Not hasFinal Then WarnMissing "FINAL"
    If Not hasMaturity Then WarnMissing "MATURITY"
    If initPrice = "" And Not isRA Then WarnMissing "INITPRICE"

    '--- 寫入 Term ---
    On Error GoTo Fail
    Application.EnableEvents = False

    wsT.Range("A2:T2").ClearContents
    wsT.Range("V2").ClearContents
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
    wsT.Range("Q2").Value = IIf(isRA, "Y", "M")
    wsT.Range("R2").Value = barrierType
    wsT.Range("S2").Value = ki
    wsT.Range("T2").Value = pgRate
    ' Floor Coupon：FCN、DACN、FCN_Express、Super Bull、BEN 固定 0；Range Accrual 是 Floor@ 或 0；其他產品留空
    If product = "FCN" Or product = "DACN" Or product = "FCN_Express" Or isSuperBull Or isBEN Then wsT.Range("V2").Value = 0
    If isRA Then wsT.Range("V2").Value = floorCpn

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
    If mNotes <> "" Then report = report & vbCrLf & vbCrLf & "請留意：" & vbCrLf & mNotes
    If Not silent Then MsgBox report, IIf(mNotes = "", vbInformation, vbExclamation), "起單回填"
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
    ws.Range("A5").Value = "TDCC（信裡沒有時手填）"
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

    ' 手填 TDCC 用文字格式，保留開頭的 0
    With ws.Range(CELL_TDCC_INPUT)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 242, 204)
    End With

    ' 保本型商品：標題在左邊一格
    ws.Range(CELL_PROTECTED_PICK).Offset(0, -1).Value = "保本型商品（保本型要先選）"
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 16
    With ws.Range(CELL_PROTECTED_PICK).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=PROTECTED_PRODUCTS
    End With
    ws.Range(CELL_PROTECTED_PICK).Interior.Color = RGB(255, 242, 204)

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
' 百分比格式的數字改成帶 % 的文字（0.8 → "80%"），和信裡看到的寫法一致
'==================================================
Private Function ReadPasteArea(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long, lastCol As Long
    Dim ur As Range
    Dim arr As Variant
    Dim r As Long, c As Long

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

    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            Select Case VarType(arr(r, c))
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                    If InStr(ws.Cells(PASTE_TOP + r - 1, c).NumberFormat, "%") > 0 Then
                        arr(r, c) = NumText(CDbl(arr(r, c)) * 100) & "%"
                    End If
            End Select
        Next c
    Next r
    ReadPasteArea = arr
End Function

' 數字轉成一般寫法的文字（小數點用「.」）：0.85 → "0.85"；Str$ 會寫成 " .85"、極小值會變成 1E-05
Private Function NumText(ByVal n As Double) As String
    Dim s As String
    s = Format$(Abs(n), "0.########")
    s = Replace(s, Mid$(Format$(0.5, "0.0"), 2, 1), ".")   ' Format$ 依地區設定的小數點
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    If n < 0 Then s = "-" & s
    NumText = s
End Function

'==================================================
' 掃描貼上區，把找到的欄位放進 mVal；兩處不一致的欄位留空並提醒
'==================================================
Private Sub CollectFields(ByRef data As Variant)
    Dim nR As Long, nC As Long
    Dim r As Long, c As Long, r2 As Long
    Dim isHeader() As Boolean, isValueRow() As Boolean
    Dim colField() As String
    Dim cnt As Long, valueRow As Long
    Dim key As String, s As String
    Dim segs As Variant, seg As Variant, k As Variant

    BuildHeaderMap
    Set mVal = CreateObject("Scripting.Dictionary")
    Set mKey = CreateObject("Scripting.Dictionary")
    Set mSrc = CreateObject("Scripting.Dictionary")
    Set mConflict = CreateObject("Scripting.Dictionary")
    Set mFlag = CreateObject("Scripting.Dictionary")
    mNotes = ""

    nR = UBound(data, 1)
    nC = UBound(data, 2)
    ReDim isHeader(1 To nR)
    ReDim isValueRow(1 To nR)

    '--- 1. 表格：一列有 3 個以上認得的表頭，就找下一列的數值 ---
    For r = 1 To nR
        cnt = 0
        For c = 1 To nC
            If mHeaderMap.Exists(CellKey(data(r, c))) Then cnt = cnt + 1
        Next c

        If cnt >= 3 Then
            isHeader(r) = True
            ReDim colField(1 To nC)
            For c = 1 To nC
                key = CellKey(data(r, c))
                If mHeaderMap.Exists(key) Then colField(c) = mHeaderMap(key)
            Next c

            ' 表頭下的說明列（例：玉山格式 Coupon 下的 Coupon frequency=Callable frequency）不是數值，
            ' 改用說明當該欄的表頭
            valueRow = 0
            For r2 = r + 1 To IIf(r + 3 > nR, nR, r + 3)
                If IsHeaderNoteRow(data, colField, r2) Then
                    For c = 1 To nC
                        key = CellKey(data(r2, c))
                        If colField(c) <> "" And mHeaderMap.Exists(key) Then colField(c) = mHeaderMap(key)
                    Next c
                Else
                    For c = 1 To nC
                        If colField(c) <> "" And Not IsBlankVal(data(r2, c)) Then
                            valueRow = r2
                            Exit For
                        End If
                    Next c
                    If valueRow > 0 Then Exit For
                End If
            Next r2

            If valueRow > 0 Then
                isValueRow(valueRow) = True
                For c = 1 To nC
                    If colField(c) = "RACOUPON" Then
                        RecordField colField(c), ColumnText(data, valueRow, c, isValueRow), PRI_TABLE
                    ElseIf colField(c) <> "" Then
                        RecordField colField(c), data(valueRow, c), PRI_TABLE
                    End If
                Next c
                ' 後面幾列也是數值（一張表多筆單）：只讀第一列，但這些列也不做左右一組
                r2 = valueRow + 1
                Do While IsExtraValueRow(data, r, r2)
                    isValueRow(r2) = True
                    r2 = r2 + 1
                Loop
                If r2 > valueRow + 1 Then
                    AddNote "表格有不只一列數值（一封信多筆單），只讀第一列，請確認是要填的那一筆"
                End If
            End If
        End If
    Next r

    '--- 2. 左右一組、3.「標題: 數值」（一格裡有好幾組 [ 標題: 數值 ] 時逐組拆開）---
    ' 表格的數值列不做左右一組：SG 的 Quote ? 欄值是 "Strike"，會把右邊的 Coupon 當成 Strike
    For r = 1 To nR
        If Not isHeader(r) Then
            For c = 1 To nC
                If VarType(data(r, c)) = vbString Then
                    s = Trim$(data(r, c))
                    key = NormKey(s)
                    If mHeaderMap.Exists(key) And c < nC And Not isValueRow(r) Then
                        If Not IsBlankVal(data(r, c + 1)) Then RecordField mHeaderMap(key), data(r, c + 1), PRI_PAIR
                    End If

                    If InStr(s, "[") > 0 Then
                        segs = Split(Replace(s, "]", "["), "[")
                    Else
                        segs = Array(s)
                    End If
                    For Each seg In segs
                        RecordInline Trim$(CStr(seg))
                    Next seg
                End If
            Next c
        End If
    Next r

    For Each k In mConflict.Keys
        mVal.Remove k
        FlagField CStr(k), FieldLabel(CStr(k)) & " 不一致：" & mConflict(k) & "，" & MissingAction(CStr(k))
    Next k
End Sub

' 「標題: 數值」或「標題：數值」
Private Sub RecordInline(ByVal s As String)
    Dim p As Long, p2 As Long, key As String

    p = InStr(s, ":")
    p2 = InStr(s, ChrW(&HFF1A))   ' 全形冒號
    If p2 > 0 And (p = 0 Or p2 < p) Then p = p2
    If p > 1 And p < 40 Then
        key = NormKey(Left$(s, p - 1))
        If mHeaderMap.Exists(key) Then RecordField mHeaderMap(key), Trim$(Mid$(s, p + 1)), PRI_INLINE
    End If
End Sub

' 依欄位型別檢查後記錄；同一欄出現好幾次時，值不同就記為不一致
Private Sub RecordField(ByVal field As String, ByVal v As Variant, ByVal pri As Long)
    Dim raw As Variant, key As String, revised As Boolean
    Dim d As Date, n As Double, pct As Boolean

    If IsBlankVal(v) Then Exit Sub
    raw = v
    If VarType(v) = vbString Then
        v = LatestText(CStr(v), revised)
        If IsBlankVal(v) Then Exit Sub
    End If

    Select Case field
        Case "TRADE", "ISSUE", "FINAL", "MATURITY"
            If Not ParseDateVal(v, d) Then
                FlagField field, FieldLabel(field) & " 格式不認得：" & ShowVal(raw) & "，" & MissingAction(field)
                Exit Sub
            End If
            key = "D" & CStr(CLng(d))
        Case "GUAR", "STRIKE", "KOBARRIER", "COUPON", "PRICE", "UPFRONT", "TENOR", "PERIODS", "KI", "PGRATE", _
             "ACCRUALBARRIER", "EXPRESSMULT", "SBLEVEL", "ACBARRIER", "ACFREQ", "FLATCPN", "CPNBARRIER", _
             "TENORY", "NCY", "RAUF"
            If ParseNumber(v, n, pct) Then
                key = "N" & Trim$(Str$(Round(n, 6)))
            ElseIf field = "TENOR" And ParseYears(v, n) Then
                key = "Y" & Trim$(Str$(Round(n, 6)))   ' Tenor 寫成年（BARC 的 RA 是 5Y）；和月數的值算不一致
            Else
                FlagField field, FieldLabel(field) & " 格式不認得：" & ShowVal(raw) & "，" & MissingAction(field)
                Exit Sub
            End If
        Case "TDCC"
            key = "C" & NormTDCC(v)
        Case Else
            key = "T" & LCase$(CollapseSpaces(CStr(v)))
    End Select
    If revised Then AddNote FieldLabel(field) & " 有修改：" & ShowVal(raw) & "，取最後的值"

    If mKey.Exists(field) Then
        If mKey(field) <> key Then
            If Not mConflict.Exists(field) Then mConflict(field) = mSrc(field)
            mConflict(field) = mConflict(field) & "／" & ShowVal(raw) & "（" & PriLabel(pri) & "）"
        End If
        Exit Sub
    End If
    mVal(field) = v
    mKey(field) = key
    mSrc(field) = ShowVal(raw) & "（" & PriLabel(pri) & "）"
End Sub

' 表頭下的說明列：表頭欄有字，而且都是認得的表頭
Private Function IsHeaderNoteRow(ByRef data As Variant, ByRef colField() As String, ByVal r As Long) As Boolean
    Dim c As Long, filled As Long

    For c = 1 To UBound(data, 2)
        If colField(c) <> "" And Not IsBlankVal(data(r, c)) Then
            If Not mHeaderMap.Exists(CellKey(data(r, c))) Then Exit Function
            filled = filled + 1
        End If
    Next c
    IsHeaderNoteRow = (filled > 0)
End Function

' 玉山格式的配息拆成好幾列（Y1 : 8% p.a. fixed／Y2+ : 6% x n/N p.a, Floor@ 2.5%）：接著讀同一欄下面的列，
' 到空白格或表頭為止；這些列不做左右一組
Private Function ColumnText(ByRef data As Variant, ByVal r As Long, ByVal c As Long, _
                            ByRef isValueRow() As Boolean) As String
    Dim s As String, r2 As Long

    s = CStr(data(r, c))
    For r2 = r + 1 To UBound(data, 1)
        If IsBlankVal(data(r2, c)) Or mHeaderMap.Exists(CellKey(data(r2, c))) Then Exit For
        s = s & " " & CStr(data(r2, c))
        isValueRow(r2) = True
    Next r2
    ColumnText = s
End Function

' 第 r 列在 3 個以上的表頭下有值（例：MS 一張表三筆單）；第 r 列本身是表頭（下一張表）時不算
Private Function IsExtraValueRow(ByRef data As Variant, ByVal headerRow As Long, ByVal r As Long) As Boolean
    Dim c As Long, filled As Long, headers As Long

    If r > UBound(data, 1) Then Exit Function
    For c = 1 To UBound(data, 2)
        If mHeaderMap.Exists(CellKey(data(r, c))) Then headers = headers + 1
        If mHeaderMap.Exists(CellKey(data(headerRow, c))) Then
            If Not IsBlankVal(data(r, c)) Then filled = filled + 1
        End If
    Next c
    IsExtraValueRow = (filled >= 3 And headers < 3)
End Function

Private Function GetVal(ByVal field As String) As Variant
    If mVal.Exists(field) Then GetVal = mVal(field) Else GetVal = ""
End Function

' 數值欄位：帶 % 的值是百分比數字（80% → 80）；沒有 % 時照填。
' decimalLimit > 0 且沒有 % 的值像小數（0 < 值 <= decimalLimit；Coupon 是 < 1）時提醒確認
Private Function NumField(ByVal field As String, ByVal decimalLimit As Double) As Variant
    Dim n As Double, pct As Boolean, looksDecimal As Boolean

    NumField = Empty
    If Not mVal.Exists(field) Then Exit Function
    If Not ParseNumber(mVal(field), n, pct) Then Exit Function

    If decimalLimit > 0 And Not pct And n > 0 Then
        If field = "COUPON" Or field = "FLATCPN" Then
            looksDecimal = (n < decimalLimit)
        Else
            looksDecimal = (n <= decimalLimit)
        End If
        If looksDecimal Then
            AddNote FieldLabel(field) & " 是 " & n & "、沒有 %，照填；請確認是不是 " & Round(n * 100, 4) & "%"
        End If
    End If
    NumField = Round(n, 4)
End Function

Private Function DateField(ByVal field As String, ByRef d As Date) As Boolean
    If mVal.Exists(field) Then DateField = ParseDateVal(mVal(field), d)
End Function

'==================================================
' 提醒
'==================================================
Private Sub AddNote(ByVal text As String)
    mNotes = AddWarn(mNotes, text)
End Sub

' 提醒並記下這欄已提醒過，之後不再重複提醒「找不到」
Private Sub FlagField(ByVal field As String, ByVal text As String)
    mFlag(field) = True
    AddNote text
End Sub

Private Function IsFlagged(ByVal field As String) As Boolean
    IsFlagged = mFlag.Exists(field)
End Function

' 信裡有這個欄位（包含格式不認得、不一致而留空的）
Private Function HasField(ByVal field As String) As Boolean
    HasField = mSrc.Exists(field) Or IsFlagged(field)
End Function

Private Sub WarnMissing(ByVal field As String)
    If Not IsFlagged(field) Then AddNote "信中找不到 " & FieldLabel(field) & "，" & MissingAction(field)
End Sub

' 讀不到時 Term 怎麼處理
Private Function MissingAction(ByVal field As String) As String
    Select Case field
        Case "PRODUCT": MissingAction = "A2 留空"
        Case "TDCC": MissingAction = "B2 留空（可在 " & CELL_TDCC_INPUT & " 手填）"
        Case "TRADE": MissingAction = "O5 填今天"
        Case "ISSUE": MissingAction = "O6 由系統推算"
        Case "FINAL": MissingAction = "O7 由系統推算"
        Case "MATURITY": MissingAction = "O8 由系統推算"
        Case "INITPRICE": MissingAction = "O9 留空"
        Case Else: MissingAction = "留空"
    End Select
End Function

Private Function FieldLabel(ByVal field As String) As String
    Select Case field
        Case "PRODUCT": FieldLabel = "Product"
        Case "CURRENCY": FieldLabel = "Currency"
        Case "GUAR": FieldLabel = "Guaranteed Period / Non-call"
        Case "UL1", "UL2", "UL3", "UL4", "UL5", "UL6": FieldLabel = "標的 " & Mid$(field, 3)
        Case "STRIKE": FieldLabel = "Strike"
        Case "KOTYPE": FieldLabel = "KO Type"
        Case "KOBARRIER": FieldLabel = "KO Barrier"
        Case "COUPON": FieldLabel = "Coupon"
        Case "PRICE": FieldLabel = "Note Price"
        Case "UPFRONT": FieldLabel = "Upfront"
        Case "TENOR": FieldLabel = "Tenor"
        Case "PERIODS": FieldLabel = "No. of Periods"
        Case "SETTLEFREQ": FieldLabel = "Settlement Frequency"
        Case "FIXEDCPN": FieldLabel = "Fixed Coupons"
        Case "BARRIERTYPE": FieldLabel = "Barrier Type"
        Case "KI": FieldLabel = "KI"
        Case "KOMEMORY": FieldLabel = "KO Memory"
        Case "MEMAUTOCALL": FieldLabel = "Memory Autocall"
        Case "DAILYKO": FieldLabel = "Daily KO"
        Case "TRADE": FieldLabel = "Trade date"
        Case "ISSUE": FieldLabel = "Issue date"
        Case "FINAL": FieldLabel = "Final valuation date"
        Case "MATURITY": FieldLabel = "Maturity date"
        Case "TDCC": FieldLabel = "TDCC"
        Case "INITPRICE": FieldLabel = "期初價格"
        Case "PGRATE": FieldLabel = "PG rate"
        Case "ACCRUALBARRIER": FieldLabel = "Accrual Barrier"
        Case "EXPRESSMULT": FieldLabel = "First Period Rebate Multiplier／Express Multiplier"
        Case "SBLEVEL": FieldLabel = "Super Bull Level"
        Case "ACBARRIER": FieldLabel = "Autocall Barrier"
        Case "ACFREQ": FieldLabel = "Autocall Frequency／Autocall Freq／Autocall Date(one-time)"
        Case "FLATCPN": FieldLabel = "Autocall Coupon Flat／Autocall Flat Coupon／Bonus Coupon／Coupon Flat"
        Case "CPNBARRIER": FieldLabel = "Coupon Barrier"
        Case "TENORY": FieldLabel = "Tenor (Y)"
        Case "NCY": FieldLabel = "Non call period (Y)"
        Case "CALLFEATURE": FieldLabel = "Non call period／Call Feature／Autocall Feature／NC"
        Case "RAUF": FieldLabel = "UF"
        Case "ISCALLABLE": FieldLabel = "Is Callable"
        Case "RAUL": FieldLabel = "Underlying／Index／Condition_1"
        Case "RACOUPON": FieldLabel = "Coupon（玉山格式）"
        Case "DRACOUPON": FieldLabel = "DRA coupon"
        Case "RARATE": FieldLabel = "RA Rate"
        Case Else: FieldLabel = field
    End Select
End Function

Private Function PriLabel(ByVal pri As Long) As String
    Select Case pri
        Case PRI_TABLE: PriLabel = "表格"
        Case PRI_PAIR: PriLabel = "左右一組"
        Case Else: PriLabel = "標題: 數值"
    End Select
End Function

Private Function ShowVal(ByVal v As Variant) As String
    If VarType(v) = vbDate Then
        ShowVal = Format$(v, "yyyy/mm/dd")
    ElseIf IsError(v) Or IsEmpty(v) Then
        ShowVal = ""
    Else
        ShowVal = CStr(v)
    End If
End Function

Private Function ShowBlank(ByVal v As Variant) As String
    If IsBlankVal(v) Then ShowBlank = "（空白）" Else ShowBlank = ShowVal(v)
End Function

'==================================================
' 表頭對照表：正規化後的表頭 → 欄位代號
'==================================================
Private Sub BuildHeaderMap()
    Dim i As Long

    If Not mHeaderMap Is Nothing Then Exit Sub
    Set mHeaderMap = CreateObject("Scripting.Dictionary")

    AddSyn "PRODUCT", "product"
    AddSyn "CURRENCY", "currency", "curr", "ccy"
    AddSyn "GUAR", "guaranteedperiodsm", "guaranteedperiods", "noncallperiodsm", _
                   "noncallableperiods", "noncallm"
    For i = 1 To 6
        AddSyn "UL" & i, "bbgcode" & i, "stock" & i, "underlying" & i
    Next i
    AddSyn "STRIKE", "strike", "putstrike"
    AddSyn "KOTYPE", "kotype"
    AddSyn "KOBARRIER", "kobarrier", "ko", "autocall"
    AddSyn "COUPON", "couponpa", "coupon"
    AddSyn "PRICE", "upfrontnoteprice", "noteprice", "cost", "reoffer", "offerprice", "issueprice"
    AddSyn "UPFRONT", "upfront"                  ' Citi 的 Upfront (%) 是 UF，不是票面價格
    AddSyn "TENOR", "tenorm", "tenor", "tenormonths"
    AddSyn "PERIODS", "noofperiods"              ' SG
    AddSyn "SETTLEFREQ", "settlementfrequency"   ' SG
    AddSyn "FIXEDCPN", "fixedcoupons"            ' SG，用來推定 Product
    AddSyn "BARRIERTYPE", "barriertype", "putki?", "kitype"
    AddSyn "KI", "kibarrier", "putki", "ki"
    AddSyn "KOMEMORY", "komemory"
    AddSyn "MEMAUTOCALL", "memoryautocall"
    AddSyn "DAILYKO", "dailyko"
    AddSyn "TRADE", "tradedate", "strikedate", "initialvaluationdate"
    AddSyn "ISSUE", "issuedate", "issuancedate", "issuedatetcalendardays"   ' MS 這欄填的是發行日
    AddSyn "FINAL", "finalvaluationdate", "finalfixingdate", "redemptionvaluationdate", "finalvaluation"
    AddSyn "MATURITY", "maturitydate", "redemptiondate", "finalmaturitydate"
    AddSyn "TDCC", "tdcc", "tdcccode"
    AddSyn "INITPRICE", "initialprice"
    AddSyn "PGRATE", "pgrate"
    AddSyn "ACCRUALBARRIER", "accrualbarrier"    ' MS 的 DACN
    AddSyn "EXPRESSMULT", "firstperiodrebatemultiplier", "expressmultiplier"   ' FCN_Express：BARC、CA
    ' Super Bull 專用，只在 Product 是 Super Bull 時才讀（BARC 的 DACN／Express 也有 Autocall Barrier，不能併進 KO Barrier）
    AddSyn "SBLEVEL", "superbulllevel"           ' BNP
    AddSyn "ACBARRIER", "autocallbarrier"        ' HSBC、CA
    AddSyn "ACFREQ", "autocallfrequencym", "autocallfreqm", "autocalldateonetime"
    ' Super Bull 的 Autocall Coupon、BEN 的 Bonus Coupon 都是 flat 配息，只在 Product 是 Super Bull 或 BEN 時才讀
    AddSyn "FLATCPN", "autocallcouponflat", "bonuscouponflat", "autocallflatcoupon", "bonuscoupon", "couponflat"
    AddSyn "CPNBARRIER", "couponbarrier"         ' BEN 專用
    ' Range Accrual 專用：值只在 Product 是 Range Accrual 時才用（表頭仍算進每個商品的表頭偵測）；
    ' RACOUPON、DRACOUPON、RARATE 也用來判斷信是不是 RA
    AddSyn "TENORY", "tenory"                    ' BNP、玉山格式：年
    AddSyn "RAUF", "uf"                          ' 玉山格式：UF，Note Price = 100 - UF
    AddSyn "NCY", "noncallperiody"               ' BNP：年
    AddSyn "CALLFEATURE", "callfeature", "autocallfeature", "nc"   ' 玉山格式 NC 1Y*3M、GS 的 Autocall Feature、BARC 的 1Y
    AddSyn "ISCALLABLE", "iscallable"            ' BNP
    AddSyn "RAUL", "underlying", "index", "condition1"   ' BNP 的 USS10Y、BARC 的 CMS10Y、玉山格式 USD SOFR CMS 10Y ≦ 5.0%
    ' 玉山格式的 Coupon 欄（說明 Coupon frequency=Callable frequency 在表頭下一列或同一格）：
    ' Y1 : 8% p.a. fixed／Y2+ : 6% x n/N p.a, Floor@ 2.5%
    AddSyn "RACOUPON", "couponfrequencycallablefrequency", "couponcouponfrequencycallablefrequency"
    AddSyn "DRACOUPON", "dracoupon"              ' BNP
    AddSyn "RARATE", "rarate"                    ' BARC
End Sub

Private Sub AddSyn(ByVal field As String, ParamArray keys() As Variant)
    Dim k As Variant
    For Each k In keys
        mHeaderMap(CStr(k)) = field
    Next k
End Sub

' 正規化：轉小寫，只留英數與「?」（用來分辨 MS 的 "Put KI?" 與 "Put KI (%)"）。
' 玉山格式的 Issue date (T+7)：T+ 後面的天數不影響表頭，一律當成 issuedate
Private Function NormKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "?" Then out = out & ch
    Next i
    If out Like "issuedatet#" Or out Like "issuedatet##" Then out = "issuedate"
    NormKey = out
End Function

Private Function CellKey(ByVal v As Variant) As String
    If VarType(v) = vbString Then CellKey = NormKey(v)
End Function

' 空白、「-」、「N/A」、「NA」都算沒有值
Private Function IsBlankVal(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        IsBlankVal = True
    ElseIf VarType(v) = vbString Then
        s = LCase$(CollapseSpaces(v))
        IsBlankVal = (s = "" Or s = "-" Or s = "n/a" Or s = "na")
    End If
End Function

' 不換行空白、全形空白換成一般空白，去頭尾並合併連續空白
Private Function CollapseSpaces(ByVal s As String) As String
    s = Trim$(Replace(Replace(s, ChrW(160), " "), ChrW(&H3000), " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CollapseSpaces = s
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
    ElseIf keys.Exists("rarate") Then   ' RA Rate：BARC 的 Range Accrual（也有 Reoffer，要比 CA 先比）
        IssuerFromSignature = "BARC"
    ElseIf keys.Exists("datescheduleconfig") Or keys.Exists("reoffer") Or keys.Exists("wrapper") Then
        IssuerFromSignature = "CA"
    ElseIf keys.Exists("settlementfrequency") Or keys.Exists("offerprice") Or InStr(allText, "AUTOPRICER") > 0 Then
        IssuerFromSignature = "SG"
    ElseIf keys.Exists("issuerprodref") Or keys.Exists("issuedatelag") Then
        IssuerFromSignature = "DBS"
    ElseIf keys.Exists("clientref") Or keys.Exists("redemptionvaluationdate") Or InStr(allText, "ESS_TW_") > 0 _
           Or keys.Exists("strategycode") Then   ' Strategy Code：BNP 的 BEN
        IssuerFromSignature = "BNP"
    ElseIf keys.Exists("systemremarks") Then
        IssuerFromSignature = "JPM"
    ElseIf keys.Exists("message") Then
        IssuerFromSignature = "Natixis"
    ElseIf (keys.Exists("comments") Or keys.Exists("commentremarks")) And keys.Exists("id") Then   ' Comment/Remarks：HSBC 的 BEN
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
' 改價寫法取最後的值：96.20%→96.15%、99.65->99.62、99.92 >> 99.89、99.92 > 99.89
' 單一個 > 只在兩邊都是數字時算改價
Private Function LatestText(ByVal s As String, ByRef revised As Boolean) As String
    Dim arrow As String, p As Long
    Dim before As String, after As String

    arrow = ChrW(&H2192)
    s = Replace(Replace(Replace(s, "->", arrow), "=>", arrow), ">>", arrow)
    p = InStr(s, ">")
    Do While p > 0
        before = Replace(RTrim$(Left$(s, p - 1)), "%", "")
        after = LTrim$(Mid$(s, p + 1))
        If Right$(before, 1) Like "#" And Left$(after, 1) Like "#" Then
            s = Left$(s, p - 1) & arrow & Mid$(s, p + 1)
        End If
        p = InStr(p + 1, s, ">")
    Loop

    p = InStrRev(s, arrow)
    revised = (p > 0)
    If revised Then LatestText = Trim$(Mid$(s, p + 1)) Else LatestText = s
End Function

' 支援：數字、" 20.00 "、"80%"、"1m"；hadPct 回傳是否帶 %
' 有逗號的不認（"105%, 100%" 是列出多個值、"98,50%" 是逗號當小數點）；這些欄位不會用到千分位
Private Function ParseNumber(ByVal v As Variant, ByRef n As Double, ByRef hadPct As Boolean) As Boolean
    Dim s As String

    hadPct = False
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            n = CDbl(v)
            ParseNumber = True
        Case vbString
            s = CollapseSpaces(v)
            hadPct = (InStr(s, "%") > 0)
            s = Replace(Replace(s, "%", ""), " ", "")
            If LCase$(Right$(s, 1)) = "m" Then s = Left$(s, Len(s) - 1)
            If Not IsPlainNumber(s) Then Exit Function
            n = Val(s)
            ParseNumber = True
    End Select
End Function

' -?數字(.數字)?
Private Function IsPlainNumber(ByVal s As String) As Boolean
    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    If s = "" Or s Like "*[!0-9.]*" Then Exit Function
    If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
    If Left$(s, 1) = "." Or Right$(s, 1) = "." Then Exit Function
    IsPlainNumber = True
End Function

' 支援：日期值、yyyy-mm-dd、yyyy/mm/dd、dd/mm/yyyy（HSBC）、20Oct26、15-Oct-26、20-Oct-2026、
' 20 Oct 2026；數字格式的日期後面可帶時間
Private Function ParseDateVal(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim s As String, p() As String
    Dim y As Long, m As Long, dd As Long

    If VarType(v) = vbDate Then
        d = DateSerial(Year(v), Month(v), Day(v))
        ParseDateVal = True
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function

    s = CollapseSpaces(v)
    If ParseMonthNameDate(s, d) Then
        ParseDateVal = True
        Exit Function
    End If

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

' 日 + 英文月份縮寫 + 2 或 4 位年：20Oct26、15-Oct-26、20-Oct-2026、20 Oct 2026
Private Function ParseMonthNameDate(ByVal s As String, ByRef d As Date) As Boolean
    Dim t As String, i As Long, pos As Long
    Dim dayPart As String, monPart As String, yearPart As String
    Dim y As Long, m As Long, dd As Long

    t = Replace(Replace(s, "-", ""), " ", "")
    For i = 1 To Len(t)
        If Mid$(t, i, 1) Like "[A-Za-z]" Then
            pos = i
            Exit For
        End If
    Next i
    If pos < 2 Or pos > 3 Then Exit Function

    dayPart = Left$(t, pos - 1)
    monPart = LCase$(Mid$(t, pos, 3))
    yearPart = Mid$(t, pos + 3)
    If dayPart Like "*[!0-9]*" Or Len(monPart) <> 3 Then Exit Function
    If Not (yearPart Like "##" Or yearPart Like "####") Then Exit Function

    m = InStr("janfebmaraprmayjunjulaugsepoctnovdec", monPart)
    If m = 0 Or (m - 1) Mod 3 <> 0 Then Exit Function
    m = (m - 1) \ 3 + 1
    dd = CLng(dayPart)
    y = CLng(yearPart)
    If y < 100 Then y = y + 2000
    If dd < 1 Or dd > 31 Then Exit Function

    d = DateSerial(y, m, dd)
    ParseMonthNameDate = (Month(d) = m And Day(d) = dd)
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

' Product 對應到 Variable!D 的寫法；上手各自的寫法先轉換：
'   Citi 的 FCA → FCN；DAC（多數上手）、DRA（DBS）、VMRAN（UBS）→ DACN；Express FCN → FCN_Express；
'   SuperBull（BNP、HSBC）→ Super Bull；BEN、BEN WITH LOOKBACK（MS）→ Regular BEN；
'   Digital Note（CA）、DCN（SG）、DIGIT、DIGIT WITH LOOKBACK（MS）、Bonus RCN（HSBC）→ Digital BEN；
'   TW DRA（BNP）、RATES CRAN、RATES AUTOCALL RAN（SG）→ Range Accrual
' 轉換後不在 Variable!D 的留空並提醒
Private Function NormProduct(ByVal v As Variant) As String
    Dim s As String
    Dim wsV As Worksheet
    Dim lastRow As Long, r As Long

    s = CollapseSpaces(CStr(v))
    If s = "" Then Exit Function
    Select Case UCase$(s)
        Case "FCA": s = "FCN"
        Case "DAC", "DRA", "VMRAN": s = "DACN"
        Case "EXPRESS FCN": s = "FCN_Express"
        Case "SUPERBULL": s = "Super Bull"
        Case "BEN", "BEN WITH LOOKBACK": s = "Regular BEN"
        Case "DIGITAL NOTE", "DCN", "DIGIT", "DIGIT WITH LOOKBACK", "BONUS RCN": s = "Digital BEN"
        Case "TW DRA", "RATES CRAN", "RATES AUTOCALL RAN": s = "Range Accrual"
    End Select

    Set wsV = ThisWorkbook.Worksheets("Variable")
    lastRow = wsV.Cells(wsV.Rows.Count, "D").End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(wsV.Cells(r, "D").Value)), s, vbTextCompare) = 0 Then
            NormProduct = Trim$(CStr(wsV.Cells(r, "D").Value))
            Exit Function
        End If
    Next r
    FlagField "PRODUCT", "Product 認不得：" & ShowVal(v) & "，A2 留空"
End Function

' KO Type 對應到 Variable!C：Daily / Daily Memory / Period End / Period End Memory（不分大小寫）
' MS：OBSERVATION = 每期，由 KO Memory（Yes / No）決定是否記憶式
' Citi：沒有 KO Type，由 Daily KO（每日觀察）、Memory Autocall（記憶式）兩欄 TRUE / FALSE 判斷
' 認不得或互相矛盾時留空並提醒
Private Function NormKOType(ByVal koType As Variant, ByVal koMemory As Variant, _
                            ByVal memAutocall As Variant, ByVal dailyKO As Variant) As String
    Dim s As String, base As String, mem As String
    Dim isDaily As Boolean, isMemory As Boolean

    s = LCase$(CollapseSpaces(CStr(koType)))
    If s = "" Then
        If IsBlankVal(memAutocall) And IsBlankVal(dailyKO) Then Exit Function
        If Not ParseFlag(dailyKO, isDaily) Or Not ParseFlag(memAutocall, isMemory) Then
            FlagField "KOTYPE", "Daily KO／Memory Autocall 認不得：" & ShowBlank(dailyKO) & "／" & _
                                ShowBlank(memAutocall) & "，L2 留空"
            Exit Function
        End If
        If isDaily Then base = "Daily" Else base = "Period End"
        If isMemory Then NormKOType = base & " Memory" Else NormKOType = base
        Exit Function
    End If

    Select Case s
        Case "daily": base = "Daily"
        Case "daily memory": base = "Daily Memory"
        Case "period end": base = "Period End"
        Case "period end memory": base = "Period End Memory"
        Case "observation": base = "Period End"   ' MS：每期觀察
        Case Else
            FlagField "KOTYPE", "KO Type 認不得：" & ShowVal(koType) & "，L2 留空"
            Exit Function
    End Select

    mem = LCase$(CollapseSpaces(CStr(koMemory)))
    If mem = "" Then
        If s = "observation" Then
            FlagField "KOTYPE", "KO Type 是 OBSERVATION，但 KO Memory 是空白，L2 留空"
        Else
            NormKOType = base
        End If
    ElseIf mem = "yes" Or mem = "y" Then
        NormKOType = Replace(base, " Memory", "") & " Memory"
    ElseIf mem = "no" Or mem = "n" Then
        If Right$(base, 7) = " Memory" Then
            FlagField "KOTYPE", "KO Type 是 " & ShowVal(koType) & "，但 KO Memory 是 " & ShowVal(koMemory) & "，L2 留空"
        Else
            NormKOType = base
        End If
    Else
        FlagField "KOTYPE", "KO Memory 認不得：" & ShowVal(koMemory) & "，L2 留空"
    End If
End Function

' TRUE / FALSE（Excel 的布林值或文字）
Private Function ParseFlag(ByVal v As Variant, ByRef b As Boolean) As Boolean
    Dim s As String
    If VarType(v) = vbBoolean Then
        b = v
        ParseFlag = True
    ElseIf VarType(v) = vbString Then
        s = LCase$(CollapseSpaces(v))
        If s = "true" Or s = "false" Then
            b = (s = "true")
            ParseFlag = True
        End If
    End If
End Function

' Barrier Type 對應到 AKI / MKI / EKI / NONE；認不得時留空並提醒
Private Function NormBarrierType(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(CollapseSpaces(CStr(v)))
    Select Case s
        Case ""
        Case "AKI", "MKI", "EKI", "NONE"
            NormBarrierType = s
        Case "EUROPEAN", "AT MATURITY"   ' MS、SG
            NormBarrierType = "EKI"
        Case "AMERICAN"
            NormBarrierType = "AKI"
        Case "NO", "NIL"
            NormBarrierType = "NONE"
        Case Else
            FlagField "BARRIERTYPE", "Barrier Type 認不得：" & ShowVal(v) & "，R2 留空"
    End Select
End Function

' 期初價格：Opening → 開盤價；Closing → 收盤價；VWAP → VWAP
Private Function NormInitPrice(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(CollapseSpaces(CStr(v)))
    If s = "" Then Exit Function
    If InStr(s, "open") > 0 Then
        NormInitPrice = "開盤價"
    ElseIf InStr(s, "clos") > 0 Then
        NormInitPrice = "收盤價"
    ElseIf InStr(s, "vwap") > 0 Then
        NormInitPrice = "VWAP"
    Else
        FlagField "INITPRICE", "期初價格認不得：" & ShowVal(v) & "，O9 留空"
    End If
End Function

'==================================================
' Range Accrual
'==================================================
' Product 是 Range Accrual，或有 RA 才有的欄位（玉山格式的配息欄、BNP 的 DRA coupon、BARC 的 RA Rate）
Private Function LooksLikeRA(ByVal product As String) As Boolean
    LooksLikeRA = (product = "Range Accrual" Or HasField("RACOUPON") Or HasField("DRACOUPON") Or HasField("RARATE"))
End Function

' 去掉空白並轉大寫：USD SOFR CMS 10Y → USDSOFRCMS10Y
Private Function CompactUpper(ByVal v As Variant) As String
    CompactUpper = UCase$(Replace(CollapseSpaces(CStr(v)), " ", ""))
End Function

' Non-call（月）：BNP 的 Non call period (Y) 乘 12；玉山格式的 Call Feature（NC 1Y*3M、NC 3M*3M、NC 0.25Y X 3M）、
' GS 的 Autocall Feature、BARC 的 NC（1Y）取 NC 後面的期間。Is Callable 是 No 或不是整月數時留空並提醒
Private Function RANonCall() As Variant
    Dim ncy As Variant, s As String, months As Double

    RANonCall = Empty
    If LCase$(CollapseSpaces(CStr(GetVal("ISCALLABLE")))) = "no" Then
        FlagField "CALLFEATURE", "Is Callable 是 No，E2 留空"
        Exit Function
    End If

    ncy = NumField("NCY", 0)
    If Not IsEmpty(ncy) Then
        months = ncy * 12
    ElseIf mVal.Exists("CALLFEATURE") Then
        s = CompactUpper(GetVal("CALLFEATURE"))
        If Left$(s, 2) = "NC" Then s = Mid$(s, 3)
        If Not LeadingPeriodMonths(s, months) Then
            FlagField "CALLFEATURE", FieldLabel("CALLFEATURE") & " 認不得：" & ShowVal(GetVal("CALLFEATURE")) & "，E2 留空"
            Exit Function
        End If
    Else
        Exit Function
    End If

    If months <> Int(months) Then
        FlagField "CALLFEATURE", "Non-call 是 " & months & " 個月，不是整月數，E2 留空"
        Exit Function
    End If
    RANonCall = months
End Function

' 開頭的期間換成月：1Y*3M → 12、0.25YX3M → 3、3M*3M → 3（後面的字不看）
Private Function LeadingPeriodMonths(ByVal s As String, ByRef months As Double) As Boolean
    Dim i As Long

    i = 1
    Do While i <= Len(s)
        If Not Mid$(s, i, 1) Like "[0-9.]" Then Exit Do
        i = i + 1
    Loop
    If Not IsPlainNumber(Left$(s, i - 1)) Then Exit Function
    Select Case Mid$(s, i, 1)
        Case "Y": months = Val(Left$(s, i - 1)) * 12
        Case "M": months = Val(Left$(s, i - 1))
        Case Else: Exit Function
    End Select
    LeadingPeriodMonths = True
End Function

' 「數字 + Y」：5Y → 5
Private Function ParseYears(ByVal v As Variant, ByRef n As Double) As Boolean
    Dim s As String

    If VarType(v) <> vbString Then Exit Function
    s = CompactUpper(v)
    If Right$(s, 1) <> "Y" Then Exit Function
    s = Left$(s, Len(s) - 1)
    If Not IsPlainNumber(s) Then Exit Function
    n = Val(s)
    ParseYears = True
End Function

' Tenor（年）：Tenor (Y)（BNP、玉山格式）；BARC 的 Tenor 寫成 5Y。沒寫單位的 Tenor 不猜
Private Function RATenor() As Variant
    Dim y As Double

    RATenor = NumField("TENORY", 0)
    If Not IsEmpty(RATenor) Or IsFlagged("TENORY") Or Not mVal.Exists("TENOR") Then Exit Function
    If ParseYears(GetVal("TENOR"), y) Then
        RATenor = y
    Else
        FlagField "TENOR", "Tenor 是 " & ShowVal(GetVal("TENOR")) & "，沒有寫年（Y），P2 留空"
    End If
End Function

' 標的只有格式裡的兩種：連 CMS → USD SOFR CMS 10Y（USS10Y、CMS10Y、USD SOFR CMS 10Y ≦ 5.0%）；
' 連 CMT → USD CMT 10Y（UCMT10、USD CMT 10Y）。其他（例：股權連結的 TSM UN ≧ 75%、EUR CMS 10Y）留空並提醒
Private Function RAUnderlying(ByVal v As Variant) As String
    Dim s As String, rateCount As Long

    If IsBlankVal(v) Then Exit Function
    s = CompactUpper(v)
    rateCount = (Len(s) - Len(Replace(Replace(s, "CMS", ""), "CMT", ""))) \ 3
    If s = "USS10Y" Or (rateCount = 1 And IsUSDRate(s, "CMS10Y")) Then
        RAUnderlying = "USD SOFR CMS 10Y"
    ElseIf s = "UCMT10" Or (rateCount = 1 And IsUSDRate(s, "CMT10Y")) Then
        RAUnderlying = "USD CMT 10Y"
    Else
        FlagField "UL1", "標的認不得：" & ShowVal(v) & "，F2 留空（Range Accrual 只有 USD SOFR CMS 10Y、USD CMT 10Y）"
    End If
End Function

' s 有 rate（CMS10Y／CMT10Y），前面只有 USD、SOFR 或條件符號（0%≦USDSOFRCMS10Y≦5.0%）；EURCMS10Y 不算
Private Function IsUSDRate(ByVal s As String, ByVal rate As String) As Boolean
    Dim p As Long, prefix As String

    p = InStr(s, rate)
    If p = 0 Then Exit Function
    prefix = Left$(s, p - 1)
    If Right$(prefix, 4) = "SOFR" Then prefix = Left$(prefix, Len(prefix) - 4)
    If Right$(prefix, 3) = "USD" Then prefix = Left$(prefix, Len(prefix) - 3)
    IsUSDRate = Not (Right$(prefix, 1) Like "[A-Z]")
End Function

' Floor Coupon：配息欄有 Floor@ 2.5% 時填該值，沒有填 0；Floor 後面讀不到數字時留空並提醒
Private Function RAFloor(ByVal v As Variant) As Variant
    Dim s As String, p As Long, n As Double, pct As Boolean

    s = CompactUpper(v)
    p = InStr(s, "FLOOR")
    If p = 0 Then
        RAFloor = 0
        Exit Function
    End If
    s = Mid$(s, p + 5)
    If Left$(s, 1) = "@" Then s = Mid$(s, 2)
    p = InStr(s, "%")
    If p > 0 Then s = Left$(s, p)
    If ParseNumber(s, n, pct) Then
        RAFloor = Round(n, 4)
    Else
        RAFloor = Empty
        FlagField "RACOUPON", "Floor 認不得：" & ShowVal(v) & "，V2 留空"
    End If
End Function

Private Function AddWarn(ByVal baseText As String, ByVal newText As String) As String
    If baseText = "" Then
        AddWarn = "- " & newText
    Else
        AddWarn = baseText & vbCrLf & "- " & newText
    End If
End Function
