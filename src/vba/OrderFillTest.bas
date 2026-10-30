Attribute VB_Name = "OrderFillTest"
Option Explicit

'==================================================
' 起單回填測試
' 1. 樣本：把「起單信件格式.xlsx」14 張上手表各自貼進「起單回填」並執行填入，
'    再和 ELN每日起單建檔中同一 TDCC 的紀錄比對。
' 2. 規則：在貼上區寫入合成的信件內容，逐條檢查對照規則。
' 會覆蓋 Term 與貼上區，請在複本上執行。
'==================================================

' 樣本與規則測試一起跑
Public Sub Test_OrderFill_AllIssuers()
    Dim path As String
    Dim result As String

    If MsgBox("測試會覆蓋 Term 與「起單回填」的內容，請確認是在複本上執行。要繼續嗎？", _
              vbYesNo + vbExclamation, "起單回填測試") = vbNo Then Exit Sub

    path = ThisWorkbook.Path & Application.PathSeparator & "起單信件格式.xlsx"
    If Dir(path) = "" Then
        path = CStr(Application.GetOpenFilename("Excel 檔案 (*.xlsx), *.xlsx", , "選擇起單信件格式.xlsx"))
        If path = "False" Then Exit Sub
    End If

    result = Test_OrderFill_Report(path) & vbCrLf & Test_OrderFill_RulesReport()
    Debug.Print result
    MsgBox Left$(result, 1000), vbInformation, "起單回填測試"
End Sub

' 只跑規則測試（不需要樣本檔）
Public Sub Test_OrderFill_Rules()
    Dim result As String

    If MsgBox("測試會覆蓋 Term 與「起單回填」的內容，請確認是在複本上執行。要繼續嗎？", _
              vbYesNo + vbExclamation, "起單回填測試") = vbNo Then Exit Sub

    result = Test_OrderFill_RulesReport()
    Debug.Print result
    MsgBox Left$(result, 1000), vbInformation, "起單回填測試"
End Sub

' 回傳測試報告文字；第一行為「PASS x/y」
Public Function Test_OrderFill_Report(ByVal samplePath As String) As String
    Dim cases As Variant, tc As Variant
    Dim wbS As Workbook, wsP As Worksheet, wsT As Worksheet
    Dim report As String, detail As String, fails As String
    Dim passCount As Long, total As Long
    Dim i As Long

    Set wsP = ThisWorkbook.Worksheets("起單回填")
    Set wsT = ThisWorkbook.Worksheets("Term")
    Set wbS = Workbooks.Open(samplePath, ReadOnly:=True)

    cases = ExpectedCases()
    For i = LBound(cases) To UBound(cases)
        tc = cases(i)
        total = total + 1

        OrderFill_Clear
        wbS.Worksheets(CStr(tc(0))).UsedRange.Copy Destination:=wsP.Range("A7")
        Application.CutCopyMode = False

        fails = ""
        If Not OrderFill_Run(True, report) Then
            fails = "  填入失敗：" & Replace(report, vbCrLf, " ")
        Else
            fails = CompareTerm(wsT, tc)
        End If

        If fails = "" Then
            passCount = passCount + 1
            detail = detail & "OK   " & tc(0) & vbCrLf
        Else
            detail = detail & "FAIL " & tc(0) & vbCrLf & fails
        End If
    Next i

    wbS.Close SaveChanges:=False
    Test_OrderFill_Report = "PASS " & passCount & "/" & total & vbCrLf & detail
End Function

'==================================================
' 清空回歸測試（#4）：選過 C3 後 Excel 會加一個隱藏的下拉框 Shape，按清空不能因此出現 1004；
' 貼上區的圖片要刪掉，按鈕與 C3 下拉清單要保留；E5 的保本型商品要清掉。
'==================================================
Public Sub Test_OrderFill_Clear()
    Dim result As String

    If MsgBox("測試會清空「起單回填」的內容，請確認是在複本上執行。要繼續嗎？", _
              vbYesNo + vbExclamation, "起單回填測試") = vbNo Then Exit Sub

    result = Test_OrderFill_ClearReport()
    Debug.Print result
    MsgBox result, vbInformation, "起單回填測試"
End Sub

' 回傳測試報告文字；第一行為 PASS 或 FAIL
Public Function Test_OrderFill_ClearReport() As String
    Const PASTED_NAME As String = "OrderFillTest_Pasted"
    Dim ws As Worksheet
    Dim shp As Shape
    Dim fails As String
    Dim btnCount As Long

    Set ws = ThisWorkbook.Worksheets("起單回填")
    btnCount = ws.Buttons.Count

    ws.Activate
    ws.Range("C3").Select    ' 讓 Excel 建立 C3 下拉清單的隱藏下拉框
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("B8").Left, ws.Range("B8").Top, 40, 20)
    shp.Name = PASTED_NAME   ' 代替信件貼上的圖片
    ws.Range("E5").Value = "Range Accrual"

    On Error Resume Next
    OrderFill_Clear
    If Err.Number <> 0 Then fails = fails & "  清空出錯：" & Err.Number & " " & Err.Description & vbCrLf
    Err.Clear
    Set shp = Nothing
    Set shp = ws.Shapes(PASTED_NAME)
    On Error GoTo 0

    If Not shp Is Nothing Then
        fails = fails & "  貼上區的圖片沒有刪掉" & vbCrLf
        shp.Delete
    End If
    If ws.Buttons.Count <> btnCount Then
        fails = fails & "  按鈕數量 預期 " & btnCount & " 實際 " & ws.Buttons.Count & vbCrLf
    End If
    If Not HasListValidation(ws.Range("C3")) Then fails = fails & "  C3 下拉清單不見了" & vbCrLf
    If Not IsEmpty(ws.Range("E5").Value) Then fails = fails & "  E5 保本型商品沒有清掉" & vbCrLf

    If fails = "" Then
        Test_OrderFill_ClearReport = "PASS"
    Else
        Test_OrderFill_ClearReport = "FAIL" & vbCrLf & fails
    End If
End Function

Private Function HasListValidation(ByVal cell As Range) As Boolean
    Dim t As Long

    t = -1
    On Error Resume Next
    t = cell.Validation.Type
    On Error GoTo 0
    HasListValidation = (t = xlValidateList)
End Function

Private Function CompareTerm(ByVal wsT As Worksheet, ByVal tc As Variant) As String
    Dim out As String
    Dim cols As Variant
    Dim i As Long
    Dim expTrade As String

    ' tc: 0 工作表, 1 C2 上手, 2 A2, 3 B2, 4 D2, 5 E2, 6-10 F2:J2, 11 K2, 12 L2, 13 M2,
    '     14 N2, 15 O2, 16 P2, 17 R2, 18 S2, 19 T2, 20 O5, 21 O6, 22 O7, 23 O8, 24 O9
    cols = Array("", "C2", "A2", "B2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", "K2", "L2", "M2", _
                 "N2", "O2", "P2", "R2", "S2", "T2", "O5", "O6", "O7", "O8", "O9")

    For i = 1 To UBound(cols)
        If i = 20 And CStr(tc(i)) = "" Then
            expTrade = Format$(Date, "yyyy/mm/dd")   ' 信裡沒有 Trade date → 今天
            out = out & CompareCell(wsT.Range("O5"), expTrade)
        Else
            out = out & CompareCell(wsT.Range(CStr(cols(i))), tc(i))
        End If
    Next i
    out = out & CompareCell(wsT.Range("Q2"), "M")
    CompareTerm = out
End Function

Private Function CompareCell(ByVal cell As Range, ByVal expected As Variant) As String
    Dim actual As Variant
    Dim ok As Boolean
    Dim a As String

    actual = cell.Value
    If VarType(actual) = vbDate Then
        a = Format$(actual, "yyyy/mm/dd")
    ElseIf IsEmpty(actual) Then
        a = ""
    Else
        a = CStr(actual)
    End If

    If VarType(expected) = vbString Then
        ok = (a = expected)
    ElseIf IsNumeric(actual) And Not IsEmpty(actual) Then
        ok = (Abs(CDbl(actual) - CDbl(expected)) < 0.00001)
    End If

    If Not ok Then CompareCell = "  " & cell.Address(False, False) & " 預期 [" & expected & "] 實際 [" & a & "]" & vbCrLf
End Function

' 預期值取自 ELN每日起單建檔（同一 TDCC 的回填紀錄）；Trade date 空白代表信裡沒有、應填今天
Private Function ExpectedCases() As Variant
    ExpectedCases = Array( _
        Array("Barclays", "BARC", "FCN", "", "USD", 1, "NVDA UW", "AMAT UW", "TSM UN", "", "", 80, "Daily Memory", 100, 14.72, 98.5, 6, "NONE", "", 0, "2026/10/15", "2026/10/22", "2027/04/22", "2027/04/27", "開盤價"), _
        Array("BBVA", "BBVA", "FCN", "", "USD", 3, "MU UW", "SNDK UW", "COHR UN", "", "", 80, "Daily Memory", 100, 23.5, 97.22, 8, "EKI", 65, 0, "2026/10/12", "2026/10/20", "2027/06/21", "2027/06/23", "收盤價"), _
        Array("BNP", "BNP", "FCN", "", "USD", 2, "INTC UW", "AVGO UW", "NVDA UW", "SPCX UW", "TSLA UW", 75, "Daily Memory", 100, 14.1, 97.49, 6, "EKI", 65, 0, "2026/10/13", "2026/10/20", "2027/04/20", "2027/04/23", "開盤價"), _
        Array("CA", "CA", "FCN", "", "EUR", 11, "TSM UN", "", "", "", "", 80, "Daily", 95, 8, 99.88, 12, "EKI", 75, 0, "2026/10/13", "2026/10/20", "2027/10/20", "2027/10/22", "收盤價"), _
        Array("Citi", "Citi", "FCN", "", "USD", 1, "TSM UN", "NVDA UW", "GOOG UW", "META UW", "", 80, "Daily Memory", 105, 15.6, 99.15, 7, "EKI", 75, 0, "2026/10/08", "2026/10/16", "2027/05/17", "2027/05/20", "收盤價"), _
        Array("DBS", "DBS", "FCN", "", "JPY", 1, "AVGO UW", "INTC UW", "TSM UN", "", "", 80, "Daily Memory", 100, 10.17, 96.89, 12, "EKI", 70, 0, "", "2026/10/26", "2027/10/26", "2027/10/28", "收盤價"), _
        Array("GS", "GS", "FCN", "148000011877", "USD", 1, "NVDA UW", "TSM UN", "AMD UW", "MSFT UW", "", 80, "Daily Memory", 105, 11.64, 97.5, 12, "EKI", 60, 0, "2026/10/13", "2026/10/20", "2027/10/20", "2027/10/22", "收盤價"), _
        Array("HSBC", "HSBC", "FCN", "325000133781", "USD", 1, "AMD UQ", "TSM UN", "AMAT UQ", "", "", 100, "Daily Memory", 120, 45.33, 96.15, 2, "NONE", "", 0, "2026/10/13", "2026/10/20", "2026/12/21", "2026/12/24", "收盤價"), _
        Array("JPM", "JPM", "FCN", "", "JPY", 2, "8035 JT", "1321 JT", "", "", "", 79.18, "Daily Memory", 100, 10, 97.9, 9, "NONE", "", 0, "", "2026/10/20", "2027/07/20", "2027/07/23", "收盤價"), _
        Array("MS", "MS", "FCN", "147000029936", "USD", 1, "TSM UN", "NVDA UW", "GOOG UW", "AMZN UW", "", 91, "Period End", 100, 8, 99.2, 1, "EKI", 80, 0, "", "2026/10/20", "2026/11/20", "2026/11/25", "收盤價"), _
        Array("Natixis", "Natixis", "FCN", "", "USD", 1, "AMD UW", "ARM UW", "MU UW", "", "", 65, "Daily Memory", 110, 11.8, 96.15, 12, "NONE", "", 0, "", "2026/10/20", "2027/10/20", "2027/10/25", "收盤價"), _
        Array("Nomura", "Nomura", "FCN", "", "USD", 3, "MU UW", "SNDK UW", "TSLA UW", "AMD UW", "", 80, "Daily Memory", 100, 20, 97.53, 8, "EKI", 65, 0, "", "2026/10/20", "2027/06/21", "2027/06/24", "收盤價"), _
        Array("SG", "SG", "FCN", "", "USD", 1, "TSM UN", "NVDA UW", "MSFT UW", "", "", 75, "Daily Memory", 105, 9, 97.7, 12, "NONE", "", 0, "2026/10/13", "2026/10/20", "2027/10/20", "2027/10/25", "收盤價"), _
        Array("UBS", "UBS", "FCN", "", "USD", 1, "LITE UW", "TSM UN", "NVDA UW", "ASML UW", "", 66.25, "Daily Memory", 100, 12, 98, 6, "NONE", "", 0, "", "2026/10/20", "2027/04/20", "2027/04/23", "開盤價") _
    )
End Function

'==================================================
' 規則測試：PasteBase 寫入一封基本的起單信件，各案例改幾格後填入並檢查
'   A 欄起依序：Product, Currency, Guaranteed Periods (m), BBG Code 1, Strike (%), KO Type,
'   KO Barrier (%), Coupon p.a. (%), Upfront / NotePrice (%), Tenor (m), Barrier Type,
'   KI Barrier (%), Trade Date, Issue Date, Final Valuation Date, Maturity Date（第 9 列表頭、第 10 列數值）
'==================================================
' 回傳測試報告文字；第一行為「RULES PASS x/y」
Public Function Test_OrderFill_RulesReport() As String
    Dim wsP As Worksheet, wsT As Worksheet
    Dim report As String, fails As String, detail As String
    Dim passCount As Long, total As Long
    Dim ok As Boolean
    Dim prod As Variant

    Set wsP = ThisWorkbook.Worksheets("起單回填")
    Set wsT = ThisWorkbook.Worksheets("Term")

    PasteBase wsP
    wsP.Range("E10").Value = 0.8
    wsP.Range("E10").NumberFormat = "0.00%"
    wsP.Range("I10").Value = 0.98
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 80) & CompareCell(wsT.Range("O2"), 0.98) & _
                               ExpectNote(report, "Note Price 是 0.98、沒有 %")
    Tally "百分比看儲存格格式；沒有 % 的小數照填並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("I9").Value = "Upfront (%)"
    wsP.Range("I10").Value = 0.0085
    wsP.Range("I10").NumberFormat = "0.00%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O2"), 99.15)
    Tally "Upfront 0.85%（% 格式）→ Note Price 99.15", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("K10"), "At Maturity"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("R2"), "EKI") & CompareCell(wsT.Range("S2"), 65)
    Tally "Barrier Type：At Maturity → EKI", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("K10"), "Knock-in"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("R2"), "") & ExpectNote(report, "Barrier Type 認不得：Knock-in")
    Tally "Barrier Type 認不得 → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("K10"), "N/A"
    PutText wsP.Range("L10"), "N/A"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("R2"), "NONE") & CompareCell(wsT.Range("S2"), "")
    Tally "Barrier Type、KI 都是 N/A → NONE", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("K10"), "NONE"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("S2"), "") & ExpectNote(report, "Barrier Type 是 NONE，但 KI 有值")
    Tally "Barrier Type 是 NONE 但 KI 有值 → KI 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("Q9").Value = "KO Memory"
    PutText wsP.Range("Q10"), "No"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "") & ExpectNote(report, "但 KO Memory 是 No")
    Tally "KO Type 是 Daily Memory、KO Memory 是 No → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("F10"), "OBSERVATION"
    wsP.Range("Q9").Value = "KO Memory"
    PutText wsP.Range("Q10"), "Yes"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "Period End Memory")
    Tally "MS：OBSERVATION + KO Memory Yes → Period End Memory", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("F10"), "Daily KO"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "") & ExpectNote(report, "KO Type 認不得：Daily KO")
    Tally "KO Type 認不得 → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("N10"), "22Oct26"
    PutText wsP.Range("O10"), "20-Oct-2027"
    PutText wsP.Range("P10"), "25 Oct 2027"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O6"), "2026/10/22") & CompareCell(wsT.Range("O7"), "2027/10/20") & _
                               CompareCell(wsT.Range("O8"), "2027/10/25")
    Tally "日期：22Oct26、20-Oct-2027、25 Oct 2027", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("I10"), "99.65->99.62"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O2"), 99.62) & ExpectNote(report, "Note Price 有修改")
    Tally "改價 99.65->99.62 → 99.62 並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A7"), "Initial price" & ChrW(&HFF1A) & "Opening price"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O9"), "開盤價")
    Tally "全形冒號：Initial price：Opening price → 開盤價", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A7"), "Initial price: VWAP"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O9"), "VWAP")
    Tally "期初價格 VWAP", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A12"), "Strike: 85%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), "") & ExpectNote(report, "Strike 不一致")
    Tally "同一欄兩處不一致 → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("O10"), "15-Oct-27"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O7"), "2027/10/15")
    Tally "日期：15-Oct-27", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("F10"), "OBSERVATION"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "") & ExpectNote(report, "KO Memory 是空白")
    Tally "MS：OBSERVATION 但沒有 KO Memory → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("F9").Value = "Daily KO"
    PutText wsP.Range("F10"), "Maybe"
    wsP.Range("Q9").Value = "Memory Autocall"
    wsP.Range("Q10").Value = True
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "") & ExpectNote(report, "Daily KO／Memory Autocall 認不得")
    Tally "Citi：Daily KO 不是 TRUE／FALSE → 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("F9").Value = "Daily KO"
    wsP.Range("F10").Value = False
    wsP.Range("Q9").Value = "Memory Autocall"
    wsP.Range("Q10").Value = True
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("L2"), "Period End Memory")
    Tally "Citi：Daily KO FALSE、Memory Autocall TRUE → Period End Memory", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("N9:N10").ClearContents
    PutText wsP.Range("A12"), "[ Request ID: HKGEQV1 ] [ Issue Date: 22Oct26 ] [ Issuer: GSI ]"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O6"), "2026/10/22")
    Tally "一格多組 [ 標題: 數值 ]（GS System Remark）", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("J9").Value = "No. of Periods"
    wsP.Range("J10").Value = 12
    wsP.Range("Q9").Value = "Settlement Frequency"
    PutText wsP.Range("Q10"), "Monthly"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("P2"), 12)
    Tally "SG：Settlement Frequency Monthly → Tenor = No. of Periods", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("J9").Value = "No. of Periods"
    wsP.Range("J10").Value = 4
    wsP.Range("Q9").Value = "Settlement Frequency"
    PutText wsP.Range("Q10"), "Quarterly"
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNote(report, "Settlement Frequency 不是 Monthly")
    Tally "SG：Settlement Frequency 不是 Monthly → Tenor 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("A9:A10").ClearContents
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & ExpectNote(report, "找不到 Product，A2 留空")
    Tally "非 SG 沒有 Product → A2 留空並提醒（不預設 FCN）", fails, passCount, total, detail

    PasteSGNoProduct wsP
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "FCN") & ExpectNote(report, "推定為 FCN")
    Tally "SG 沒有 Product、條件都符合 → 推定 FCN 並提醒", fails, passCount, total, detail

    PasteSGNoProduct wsP
    PutText wsP.Range("Q10"), "Coupon Barrier"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & ExpectNote(report, "找不到 Product，A2 留空")
    Tally "SG：Fixed Coupons 不是 All Periods／First Period／First Two Periods → A2 留空並提醒", fails, passCount, total, detail

    PasteSGNoProduct wsP
    PutText wsP.Range("G10"), "105%/100%/95%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & ExpectNote(report, "找不到 Product，A2 留空")
    Tally "SG：AutoCall 不是單一數字 → A2 留空並提醒", fails, passCount, total, detail

    PasteSGNoProduct wsP
    wsP.Range("E9:E10").ClearContents
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & ExpectNote(report, "找不到 Product，A2 留空")
    Tally "SG：沒有 Put Strike → A2 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("G10"), "105%, 100%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("M2"), "") & ExpectNote(report, "KO Barrier 格式不認得：105%, 100%")
    Tally "KO Barrier 用逗號列出多個值 → 留空並提醒（不讀成 105100）", fails, passCount, total, detail

    PasteSGNoProduct wsP
    PutText wsP.Range("G10"), "105%, 100%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & ExpectNote(report, "找不到 Product，A2 留空")
    Tally "SG：AutoCall 是 105%, 100% → 不推定 FCN", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("E10"), "98,50%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), "") & ExpectNote(report, "Strike 格式不認得：98,50%")
    Tally "Strike 逗號當小數點（98,50%）→ 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    SetTermCell wsT.Range("V2"), 5
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("V2"), 0)
    Tally "Product 是 FCN → V2（Floor Coupon）填 0", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "FCN_Stepdown"
    SetTermCell wsT.Range("V2"), 0
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "FCN_Stepdown") & CompareCell(wsT.Range("V2"), "")
    Tally "Product 是 FCN_Stepdown → V2 清空", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("A9:A10").ClearContents
    SetTermCell wsT.Range("V2"), 0
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & CompareCell(wsT.Range("V2"), "")
    Tally "沒有 Product、A2 留空 → V2 清空", fails, passCount, total, detail

    PasteSGNoProduct wsP
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "FCN") & CompareCell(wsT.Range("V2"), 0)
    Tally "SG 推定為 FCN → V2 填 0", fails, passCount, total, detail

    ' DACN、FCN_Express（#16）
    For Each prod In Array("DAC", "dra", "VMRAN", "dacn")
        PasteBase wsP
        PutText wsP.Range("A10"), CStr(prod)
        SetTermCell wsT.Range("V2"), Empty
        fails = RunFill(report)
        If fails = "" Then fails = CompareCell(wsT.Range("A2"), "DACN") & CompareCell(wsT.Range("V2"), 0)
        Tally "Product " & prod & " → DACN、V2 填 0", fails, passCount, total, detail
    Next prod

    PasteBase wsP
    PutText wsP.Range("A10"), "XYZ Note"
    SetTermCell wsT.Range("V2"), 0
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "") & CompareCell(wsT.Range("V2"), "") & _
                               ExpectNote(report, "Product 認不得：XYZ Note，A2 留空")
    fails = fails & ExpectNoNote(report, "找不到 Product")
    Tally "Product 不在對照表與 Variable 清單 → A2 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "express fcn"
    wsP.Range("C9:C10").ClearContents
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "FCN_Express") & CompareCell(wsT.Range("E2"), 0) & _
                               CompareCell(wsT.Range("V2"), 0)
    fails = fails & ExpectNoNote(report, "Guaranteed Period")
    Tally "express fcn → FCN_Express、沒有 Guaranteed Period 時 E2 填 0、V2 填 0", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "Express FCN"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), 0) & ExpectNote(report, "Guaranteed Period 固定 0，信裡是 1")
    Tally "FCN_Express 信裡 Guaranteed Period 是 1 → E2 仍填 0 並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "Express FCN"
    PutText wsP.Range("C10"), "TBD"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), 0) & ExpectNote(report, "E2 已填 0（不是留空）")
    Tally "FCN_Express 信裡 Guaranteed Period 認不得 → E2 仍填 0，提醒不寫留空", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "Express FCN"
    wsP.Range("Q9").Value = "First Period Rebate Multiplier"
    wsP.Range("Q10").Value = 1.5
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNote(report, "是 1.5，不是 1")
    Tally "FCN_Express 倍數不是 1（BARC）→ 提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "Express FCN"
    wsP.Range("Q9").Value = "Express Multiplier"
    wsP.Range("Q10").Value = 1
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNoNote(report, "不是 1")
    Tally "FCN_Express 倍數是 1（CA）→ 不提醒", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "DACN") & CompareCell(wsT.Range("K2"), 80) & _
                               CompareCell(wsT.Range("V2"), 0) & ExpectNote(report, "判定為 DACN")
    Tally "MS：Range Accrual 有 Put Strike → DACN 並提醒", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    wsP.Range("E9:E10").ClearContents
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "看起來是保本型商品") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "MS：Range Accrual 沒有 Strike、沒選保本型商品 → 不填入（V2 不變）", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    wsP.Range("C3").Value = "BARC"
    wsP.Range("E5").Value = "Range Accrual"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Range Accrual") & CompareCell(wsT.Range("K2"), "")
    Tally "不是 MS 的 Range Accrual（選了保本型商品）→ 不判定 DACN", fails, passCount, total, detail

    PasteRA wsP, "ESUN"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Range Accrual") & CompareCell(wsT.Range("C2"), "CA") & _
                               CompareCell(wsT.Range("D2"), "USD") & CompareCell(wsT.Range("E2"), 12) & _
                               CompareCell(wsT.Range("F2"), "USD SOFR CMS 10Y") & CompareCell(wsT.Range("G2"), "") & _
                               CompareCell(wsT.Range("K2"), "") & CompareCell(wsT.Range("L2"), "") & _
                               CompareCell(wsT.Range("M2"), "") & CompareCell(wsT.Range("N2"), "") & _
                               CompareCell(wsT.Range("O2"), 95.5) & CompareCell(wsT.Range("P2"), 10) & _
                               CompareCell(wsT.Range("Q2"), "Y") & CompareCell(wsT.Range("R2"), "") & _
                               CompareCell(wsT.Range("S2"), "") & CompareCell(wsT.Range("T2"), 100) & _
                               CompareCell(wsT.Range("V2"), 0) & CompareCell(wsT.Range("O9"), "")
    fails = fails & ExpectNoNote(report, "找不到 Strike") & ExpectNoNote(report, "找不到 KO") & _
                    ExpectNoNote(report, "Coupon") & ExpectNoNote(report, "找不到 Barrier Type") & _
                    ExpectNoNote(report, "找不到 期初價格") & ExpectNoNote(report, "不只一列")
    Tally "RA（玉山格式，表頭說明列、配息兩列）：E2 = NC、F2 = CMS 10Y、O2 = 100 - UF、P2 年、Q2 Y、V2 0，其他欄留空", _
          fails, passCount, total, detail

    PasteRA wsP, "ESUN1"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), 12) & CompareCell(wsT.Range("F2"), "USD SOFR CMS 10Y") & _
                               CompareCell(wsT.Range("O2"), 95.5) & CompareCell(wsT.Range("V2"), 0)
    Tally "RA（玉山格式，Coupon 說明和表頭同一格）", fails, passCount, total, detail

    PasteRA wsP, "ESUN", False
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "看起來是保本型商品") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "RA（玉山格式）沒選保本型商品 → 不填入（V2 不變）", fails, passCount, total, detail

    PasteRA wsP, "ESUN"
    PutText wsP.Range("D12"), "Y2+: 5.0% x n/N p.a, Floor@ 1.5%"
    PutText wsP.Range("G11"), "NC 0.25Y X 3M"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("V2"), 1.5) & CompareCell(wsT.Range("E2"), 3)
    Tally "RA：配息第二列的 Floor@ → V2、NC 0.25Y → E2 3", fails, passCount, total, detail

    PasteRA wsP, "ESUN"
    PutText wsP.Range("E11"), "TSM UN " & ChrW(&H2267) & " 75%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("F2"), "") & ExpectNote(report, "標的認不得") & _
                               ExpectNoNote(report, "找不到 標的")
    Tally "RA：標的不是 CMS／CMT 10Y → F2 留空並提醒", fails, passCount, total, detail

    PasteRA wsP, "BNP"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Range Accrual") & CompareCell(wsT.Range("C2"), "BNP") & _
                               CompareCell(wsT.Range("E2"), 12) & CompareCell(wsT.Range("F2"), "USD CMT 10Y") & _
                               CompareCell(wsT.Range("O2"), 96) & CompareCell(wsT.Range("P2"), 15) & _
                               CompareCell(wsT.Range("Q2"), "Y") & CompareCell(wsT.Range("V2"), 0) & _
                               CompareCell(wsT.Range("O6"), "2026/10/20") & CompareCell(wsT.Range("O8"), "2041/10/21")
    Tally "RA（BNP）：TW DRA、Non call period (Y)、UCMT10 → USD CMT 10Y", fails, passCount, total, detail

    PasteRA wsP, "BNP", False
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "看起來是保本型商品")
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "RA（BNP）TW DRA 沒選保本型商品 → 不填入", fails, passCount, total, detail

    PasteRA wsP, "BARC"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("C2"), "BARC") & CompareCell(wsT.Range("E2"), 12) & _
                               CompareCell(wsT.Range("F2"), "USD SOFR CMS 10Y") & CompareCell(wsT.Range("O2"), 96) & _
                               CompareCell(wsT.Range("P2"), 5) & CompareCell(wsT.Range("Q2"), "Y")
    fails = fails & ExpectNoNote(report, "Tenor")
    Tally "RA（BARC，RA Rate 判斷上手，有 Reoffer 也不是 CA）：Tenor 5Y → P2 5、NC 1Y → E2 12、CMS10Y", _
          fails, passCount, total, detail

    PasteRA wsP, "BARC", False
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "看起來是保本型商品")
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "RA（BARC）有 RA Rate、沒選保本型商品 → 不填入", fails, passCount, total, detail

    PasteRA wsP, "BNP"
    PutText wsP.Range("L10"), "No"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), "") & ExpectNote(report, "Is Callable 是 No") & _
                               ExpectNoNote(report, "找不到 Non call")
    Tally "RA（BNP）：Is Callable 是 No → E2 留空並提醒", fails, passCount, total, detail

    PasteRA wsP, "ESUN"
    wsP.Range("G11").ClearContents
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), "") & ExpectNote(report, "找不到 Non call period") & _
                               ExpectNoNote(report, "Guaranteed")
    Tally "RA：沒有 Call Feature → E2 留空，提醒找不到 Non call period", fails, passCount, total, detail

    PasteRA wsP, "BNP"
    wsP.Range("M10").Value = 0.1
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), "") & ExpectNote(report, "不是整月數")
    Tally "RA（BNP）：Non call period 0.1 年 → 1.2 個月不是整月數，E2 留空並提醒", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("J10").Value = 60
    PutText wsP.Range("A12"), "Tenor: 5Y"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("P2"), "") & ExpectNote(report, "Tenor") & _
                               ExpectNote(report, "不一致")
    Tally "Tenor 60（月）和 5Y 並存 → 算不一致，P2 留空", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("I9").Value = "UF"
    PutText wsP.Range("I10"), "1.5%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("O2"), "") & ExpectNote(report, "找不到")
    Tally "FCN 只有 UF → 不用 100 - UF（UF 只給 Range Accrual），O2 留空並提醒", fails, passCount, total, detail

    PasteRA wsP, "ESUN"
    PutText wsP.Range("E11"), "EUR CMS 10Y " & ChrW(&H2266) & " 3.0%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("F2"), "") & ExpectNote(report, "標的認不得")
    Tally "RA：EUR CMS 10Y 不是格式裡的標的 → F2 留空並提醒", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    wsP.Range("E5").Value = "Range Accrual"
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "這封信的商品是 DACN") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "選了保本型商品、MS 有 Strike（判定 DACN）→ 不填入", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("A10"), "XYZ NOTE"
    wsP.Range("E5").Value = "Range Accrual"
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "Product（XYZ NOTE）認不得")
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "選了保本型商品、Product 認不得 → 不填入", fails, passCount, total, detail

    PasteSGNoProduct wsP
    wsP.Range("E5").Value = "Range Accrual"
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "信裡沒有 Range Accrual 的欄位")
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "選了保本型商品、信沒有 Product 也沒有 RA 欄位（SG FCN）→ 不填入", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("E5").Value = "Range Accrual"
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "這封信的商品是 FCN") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "選了保本型商品、信裡是 FCN → 不填入（V2 不變）", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("J10"), "5Y"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("P2"), "") & CompareCell(wsT.Range("Q2"), "M") & _
                               ExpectNote(report, "Tenor 是 5Y（年）")
    Tally "不是 RA 的 Tenor 寫成年 → P2 留空並提醒", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    wsP.Range("Q9").Value = "Accrual Barrier"
    PutText wsP.Range("Q10"), "65.00%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 80) & _
                               ExpectNote(report, "Accrual Barrier（65）和 Strike（80）不同，K2 填 Strike")
    Tally "MS：Accrual Barrier 和 Put Strike 不同 → K2 填 Put Strike 並提醒", fails, passCount, total, detail

    PasteMSRangeAccrual wsP
    wsP.Range("Q9").Value = "Accrual Barrier"
    PutText wsP.Range("Q10"), "80.00%"
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNoNote(report, "Accrual Barrier")
    Tally "MS：Accrual Barrier 和 Put Strike 相同 → 不提醒", fails, passCount, total, detail

    PasteSGNoProduct wsP
    PutText wsP.Range("Q10"), "First Period"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "DACN") & CompareCell(wsT.Range("V2"), 0) & _
                               ExpectNote(report, "Fixed Coupons 是 First Period") & ExpectNote(report, "推定為 DACN")
    Tally "SG：Fixed Coupons 是 First Period → 推定 DACN 並提醒", fails, passCount, total, detail

    PasteSGNoProduct wsP
    PutText wsP.Range("Q10"), "First Two Periods"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "DACN") & CompareCell(wsT.Range("V2"), 0) & _
                               ExpectNote(report, "Fixed Coupons 是 First Two Periods") & ExpectNote(report, "推定為 DACN")
    Tally "SG：Fixed Coupons 是 First Two Periods → 推定 DACN 並提醒", fails, passCount, total, detail

    ' SG 的 Quote ? 欄值是 "Strike"：數值列的值不能當成左右一組的標題（#14）
    PasteSGNoProduct wsP
    wsP.Range("R9").Value = "Quote ?"
    PutText wsP.Range("R10"), "Strike"
    wsP.Range("S9").Value = "Coupon p.a."
    wsP.Range("S10").Value = 12
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 80) & CompareCell(wsT.Range("A2"), "FCN")
    fails = fails & ExpectNoNote(report, "Strike 不一致")
    Tally "SG 的 Quote ? 是 Strike → 不當成標題，Strike 照表格、推定 FCN", fails, passCount, total, detail

    ' Super Bull（#19）
    PasteSuperBull wsP, "BNP"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Super Bull") & CompareCell(wsT.Range("E2"), 3) & _
                               CompareCell(wsT.Range("K2"), 70) & CompareCell(wsT.Range("L2"), "Period End") & _
                               CompareCell(wsT.Range("M2"), 70) & CompareCell(wsT.Range("N2"), 6.79) & _
                               CompareCell(wsT.Range("O2"), 97) & CompareCell(wsT.Range("P2"), 6) & _
                               CompareCell(wsT.Range("R2"), "") & CompareCell(wsT.Range("S2"), "") & _
                               CompareCell(wsT.Range("T2"), 0) & CompareCell(wsT.Range("V2"), 0)
    fails = fails & ExpectNoNote(report, "找不到 Barrier Type") & ExpectNoNote(report, "找不到 KI") & _
                    ExpectNoNote(report, "找不到 KO")
    Tally "Super Bull（BNP）：K2、M2 是 Super Bull Level，E2 是 Autocall Frequency，L2 Period End，R2、S2 留空", _
          fails, passCount, total, detail

    PasteSuperBull wsP, "HSBC"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Super Bull") & CompareCell(wsT.Range("D2"), "USD") & _
                               CompareCell(wsT.Range("E2"), 3) & CompareCell(wsT.Range("F2"), "TSM UN") & _
                               CompareCell(wsT.Range("K2"), 83) & CompareCell(wsT.Range("L2"), "Period End") & _
                               CompareCell(wsT.Range("M2"), 90) & CompareCell(wsT.Range("N2"), 5.75) & _
                               CompareCell(wsT.Range("O2"), 97.5) & CompareCell(wsT.Range("P2"), 6) & _
                               CompareCell(wsT.Range("R2"), "") & CompareCell(wsT.Range("S2"), "") & _
                               CompareCell(wsT.Range("T2"), 0) & CompareCell(wsT.Range("V2"), 0) & _
                               CompareCell(wsT.Range("O5"), "2026/10/13") & CompareCell(wsT.Range("O8"), "2027/04/23")
    fails = fails & ExpectNoNote(report, "找不到 Barrier Type")
    Tally "Super Bull（HSBC）：K2 是 Strike、M2 是 Autocall Barrier、N2 是 Bonus Coupon", fails, passCount, total, detail

    ' CA 的主要寫法和 HSBC 相同；不選 C3，由 TDCC 前 3 碼判斷上手
    PasteSuperBull wsP, "CA"
    wsP.Range("C3").ClearContents
    PutText wsP.Range("C5"), "086000000000"
    PutText wsP.Range("A10"), "Super Bull"
    wsP.Range("G10").Value = 1
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Super Bull") & CompareCell(wsT.Range("C2"), "CA") & _
                               CompareCell(wsT.Range("E2"), 3) & CompareCell(wsT.Range("K2"), 83) & _
                               CompareCell(wsT.Range("M2"), 90) & CompareCell(wsT.Range("N2"), 1)
    fails = fails & ExpectNoNote(report, "沒有 %")
    Tally "Super Bull（CA，TDCC 判斷上手）：Autocall Freq、Bonus Coupon；Coupon 1 不當成小數", fails, passCount, total, detail

    PasteSuperBull wsP, "CA"
    PutText wsP.Range("A10"), "Super Bull"
    wsP.Range("E9").Value = "Autocall Date(one-time)"
    PutText wsP.Range("E10"), "3M"
    wsP.Range("G9").Value = "Autocall Flat Coupon (%)"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Super Bull") & CompareCell(wsT.Range("E2"), 3) & _
                               CompareCell(wsT.Range("K2"), 83) & CompareCell(wsT.Range("M2"), 90) & _
                               CompareCell(wsT.Range("N2"), 5.75)
    Tally "Super Bull（CA）：Autocall Date(one-time) 3M → E2 3、Autocall Flat Coupon → N2", fails, passCount, total, detail

    PasteSuperBull wsP, "BNP"
    wsP.Range("E9:E10").ClearContents
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), "") & CompareCell(wsT.Range("M2"), "") & _
                               ExpectNote(report, "信中找不到 Super Bull Level，留空")
    Tally "Super Bull（BNP）沒有 Super Bull Level → K2、M2 留空並提醒", fails, passCount, total, detail

    PasteSuperBull wsP, "HSBC"
    wsP.Range("F9:F10").ClearContents
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 83) & CompareCell(wsT.Range("M2"), "") & _
                               ExpectNote(report, "信中找不到 Autocall Barrier，留空")
    Tally "Super Bull（HSBC）沒有 Autocall Barrier → M2 留空並提醒", fails, passCount, total, detail

    ' BARC 的 DACN／Express 也有 Autocall Barrier：不是 Super Bull 時不讀，不影響 KO Barrier
    PasteBase wsP
    wsP.Range("Q9").Value = "Autocall Barrier (%)"
    wsP.Range("Q10").Value = 95
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("M2"), 100)
    fails = fails & ExpectNoNote(report, "Autocall Barrier")
    Tally "不是 Super Bull 的 Autocall Barrier → 不讀，M2 照 KO Barrier", fails, passCount, total, detail

    PasteBEN wsP, "MS"
    SetTermCell wsT.Range("V2"), Empty
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Regular BEN") & CompareCell(wsT.Range("C2"), "MS") & _
                               CompareCell(wsT.Range("D2"), "USD") & CompareCell(wsT.Range("E2"), 3) & _
                               CompareCell(wsT.Range("F2"), "MU UW") & CompareCell(wsT.Range("K2"), 90) & _
                               CompareCell(wsT.Range("L2"), "Period End") & CompareCell(wsT.Range("M2"), 90) & _
                               CompareCell(wsT.Range("N2"), 20.5) & CompareCell(wsT.Range("O2"), 98) & _
                               CompareCell(wsT.Range("P2"), 3) & CompareCell(wsT.Range("R2"), "NONE") & _
                               CompareCell(wsT.Range("S2"), "") & CompareCell(wsT.Range("V2"), 0)
    fails = fails & ExpectNoNote(report, "找不到 Guaranteed") & ExpectNoNote(report, "找不到 KO") & _
                    ExpectNoNote(report, "找不到 Coupon") & ExpectNoNote(report, "Coupon Barrier")
    Tally "BEN（MS）：E2 = Tenor、L2 Period End、M2 是 Coupon Barrier、N2 是 Bonus Coupon、V2 0", fails, passCount, total, detail

    PasteBEN wsP, "HSBC"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Regular BEN") & CompareCell(wsT.Range("C2"), "HSBC") & _
                               CompareCell(wsT.Range("D2"), "USD") & CompareCell(wsT.Range("E2"), 6) & _
                               CompareCell(wsT.Range("K2"), 85) & CompareCell(wsT.Range("M2"), 85) & _
                               CompareCell(wsT.Range("N2"), 13.5) & CompareCell(wsT.Range("O2"), 97.5) & _
                               CompareCell(wsT.Range("P2"), 6)
    Tally "BEN（HSBC，表頭判斷上手）：CCY、Coupon (Flat)、Issue Price；沒有 Coupon Barrier → M2 = Strike", _
          fails, passCount, total, detail

    PasteBEN wsP, "MS"
    PutText wsP.Range("B10"), "DIGIT WITH LOOKBACK"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Digital BEN") & CompareCell(wsT.Range("V2"), 0)
    Tally "BEN：MS 的 DIGIT WITH LOOKBACK → Digital BEN", fails, passCount, total, detail

    ' SG 的 DCN 有 Tenor (Months)；HSBC 的 Bonus RCN 有 KO Barrier 欄（N/A）、Cost
    PasteBEN wsP, "SG"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Digital BEN") & CompareCell(wsT.Range("P2"), 3) & _
                               CompareCell(wsT.Range("E2"), 3) & CompareCell(wsT.Range("M2"), 88)
    Tally "BEN（SG）：DCN → Digital BEN、Tenor (Months) → P2", fails, passCount, total, detail

    PasteBEN wsP, "HSBC RCN"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Digital BEN") & CompareCell(wsT.Range("E2"), 4) & _
                               CompareCell(wsT.Range("O2"), 97.6)
    fails = fails & ExpectNoNote(report, "KO Barrier")
    Tally "BEN（HSBC）：Bonus RCN → Digital BEN、KO Barrier N/A → E2 = Tenor、Cost → O2", fails, passCount, total, detail

    PasteBEN wsP, "HSBC RCN"
    PutText wsP.Range("G10"), "105%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("E2"), "") & CompareCell(wsT.Range("L2"), "Period End") & _
                               ExpectNote(report, "可能是 BEN with KO，E2 留空")
    Tally "BEN：KO Barrier 有值 → E2 留空並提醒（BEN with KO）", fails, passCount, total, detail

    PasteBEN wsP, "MS"
    PutText wsP.Range("G10"), "95.00%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 90) & CompareCell(wsT.Range("M2"), 95) & _
                               ExpectNote(report, "Coupon Barrier（95）和 Strike（90）不同")
    Tally "BEN：Coupon Barrier 和 Strike 不同 → M2 填 Coupon Barrier 並提醒", fails, passCount, total, detail

    PasteBEN wsP, "BNP"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("A2"), "Regular BEN") & CompareCell(wsT.Range("C2"), "BNP") & _
                               CompareCell(wsT.Range("M2"), 100) & CompareCell(wsT.Range("N2"), 50)
    Tally "BEN（BNP，Strategy Code 判斷上手）", fails, passCount, total, detail

    ' MS 一張表多筆單：只讀第一列並提醒
    PasteBEN wsP, "MS"
    wsP.Range("A10:L10").Copy wsP.Range("A11")
    PutText wsP.Range("C11"), "LITE UW"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("F2"), "MU UW") & ExpectNote(report, "表格有不只一列數值")
    Tally "一張表有多列數值 → 讀第一列並提醒", fails, passCount, total, detail

    ' 第二列數值裡剛好有表頭字樣：不做左右一組（否則 Strike 會變成不一致）
    PasteBEN wsP, "MS"
    wsP.Range("A10:L10").Copy wsP.Range("A11")
    PutText wsP.Range("M11"), "Put Strike"
    PutText wsP.Range("N11"), "50%"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("K2"), 90) & ExpectNoNote(report, "Strike 不一致")
    Tally "一張表的第二列數值不做左右一組", fails, passCount, total, detail

    PasteBEN wsP, "MS"
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNoNote(report, "表格有不只一列數值")
    Tally "只有一列數值 → 不提醒多筆單", fails, passCount, total, detail

    PasteBase wsP
    wsP.Range("Q9").Value = "BBG Code 6"
    PutText wsP.Range("Q10"), "AAPL UW"
    fails = RunFill(report)
    If fails = "" Then fails = ExpectNote(report, "標的超過 5 檔")
    Tally "第 6 檔標的 → 提醒", fails, passCount, total, detail

    PasteBase wsP, False
    PutText wsP.Range("C5"), "029000000000"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("B2"), "029000000000") & CompareCell(wsT.Range("C2"), "BARC")
    Tally "C5 手填 TDCC → B2，並用來判斷上手", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("C5"), "029000000000"
    PutText wsP.Range("A12"), "TDCC code: 029000000001"
    fails = RunFill(report)
    If fails = "" Then fails = CompareCell(wsT.Range("B2"), "029000000000") & _
                               ExpectNote(report, "和信裡的 029000000001 不同")
    Tally "C5 和信裡的 TDCC 不同 → 以 C5 為準並提醒", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("C5"), "12345"
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "12 碼") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "C5 不是 12 碼 → 不填入（V2 不變）", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("C5"), "29000000000"
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "12 碼")
    If ok Then fails = fails & "  應該不填入（11 碼不補 0）" & vbCrLf
    Tally "C5 只有 11 碼 → 不填入", fails, passCount, total, detail

    PasteBase wsP, False
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "判斷不出是哪一家上手") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "判斷不出上手 → 不填入（V2 不變）", fails, passCount, total, detail

    OrderFill_Clear
    SetTermCell wsT.Range("V2"), 5
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "貼上區沒有資料") & CompareCell(wsT.Range("V2"), 5)
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "貼上區沒有資料 → 不填入（V2 不變）", fails, passCount, total, detail

    OrderFill_Clear
    Test_OrderFill_RulesReport = "RULES PASS " & passCount & "/" & total & vbCrLf & detail
End Function

Private Sub PasteBase(ByVal ws As Worksheet, Optional ByVal pickIssuer As Boolean = True)
    Dim headers As Variant, values As Variant

    OrderFill_Clear
    If pickIssuer Then ws.Range("C3").Value = "BARC"
    PutText ws.Range("A7"), "Initial price: Closing price"
    PutText ws.Range("A8"), "PG rate: 0%"

    headers = Array("Product", "Currency", "Guaranteed Periods (m)", "BBG Code 1", "Strike (%)", "KO Type", _
                    "KO Barrier (%)", "Coupon p.a. (%)", "Upfront / NotePrice (%)", "Tenor (m)", "Barrier Type", _
                    "KI Barrier (%)", "Trade Date", "Issue Date", "Final Valuation Date", "Maturity Date")
    values = Array("FCN", "USD", 1, "TSM UN", 80, "Daily Memory", 100, 12, 98, 6, "EKI", 65, _
                   "2026-10-13", "2026-10-20", "2027-04-20", "2027-04-23")
    PutTable ws, headers, values
End Sub

' SG 的寫法：沒有 Product，E 欄 Put Strike、G 欄 AutoCall，Q 欄 Fixed Coupons = All Periods
Private Sub PasteSGNoProduct(ByVal ws As Worksheet)
    PasteBase ws
    ws.Range("C3").Value = "SG"
    ws.Range("A9:A10").ClearContents
    ws.Range("E9").Value = "Put Strike"
    ws.Range("G9").Value = "AutoCall"
    ws.Range("Q9").Value = "Fixed Coupons"
    PutText ws.Range("Q10"), "All Periods"
End Sub

' MS 的 DACN 寫法：Product 是 Range Accrual，E 欄 Put Strike
Private Sub PasteMSRangeAccrual(ByVal ws As Worksheet)
    PasteBase ws
    ws.Range("C3").Value = "MS"
    PutText ws.Range("A10"), "Range Accrual"
    ws.Range("E9").Value = "Put Strike"
End Sub

' Super Bull 的寫法：BNP 是 Super Bull Level（E 欄）；HSBC、CA 是 Strike（D 欄）、
' Autocall Freq（E 欄）、Autocall Barrier（F 欄）、Bonus Coupon（G 欄）
Private Sub PasteSuperBull(ByVal ws As Worksheet, ByVal issuer As String)
    Dim headers As Variant, values As Variant

    OrderFill_Clear
    ws.Range("C3").Value = issuer
    PutText ws.Range("A7"), "Initial price: VWAP"
    PutText ws.Range("A8"), "PG rate: 0%"

    If issuer = "BNP" Then
        headers = Array("Product", "Currency", "BBG Code 1", "Tenor (m)", "Super Bull Level (%)", _
                        "Upfront/Note Price (%)", "Autocall Frequency(M)", "Autocall Coupon Flat (%)", _
                        "Trade Date", "Issue Date", "Redemption Valuation Date", "Redemption Date")
        values = Array("SuperBull", "USD", "AMD UW", 6, 70, 97, 3, 6.79, _
                       "2026-10-13", "2026-10-20", "2027-04-20", "2027-04-23")
    Else
        headers = Array("Product", "Currency", "BBG Code 1", "Strike (%)", "Autocall Freq (m)", _
                        "Autocall Barrier (%)", "Bonus Coupon (%,flat)", "Upfront / NotePrice (%)", "Tenor (m)", _
                        "Strike Date", "Issue Date", "Final Valuation Date", "Maturity Date")
        values = Array("SuperBull", "USD", "TSM UN", 83, 3, 90, 5.75, 97.5, 6, _
                       "2026-10-13", "2026-10-20", "2027-04-20", "2027-04-23")
    End If
    PutTable ws, headers, values
End Sub

' BEN 各家的寫法（不選 C3 的由表頭判斷上手）：
'   MS：B 欄 Product、G 欄 Coupon Barrier；HSBC：沒有 Coupon Barrier；HSBC RCN：G 欄 KO Barrier；
'   SG：Tenor (Months)；BNP：Strategy Code
Private Sub PasteBEN(ByVal ws As Worksheet, ByVal issuer As String)
    Dim headers As Variant, values As Variant

    OrderFill_Clear
    PutText ws.Range("A7"), "Initial price: Closing"
    PutText ws.Range("A8"), "PG rate: 0%"

    Select Case issuer
        Case "MS"
            ws.Range("C3").Value = "MS"
            headers = Array("MS ID", "Product", "Stock 1", "Currency", "Tenor", "Put Strike", "Coupon Barrier", _
                            "Bonus Coupon", "KI Barrier Type", "KI Barrier", "Lookback Period", "Note Price")
            values = Array("123456789", "BEN", "MU UW", "USD", "3m", "90.00%", "90.00%", _
                           "20.50%", "N/A", "N/A", "N/A", "98.00%")
        Case "HSBC"
            headers = Array("Product", "CCY", "Underlying 1", "Underlying 2", "Strike (%)", "Coupon (Flat)", _
                            "Issue Price (%)", "Tenor (m)", "Size Cap", "Comment/Remarks", "ID")
            values = Array("BEN", "USD", "TSM UN", "NVDA UW", "85.00%", "13.5", "97.5%", 6, "N/A", "", "12345")
        Case "HSBC RCN"
            headers = Array("Product", "Currency", "BBG Code 1", "Tenor (m)", "Put Strike (%)", "Bonus Coupon (%, flat)", _
                            "KO Barrier (%)", "PR (%)", "Cost (%)", "Comment/Remarks", "ID")
            values = Array("Bonus RCN", "USD", "AMD UQ", 4, "100%", "8.5%", "N/A", "N/A", "97.6%", "", "12345")
        Case "SG"
            ws.Range("C3").Value = "SG"
            headers = Array("Product", "Underlying 1", "Tenor (Months)", "Currency", "Bonus Coupon (flat)", _
                            "Coupon Barrier", "Put Strike", "KI Type", "KI", "Offer Price", "Participation Rate")
            values = Array("DCN", "AAPL UW", 3, "USD", "12%", "88%", "88%", "N/A", "N/A", "99.5%", "0%")
        Case "BNP"
            headers = Array("Strategy Code", "Product", "Currency", "BBG Code 1", "Tenor (m)", "Put Strike (%)", _
                            "Coupon Barrier (%)", "KI Barrier Type", "KI Barrier (%)", "Bonus Coupon (%, flat)", _
                            "Note Price (%)", "Solve For")
            values = Array("STRATEGY_1", "BEN", "USD", "MU UW", 3, 100, 100, "N/A", "N/A", 50, 96.5, "Issue Price")
    End Select
    PutTable ws, headers, values
End Sub

' Range Accrual 各家的寫法（pickProtected = True 時在 E5 選 Range Accrual）：
'   ESUN：玉山的詢價格式（CA、GS、JPM、SG、Nomura），沒有 Product。貼上後和信裡一樣分成好幾列：
'     第 9 列表頭（D 欄 Coupon、E 欄 Condition_1、F 欄 Target redemption、G 欄 Call Feature、H 欄 UF）、
'     第 10 列表頭說明、第 11 列數值、第 12 列配息的第二行；
'   ESUN1：同上，但 Coupon 的說明和表頭在同一格、配息只有一列；
'   BNP：Product 是 TW DRA，不選 C3 由 Client Ref 判斷上手；BARC：Tenor 5Y、NC 1Y，不選 C3 由 RA Rate 判斷上手
Private Sub PasteRA(ByVal ws As Worksheet, ByVal style As String, Optional ByVal pickProtected As Boolean = True)
    Dim headers As Variant, values As Variant

    OrderFill_Clear
    If pickProtected Then ws.Range("E5").Value = "Range Accrual"
    PutText ws.Range("A8"), "PG rate: 100%"

    Select Case style
        Case "ESUN"
            ws.Range("C3").Value = "CA"
            headers = Array("Tag", "Tenor (Y)", "Currency", "Coupon", "Condition_1", "Target redemption", _
                            "Call Feature", "UF")
            values = Array("E", 10, "USD", "Y1 : 7.00% p.a. fixed", "USD SOFR CMS 10Y " & ChrW(&H2266) & " 5.0%", _
                           "25%", "NC 1Y*3M", "4.50%")
            PutTable ws, headers, values
            ws.Rows(10).Insert
            PutText ws.Range("D10"), "Coupon frequency=Callable frequency"
            PutText ws.Range("F10"), "(No Guaranteed)"
            PutText ws.Range("D12"), "Y2+ : 6.0% x n/N p.a"
            Exit Sub
        Case "ESUN1"
            ws.Range("C3").Value = "CA"
            headers = Array("Tag", "Tenor (Y)", "Currency", "Coupon Coupon frequency=Callable frequency", "Condition_1", _
                            "Call Feature", "UF")
            values = Array("E", 10, "USD", "Y1 : 7.00% p.a. fixed Y2+ : 6.0% x n/N p.a", _
                           "USD SOFR CMS 10Y " & ChrW(&H2266) & " 5.0%", "NC 1Y*3M", "4.50%")
        Case "BNP"
            headers = Array("Client Ref", "Product", "Currency", "Tenor (Y)", "Underlying", "Lower barrier (%)", _
                            "Upper barrier (%)", "Fixed coupon period (Y)", "Fixed coupon (%)", "DRA coupon (%)", _
                            "Upfront/Note Price (%)", "Is Callable", "Non call period (Y)", "Issue Date", "Redemption Date")
            values = Array("REF_1", "TW DRA", "USD", 15, "UCMT10", "NA", 5, 1, 7, 6, 96, "Yes", 1, _
                           "2026-10-20", "2041-10-21")
        Case "BARC"
            headers = Array("Ccy", "Tenor", "NC", "Call Freq", "Fixed Rate", "RA Rate", "Index", "Lower Barrier", _
                            "Upper Barrier", "Reoffer")
            values = Array("USD", "5Y", "1Y", "3M", "7.00%", "5.00%", "CMS10Y", "NA", "5.00%", "96.00%")
    End Select
    PutTable ws, headers, values
End Sub

' 第 9 列寫表頭、第 10 列寫數值（文字用文字格式寫入）
Private Sub PutTable(ByVal ws As Worksheet, ByVal headers As Variant, ByVal values As Variant)
    Dim i As Long

    For i = 0 To UBound(headers)
        ws.Cells(9, i + 1).Value = headers(i)
        If VarType(values(i)) = vbString Then
            PutText ws.Cells(10, i + 1), CStr(values(i))
        Else
            ws.Cells(10, i + 1).Value = values(i)
        End If
    Next i
End Sub

' 直接改 Term 的格子時暫停事件，避免觸發 Term 的 Worksheet_Change（例如跳選 O5）
Private Sub SetTermCell(ByVal cell As Range, ByVal v As Variant)
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    cell.Value = v
    Application.EnableEvents = prevEvents
End Sub

' 寫成文字，避免 Excel 自動轉成日期或數字
Private Sub PutText(ByVal cell As Range, ByVal text As String)
    cell.NumberFormat = "@"
    cell.Value = text
End Sub

Private Function RunFill(ByRef report As String) As String
    If Not OrderFill_Run(True, report) Then RunFill = "  填入失敗：" & Replace(report, vbCrLf, " ") & vbCrLf
End Function

Private Function ExpectNote(ByVal report As String, ByVal text As String) As String
    If InStr(report, text) = 0 Then ExpectNote = "  摘要沒有「" & text & "」：" & Replace(report, vbCrLf, " ") & vbCrLf
End Function

Private Function ExpectNoNote(ByVal report As String, ByVal text As String) As String
    If InStr(report, text) > 0 Then ExpectNoNote = "  摘要不應有「" & text & "」：" & Replace(report, vbCrLf, " ") & vbCrLf
End Function

Private Sub Tally(ByVal name As String, ByVal fails As String, ByRef passCount As Long, _
                  ByRef total As Long, ByRef detail As String)
    total = total + 1
    If fails = "" Then
        passCount = passCount + 1
        detail = detail & "OK   " & name & vbCrLf
    Else
        detail = detail & "FAIL " & name & vbCrLf & fails
    End If
End Sub
