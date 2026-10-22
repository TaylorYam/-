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
' 貼上區的圖片要刪掉，按鈕與 C3 下拉清單要保留。
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
    Tally "SG：Fixed Coupons 不是 All Periods → A2 留空並提醒", fails, passCount, total, detail

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
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "12 碼")
    If ok Then fails = fails & "  應該不填入" & vbCrLf
    Tally "C5 不是 12 碼 → 不填入", fails, passCount, total, detail

    PasteBase wsP
    PutText wsP.Range("C5"), "29000000000"
    ok = OrderFill_Run(True, report)
    fails = ExpectNote(report, "12 碼")
    If ok Then fails = fails & "  應該不填入（11 碼不補 0）" & vbCrLf
    Tally "C5 只有 11 碼 → 不填入", fails, passCount, total, detail

    OrderFill_Clear
    Test_OrderFill_RulesReport = "RULES PASS " & passCount & "/" & total & vbCrLf & detail
End Function

Private Sub PasteBase(ByVal ws As Worksheet, Optional ByVal pickIssuer As Boolean = True)
    Dim headers As Variant, values As Variant
    Dim i As Long

    OrderFill_Clear
    If pickIssuer Then ws.Range("C3").Value = "BARC"
    PutText ws.Range("A7"), "Initial price: Closing price"
    PutText ws.Range("A8"), "PG rate: 0%"

    headers = Array("Product", "Currency", "Guaranteed Periods (m)", "BBG Code 1", "Strike (%)", "KO Type", _
                    "KO Barrier (%)", "Coupon p.a. (%)", "Upfront / NotePrice (%)", "Tenor (m)", "Barrier Type", _
                    "KI Barrier (%)", "Trade Date", "Issue Date", "Final Valuation Date", "Maturity Date")
    values = Array("FCN", "USD", 1, "TSM UN", 80, "Daily Memory", 100, 12, 98, 6, "EKI", 65, _
                   "2026-10-13", "2026-10-20", "2027-04-20", "2027-04-23")
    For i = 0 To UBound(headers)
        ws.Cells(9, i + 1).Value = headers(i)
        If VarType(values(i)) = vbString Then
            PutText ws.Cells(10, i + 1), CStr(values(i))
        Else
            ws.Cells(10, i + 1).Value = values(i)
        End If
    Next i
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
