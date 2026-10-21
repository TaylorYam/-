Attribute VB_Name = "OrderFillTest"
Option Explicit

'==================================================
' 起單回填測試：把「起單信件格式.xlsx」14 張上手表各自貼進「起單回填」並執行填入，
' 再和 ELN每日起單建檔中同一 TDCC 的紀錄比對。
' 會覆蓋 Term 與貼上區，請在複本上執行。
'==================================================

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

    result = Test_OrderFill_Report(path)
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
