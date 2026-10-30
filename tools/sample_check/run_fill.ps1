param(
  [Parameter(Mandatory)][string]$Workbook,
  [Parameter(Mandatory)][string]$HtmlDir,
  [Parameter(Mandatory)][string]$OutFile,
  [string]$Root,
  [string]$ProtectedProduct,
  [string]$Manifest
)
# PowerShell 5.1 的參數預設值裡 $PSScriptRoot 是空的，所以在這裡才指定預設的 repo 根目錄
if (-not $Root) { $Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path }
# 在活頁簿「複本」匯入 src/vba/OrderFill.bas，把 extract.py 抽出的每封 .htm 貼到「起單回填」A7、
# 按「填入」，再讀回 Term!A2:V2 與 O5:O9。原活頁簿不會被修改。
# 保本型商品（例：Range Accrual）的樣本要用 -ProtectedProduct 指定，每封信都先在 E5 選它再填入。
# 指定 -Manifest（extract.py 的 manifest.json）時，判斷不出上手的信依檔名 TDCC 的上手在 C3 選好再填一次，
# 和使用者手動選上手的操作相同；摘要最後會加一行說明。
$run = Join-Path $env:TEMP ("orderfill-check-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $run | Out-Null
$copy = Join-Path $run "copy.xlsm"
$issuerOf = @{}
if ($Manifest) {
  foreach ($r in (Get-Content -Raw -Encoding UTF8 -LiteralPath $Manifest | ConvertFrom-Json)) {
    if ($r.issuer) { $issuerOf[$r.html] = $r.issuer }
  }
}
Copy-Item -LiteralPath $Workbook -Destination $copy

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 1
$out = @()
try {
  $wb = $x.Workbooks.Open($copy)
  try { $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item("OrderFill")) } catch {}
  [void]$wb.VBProject.VBComponents.Import((Join-Path $Root "src\vba\OrderFill.bas"))
  # Application.Run 拿不到 ByRef 的摘要，加一個只存在於複本的包裝函式
  $helper = $wb.VBProject.VBComponents.Add(1)
  $helper.CodeModule.AddFromString("Public Function Check_Fill() As String`r`n Dim r As String`r`n OrderFill_Run True, r`r`n Check_Fill = r`r`nEnd Function")
  $wsP = $wb.Worksheets.Item("起單回填")
  $wsT = $wb.Worksheets.Item("Term")
  foreach ($h in (Get-ChildItem $HtmlDir -Filter "*.htm" | Sort-Object Name)) {
    $x.Run("'copy.xlsm'!OrderFill_Clear")
    if ($ProtectedProduct) { $wsP.Range("E5").Value2 = $ProtectedProduct }
    $hw = $x.Workbooks.Open($h.FullName)
    [void]$hw.Worksheets.Item(1).UsedRange.Copy($wsP.Range("A7"))
    $hw.Close($false)
    $report = $x.Run("'copy.xlsm'!Check_Fill")
    if ($report -like "*判斷不出是哪一家上手*" -and $issuerOf.ContainsKey($h.BaseName)) {
      $wsP.Range("C3").Value2 = $issuerOf[$h.BaseName]
      $report = $x.Run("'copy.xlsm'!Check_Fill") + "`r`n- 判斷不出上手，依 manifest 在 C3 選 " + $issuerOf[$h.BaseName] + " 後重填"
    }
    $row = @()
    foreach ($c in "A2","B2","C2","D2","E2","F2","G2","H2","I2","J2","K2","L2","M2","N2","O2","P2","Q2","R2","S2","T2","V2","O5","O6","O7","O8","O9") {
      $row += ("{0}={1}" -f $c, ($wsT.Range($c).Text -replace "[`r`n]+", " "))
    }
    $out += "=== " + $h.BaseName
    $out += ($row -join " | ")
    $out += $report
  }
  $wb.Close($false)
} catch {
  $out += "ERROR: " + $_.Exception.Message
} finally {
  try { $x.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $run -ErrorAction SilentlyContinue
}
$out -join "`r`n" | Out-File -FilePath $OutFile -Encoding utf8
