param(
  [Parameter(Mandatory)][string]$Workbook,
  [Parameter(Mandatory)][string]$HtmlDir,
  [Parameter(Mandatory)][string]$OutFile,
  [string]$Root
)
# PowerShell 5.1 的參數預設值裡 $PSScriptRoot 是空的，所以在這裡才指定預設的 repo 根目錄
if (-not $Root) { $Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path }
# 在活頁簿「複本」匯入 src/vba/OrderFill.bas，把 extract.py 抽出的每封 .htm 貼到「起單回填」A7、
# 按「填入」，再讀回 Term!A2:V2 與 O5:O9。原活頁簿不會被修改。
$run = Join-Path $env:TEMP ("orderfill-check-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $run | Out-Null
$copy = Join-Path $run "copy.xlsm"
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
    $hw = $x.Workbooks.Open($h.FullName)
    [void]$hw.Worksheets.Item(1).UsedRange.Copy($wsP.Range("A7"))
    $hw.Close($false)
    $report = $x.Run("'copy.xlsm'!Check_Fill")
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
