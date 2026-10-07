# 由 post.js 呼叫：把「狀態」「草稿網址」寫回 Excel（用 Excel COM，保留原本格式）
param(
  [Parameter(Mandatory=$true)][string]$File,
  [Parameter(Mandatory=$true)][string]$Sheet,
  [Parameter(Mandatory=$true)][int]$Row,
  [Parameter(Mandatory=$true)][string]$Status,
  [string]$Url = ''
)
$ErrorActionPreference = 'Stop'
$xl = $null; $wb = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($File)
  $ws = $wb.Worksheets.Item($Sheet)

  # 依標題列找欄位；沒有就補在最後面
  $lastCol = $ws.UsedRange.Columns.Count
  $statusCol = 0; $urlCol = 0
  for ($c = 1; $c -le $lastCol; $c++) {
    $h = [string]$ws.Cells.Item(1, $c).Text
    if ($h -eq '狀態') { $statusCol = $c }
    if ($h -eq '草稿網址') { $urlCol = $c }
  }
  if ($statusCol -eq 0) { $statusCol = $lastCol + 1; $ws.Cells.Item(1, $statusCol).Value2 = '狀態'; $lastCol = $statusCol }
  if ($urlCol -eq 0)    { $urlCol = $lastCol + 1;    $ws.Cells.Item(1, $urlCol).Value2 = '草稿網址' }

  $ws.Cells.Item($Row, $statusCol).Value2 = $Status
  if ($Url -ne '') { $ws.Cells.Item($Row, $urlCol).Value2 = $Url }

  $wb.Save()
  Write-Output "OK row=$Row statusCol=$statusCol"
} finally {
  if ($wb) { $wb.Close($false) }
  if ($xl) { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
}
