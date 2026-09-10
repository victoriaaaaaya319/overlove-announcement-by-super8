# 用電腦上的 Excel 產生「癡心情報上稿.xlsx」樣板
# 執行：powershell -ExecutionPolicy Bypass -File make-template.ps1
# 若檔案已存在會先備份成 .bak 再重建

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$out  = Join-Path $root '癡心情報上稿.xlsx'

if (Test-Path $out) { Copy-Item $out ($out + '.bak') -Force; Remove-Item $out -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

# ---------- 工作表 1：上稿 ----------
$ws = $wb.Worksheets.Item(1)
$ws.Name = '上稿'

$headers = @('發送日期','主題','卡片標題','推播通知','卡片敘述','圖片路徑','電子報網址','文字訊息','備註','狀態','草稿網址')
for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $headers[$i] }

$hdr = $ws.Range('A1:K1')
$hdr.Font.Bold = $true
$hdr.Interior.Color = 0x4F4F4F   # 深灰
$hdr.Font.Color = 0xFFFFFF
$hdr.HorizontalAlignment = -4108  # center

$maxRow = 200
# 輸入欄位：黃底
$inputCols = @('A','B','E','F','G','H','I')
foreach ($c in $inputCols) { $ws.Range("$($c)2:$($c)$maxRow").Interior.Color = 0xCCFFFF }  # 淡黃（BGR）
# 公式欄位：淡灰底、黑字
$ws.Range("C2:D$maxRow").Interior.Color = 0xF2F2F2
# 腳本寫回的欄位（狀態、草稿網址）：淡綠底，不要手動填
$ws.Range("J2:K$maxRow").Interior.Color = 0xE0F0E0

# 公式：卡片標題 / 推播通知 由主題組成（可直接覆寫）
$ws.Range("C2:C$maxRow").Formula = '=IF(B2="","","癡心情報｜"&B2)'
$ws.Range("D2:D$maxRow").Formula = '=IF(B2="","","叮咚！癡心情報來囉！"&B2)'

# 範例列（2026-09-10 實際上稿內容）
$ws.Cells.Item(2, 1).Value2 = [double]((Get-Date '2026-09-10').ToOADate())
$ws.Cells.Item(2, 2).Value2 = '市場疊放陳列生活的滋味'
$ws.Cells.Item(2, 5).Value2 = '看著大家帶著不同需求來到市場，但帶著比預期更多的大包小包從四面八方離開，也是逛市場的一件樂事'
$ws.Cells.Item(2, 6).Value2 = 'C:\Users\USER\Downloads\未命名設計.png'
$ws.Cells.Item(2, 7).Value2 = 'https://clt1443557.bmeurl.co/143A5931'
$ws.Cells.Item(2, 8).Value2 = "是說……是不是還有工友不知道……`n`n《癡心情報》是一個可以跟我們「聊天」的電子報 (ŏ_ŏ)`n`n只要點擊下面按鈕，從表單填寫想對我們說的話，阮編就會收到你的回覆唷！"
$ws.Cells.Item(2, 9).Value2 = '範例列（2026-09-10 實際上稿）。可以整列刪掉。'

# 格式
$ws.Range("A2:A$maxRow").NumberFormat = 'yyyy/mm/dd'
$ws.Range("A2:A$maxRow").HorizontalAlignment = -4108
$ws.Range("H2:H$maxRow").WrapText = $true
$ws.Range("E2:E$maxRow").WrapText = $true
$ws.Range("A1:K$maxRow").VerticalAlignment = -4160  # top
$ws.Range("A1:K$maxRow").Font.Name = 'Microsoft JhengHei'
$ws.Range("A1:K$maxRow").Font.Size = 11

$widths = @{ 'A'=12; 'B'=26; 'C'=34; 'D'=40; 'E'=44; 'F'=44; 'G'=40; 'H'=56; 'I'=24; 'J'=26; 'K'=60 }
foreach ($k in $widths.Keys) { $ws.Columns.Item($k).ColumnWidth = $widths[$k] }
$ws.Rows.Item(2).RowHeight = 110

# 凍結標題列
$ws.Activate()
$xl.ActiveWindow.SplitRow = 1
$xl.ActiveWindow.SplitColumn = 0
$xl.ActiveWindow.FreezePanes = $true

# ---------- 工作表 2：說明 ----------
$ws2 = $wb.Worksheets.Add([Type]::Missing, $ws)
$ws2.Name = '說明'
$lines = @(
  '癡心情報上稿樣板 — 使用說明',
  '',
  '1. 在「上稿」工作表新增一列，一列 = 一期。黃底的欄位要自己填，灰底是公式（可以直接覆寫）。',
  '2. 發送日期：填日期即可，腳本固定排程當天 21:00，群發名稱自動變成 YYMMDD-癡心情報推播。',
  '3. 主題：只填主題本身，例如「市場疊放陳列生活的滋味」。卡片標題、推播通知會自動組出來。',
  '4. 卡片標題（上限 40 字）/ 推播通知（上限 100 字，建議 35 字內）：想改就直接在格子裡打字覆蓋公式。',
  '5. 卡片敘述：上限 60 字。',
  '6. 圖片路徑：圖片檔的完整路徑，例如 C:\Users\USER\Downloads\封面.png（支援 png / jpg / gif）。',
  '7. 電子報網址：卡片「馬上閱讀電子報」按鈕要開的連結，必須 http 或 https 開頭。',
  '8. 文字訊息：卡片後面那則文字。儲存格內換行請用 Alt+Enter，空行就是連按兩次。',
  '9. 備註：自己用，腳本不會讀。',
  '10. 狀態 / 草稿網址（綠底）：腳本建好草稿後自動寫入，不用手填。「狀態」有字的列會被跳過；要重跑就把那格清空。',
  '',
  '執行：在資料夾開終端機，輸入 npm run post（用最後一列）或 node post.js 2026-09-10（指定日期）。',
  '先看不送：npm run dry',
  '',
  '固定不變的東西（兩顆按鈕文字、上標標籤、快速回覆、發送對象標籤、21:00）在 config.json，要改再改那裡。'
)
for ($i = 0; $i -lt $lines.Count; $i++) { $ws2.Cells.Item($i + 1, 1).Value2 = $lines[$i] }
$ws2.Cells.Item(1, 1).Font.Bold = $true
$ws2.Cells.Item(1, 1).Font.Size = 14
$ws2.Columns.Item(1).ColumnWidth = 110
$ws2.Range('A1:A30').Font.Name = 'Microsoft JhengHei'

$ws.Activate()
$ws.Range('A2').Select() | Out-Null

$wb.SaveAs($out, 51)  # 51 = xlOpenXMLWorkbook
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws2) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)  | Out-Null

Write-Output "OK: $out"
