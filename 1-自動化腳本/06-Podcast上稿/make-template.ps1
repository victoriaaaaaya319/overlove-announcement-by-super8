# 用電腦上的 Excel 產生「Podcast上稿.xlsx」樣板
# 執行：powershell -ExecutionPolicy Bypass -File make-template.ps1
# 若檔案已存在會先備份成 .bak 再重建

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$out  = Join-Path $root 'Podcast上稿.xlsx'

if (Test-Path $out) { Copy-Item $out ($out + '.bak') -Force; Remove-Item $out -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

# ---------- 工作表 1：上稿 ----------
$ws = $wb.Worksheets.Item(1)
$ws.Name = '上稿'

$headers = @('發送日期','發送時間','標題','卡片標題','推播通知','卡片敘述','圖片路徑','Apple Podcast 網址','Spotify 網址','卡片前文字','卡片後自填文字','快速回覆顯示文字','備註','狀態','草稿網址')
for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $headers[$i] }

$hdr = $ws.Range('A1:O1')
$hdr.Font.Bold = $true
$hdr.Interior.Color = 0x4F4F4F   # 深灰
$hdr.Font.Color = 0xFFFFFF
$hdr.HorizontalAlignment = -4108  # center

$maxRow = 200
# 輸入欄位：黃底
$inputCols = @('A','B','C','F','G','H','I','J','K','L','M')
foreach ($c in $inputCols) { $ws.Range("$($c)2:$($c)$maxRow").Interior.Color = 0xCCFFFF }  # 淡黃（BGR）
# 公式欄位：淡灰底、黑字
$ws.Range("D2:E$maxRow").Interior.Color = 0xF2F2F2
# 腳本寫回的欄位（狀態、草稿網址）：淡綠底，不要手動填
$ws.Range("N2:O$maxRow").Interior.Color = 0xE0F0E0

# 發送時間用文字格式，打 21:00 就是 21:00
$ws.Range("B2:B$maxRow").NumberFormat = '@'
$ws.Range("B2:B$maxRow").HorizontalAlignment = -4108

# 公式：卡片標題 / 推播通知 由標題組成（可直接覆寫）
$ws.Range("D2:D$maxRow").Formula = '=IF(C2="","",C2)'
$ws.Range("E2:E$maxRow").Formula = '=IF(C2="","","🎧【這聲好啊！】 "&C2)'

# 範例列
$ws.Cells.Item(2, 1).Value2 = [double]((Get-Date '2026-10-08').ToOADate())
$ws.Cells.Item(2, 2).Value2 = '21:00'
$ws.Cells.Item(2, 3).Value2 = '範例標題'
$ws.Cells.Item(2, 6).Value2 = '這一集的簡短介紹（可留空）'
$ws.Cells.Item(2, 7).Value2 = 'C:\Users\USER\Downloads\封面.jpg'
$ws.Cells.Item(2, 8).Value2 = 'https://podcasts.apple.com/'
$ws.Cells.Item(2, 9).Value2 = 'https://open.spotify.com/'
$ws.Cells.Item(2, 10).Value2 = "卡片上面那則文字訊息（本集介紹）。`n`n換行用 Alt+Enter。"
$ws.Cells.Item(2, 11).Value2 = '接在固定公版後面的補充，可留空。'
$ws.Cells.Item(2, 12).Value2 = '聽完寫心得，還能拿小心點！'
$ws.Cells.Item(2, 13).Value2 = '範例列，可以整列刪掉。'
$ws.Cells.Item(2, 14).Value2 = '範例（不會跑）'

# 格式
$ws.Range("A2:A$maxRow").NumberFormat = 'yyyy/mm/dd'
$ws.Range("A2:A$maxRow").HorizontalAlignment = -4108
$ws.Range("F2:F$maxRow").WrapText = $true
$ws.Range("J2:K$maxRow").WrapText = $true
$ws.Range("A1:O$maxRow").VerticalAlignment = -4160  # top
$ws.Range("A1:O$maxRow").Font.Name = 'Microsoft JhengHei'
$ws.Range("A1:O$maxRow").Font.Size = 11

$widths = @{ 'A'=12; 'B'=10; 'C'=26; 'D'=34; 'E'=34; 'F'=40; 'G'=44; 'H'=40; 'I'=40; 'J'=56; 'K'=44; 'L'=34; 'M'=24; 'N'=26; 'O'=60 }
foreach ($k in $widths.Keys) { $ws.Columns.Item($k).ColumnWidth = $widths[$k] }
$ws.Rows.Item(2).RowHeight = 70

# 凍結標題列
$ws.Activate()
$xl.ActiveWindow.SplitRow = 1
$xl.ActiveWindow.SplitColumn = 0
$xl.ActiveWindow.FreezePanes = $true

# ---------- 工作表 2：說明 ----------
$ws2 = $wb.Worksheets.Add([Type]::Missing, $ws)
$ws2.Name = '說明'
$lines = @(
  'Podcast《這聲好啊！》上稿樣板 — 使用說明',
  '',
  '1. 在「上稿」工作表新增一列，一列 = 一期。黃底的欄位要自己填，灰底是公式（可以直接覆寫）。',
  '2. 發送日期：填日期。群發名稱自動變成 YYMMDD-這聲好啊｜標題。',
  '3. 發送時間：填像 21:00 這樣的時間。留空就是 21:00。可以先只填日期，時間之後再補。',
  '4. 標題：只填標題本身。卡片標題 = 標題；推播通知 = 🎧【這聲好啊！】 + 標題。',
  '5. 卡片標題（上限 40 字）/ 推播通知（上限 100 字，建議 35 字內）：想改就直接在格子裡打字覆蓋公式。',
  '6. 卡片敘述：卡片裡標題下面那行小字，上限 60 字，可留空。',
  '7. 圖片路徑：正方形圖片的完整路徑，例如 C:\Users\USER\Downloads\封面.jpg（支援 png / jpg / gif）。',
  '8. Apple Podcast 網址 / Spotify 網址：卡片前兩顆按鈕要開的連結，必須 http 或 https 開頭。',
  '9. 卡片前文字：卡片「上面」那則文字（本集介紹）。儲存格內換行請用 Alt+Enter，空行就是連按兩次。',
  '10. 卡片後自填文字：卡片「下面」那則文字是固定公版（訂閱提醒＋小心得換小心點）；這格有填的話，會用 ♡♡♡♡♡♡♡♡♡♡ 隔開接在公版後面。可留空。',
  '11. 快速回覆顯示文字：必填，上限 20 字。按下去回傳的內容固定是「《這聲好啊！》小心得換小心點」。',
  '12. 備註：自己用，腳本不會讀。',
  '13. 狀態 / 草稿網址（綠底）：腳本建好草稿後自動寫入，不用手填。「狀態」有字的列會被跳過；要重跑就把那格清空。',
  '',
  '執行：存檔、關閉 Excel，雙擊「上稿.bat」（用狀態空白的最後一列）。',
  '先看不送：npm run dry',
  '',
  '發送對象：腳本不設定，草稿建好後自己到 Super 8 選對象、計算人數、按排程。',
  '固定不變的東西（三顆按鈕文字與標籤、卡片後的固定公版、快速回覆回傳內容、預設時間 21:00）在 config.json，要改再改那裡。'
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
