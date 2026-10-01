# 癡心情報上稿

把「癡心情報」電子報推播自動建成 Super 8 Studio 的群發訊息**草稿**（排程當天 21:00）。
你只要填 Excel、跑一個指令。實際發送前照舊自己開草稿檢查、重算人數、按排程。

下面說的「資料夾」都是指這份 README 所在的 `1-自動化腳本\01-癡心情報上稿\`。

---

## 每期上稿（平常就看這段）

1. 打開 `癡心情報上稿.xlsx`，在「上稿」工作表新增一列。欄位規則寫在「說明」工作表。
2. **存檔、關閉 Excel** —— 檔案開著的時候腳本讀不到。
3. **雙擊 `上稿.bat`**。

   腳本會挑「狀態欄是空白的最後一列」來跑。建好草稿後會自動在那一列寫入「狀態」和
   「草稿網址」，之後這列就會被跳過。要重跑某一列，把它的「狀態」清空即可。

4. 終端機會印出草稿的編輯網址。到 Super 8 開草稿確認，之後照平常流程按
   「計算符合條件的客戶數」→「排程」。

只想檢查、不真的建立草稿：在資料夾裡開 PowerShell 跑 `npm run dry`

要指定日期：`node post.js 2026-09-17`

---

## 第一次設定（換電腦或重新 clone 才要做）

1. 安裝 Node.js（LTS 版）：

   ```
   winget install OpenJS.NodeJS.LTS
   ```

   裝完**關掉終端機再重開**，確認 `node --version` 有印出版本。

2. 進到資料夾安裝套件：

   ```
   npm install
   ```

3. 複製 `.env.example`，改名成 `.env`，填入：

   ```
   SUPER8_ORG_ID=阮劇團的 Super 8 組織 ID
   SUPER8_EMAIL=你的 Super 8 登入信箱
   SUPER8_PASSWORD=你的密碼
   ```

   組織 ID 在登入 Super 8 後看網址就有：`console.no8.io/channel-setting/XXXXXXXX/linked-app`
   中間那段。

   腳本會在背景用這組帳密自動登入，登入狀態存在 `browser-profile/`，之後不用每次登入。
   **`.env` 和 `browser-profile/` 都含登入資訊，不要分享或上傳**（已寫進 `.gitignore`）。

   不想把密碼放檔案的話，把 `SUPER8_EMAIL` / `SUPER8_PASSWORD` 兩行刪掉即可，腳本會改成
   開視窗讓你手動登入（`npm run login` 可以先登一次）。`SUPER8_ORG_ID` 一定要留著。

4. 沒有 `癡心情報上稿.xlsx` 的話，跑 `make-template.ps1` 產生空白樣板：

   ```
   powershell -ExecutionPolicy Bypass -File make-template.ps1
   ```

---

## 出問題的時候

| 症狀 | 處理 |
|---|---|
| `缺少 SUPER8_ORG_ID` | `.env` 裡沒填這一行，補上 |
| `找不到 Excel` / 讀不到內容 | 確認檔名沒改，而且 Excel 已經關閉 |
| 一直要我登入 | Super 8 的登入有閒置逾時，久沒用會過期，登一次就好 |
| `找不到 Edge 或 Chrome` | 腳本用內建的 Edge 開登入視窗，Windows 11 都有；沒有的話裝 Chrome |
| Super 8 改版後壞掉 | 看終端機印的錯誤訊息，通常是 API 欄位變了。本機的 `觀察紀錄.md`（不進版控）有完整請求格式可對照 |
| 要改按鈕文字、快速回覆、發送對象標籤 | 改 `config.json`，不要改 `post.js` |

---

## 檔案說明

| 檔案 | 用途 |
|---|---|
| `癡心情報上稿.xlsx` | 每期填的資料（黃底自己填、灰底是公式可覆寫）。不進版控 |
| `config.json` | 固定設定：按鈕文字、上標標籤、快速回覆、發送對象標籤、21:00。不含帳密與組織 ID |
| `post.js` | 腳本本體 |
| `上稿.bat` | 雙擊執行上稿 |
| `make-template.ps1` | 重新產生空白 Excel 樣板（會先備份舊的成 `.bak`） |
| `write-status.ps1` | 腳本用來把「狀態」「草稿網址」寫回 Excel，不用自己跑 |
| `.env.example` | `.env` 的範本 |
| `browser-profile/` | 登入狀態，自動產生。不進版控 |

## 腳本實際做了什麼

1. 讀 Excel 那一列，檢查必填欄位、字數上限、圖片存不存在
2. 用保存的登入狀態拿 session token（過期就開視窗請你重新登入）
3. 把圖片上傳到 Super 8 的圖庫（S3）
4. 用 `config.json` 裡的標籤條件計算符合人數
5. 呼叫建立群發的 API（`isDraft: true`），內容為：一張選項型卡片（圖片、標題、敘述、
   兩顆按鈕）、一則文字訊息、一顆快速回覆、推播通知文字
