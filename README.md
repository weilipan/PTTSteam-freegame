# PTT Steam 限免遊戲查詢工具

## 📌 功能介紹

本 Google Apps Script 專案可自動爬取 PTT Steam 看板的 `[限免]` 遊戲資訊，並將結果記錄於 Google Sheets，超過 2 週的舊紀錄將自動刪除。此外，當發現新遊戲時，系統會自動發送 Email 通知。

## 🚀 主要功能

- **📄 建立空白表格**：初始化 Google Sheets，包含 `PTTSteam Tracking`（記錄限免遊戲）與 `Settings`（存放 Email 設定與 `maxPages`）。
- **✉️ 設定 Email**：用戶可輸入 Email 接收新發現的 `[限免]` 遊戲通知。
- **🔄 設定 maxPages**：設定爬取的最大頁數。
- **🔍 查詢並寄送相關資訊**：自動爬取 PTT Steam 看板，篩選 `[限免]` 文章，並發送 Email 通知。
- **🗑️ 自動清理舊資料**：每日執行時，會刪除 `PTTSteam Tracking` 表格內 **超過 2 週** 的記錄，確保 Google Sheets 不會變得過於臃腫。

## 📂 安裝與設定

1. **打開 Google Sheets**
2. \*\*點擊 \*\***`擴充功能 > Apps Script`**
3. **貼上 ********`Google Apps Script`******** 程式碼**（見 `script.js`）
4. **執行 \*\*\*\*****`onOpen()`**，初始化 Google Sheets 選單
5. **執行 \*\*\*\*****`initializeSheet()`**，初始化 `PTTSteam Tracking` 和 `Settings`
6. **設定每日觸發器**，讓 `checkPTTSteamFreeGames()` 自動執行

## 📌 使用方式

### 1️⃣ **建立空白表格**

- 點擊 `PTT Steam 限免遊戲查詢 > 📄 建立空白表格`
- 這將建立兩個工作表：
  - `PTTSteam Tracking`（記錄限免遊戲）
  - `Settings`（存放 Email 設定與 `maxPages`）

### 2️⃣ **設定 Email 通知**

- 點擊 `PTT Steam 限免遊戲查詢 > ✉️ 設定 Email`
- 輸入 Email 地址，該地址將用來接收 `[限免]` 遊戲通知

### 3️⃣ **設定爬取頁數**

- 點擊 `PTT Steam 限免遊戲查詢 > 🔄 設定 maxPages`
- 輸入要爬取的最大頁數（預設為 5 頁）

### 4️⃣ **手動查詢並發送 Email 通知**

- 點擊 `PTT Steam 限免遊戲查詢 > 🔍 查詢並寄送相關資訊`
- 這將爬取 PTT Steam 看板，篩選 `[限免]` 文章，並發送 Email 通知

## 🔄 **設定自動執行**

1. **打開 ********`Apps Script`******** 編輯器**
2. **點擊左側 ********`鬧鐘圖示`********（觸發器 Triggers）**
3. \*\*點擊 \*\***`+ 新增觸發器`**
   - **選擇函式**：`checkPTTSteamFreeGames`
   - **選擇觸發條件**：時間驅動 (Time-driven)
   - **頻率**：每天
4. \*\*按 \*\***`儲存`**

## 🛠 **技術細節**

### ✅ **自動刪除超過 2 週的資料**

- `cleanupOldEntries(trackingSheet)` 會檢查 `PTTSteam Tracking` 表內的 `Send Time`
- 若 `Send Time` 超過 **14 天**，該行資料將被刪除
- 避免 Google Sheets 變得過於臃腫，提高運行效率

### ✅ **爬取 ********`[限免]`******** 文章並發送 Email**

- **每次爬取 ********`maxPages`******** 頁**（預設 5 頁）
- **篩選標題內含 ********`[限免]`******** 的文章**
- **發現新 ********`[限免]`******** 遊戲時，記錄至 ********`PTTSteam Tracking`********，並發送 Email 通知**

### ✅ **自動翻頁**

- **尋找 ********`上頁`******** 按鈕 (********`<a class="btn wide" href="/bbs/Steam/index4000.html">&lsaquo; 上頁</a>`********)**
- **確保爬取 ********`maxPages`******** 頁，或直到沒有 ********`上頁`******** 按鈕時停止**

### ✅ **避免 ********`getUi()`******** 問題**

- **若 ********`SpreadsheetApp.getUi()`******** 因未開啟 Google Sheet 而發生錯誤**，則使用 `Logger.log()` 記錄錯誤，而不執行 `alert()`

## 🚀 **結論**

透過本 Google Apps Script，你可以 **每日自動追蹤 PTT Steam 看板的 ********`[限免]`******** 遊戲資訊**，並 **自動清理舊紀錄**，確保 Google Sheets 保持整潔。

📩 **現在，你可以輕鬆獲得最新的限免遊戲資訊，而不必手動查詢！🎮**

