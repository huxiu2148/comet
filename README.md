# 🌠 Comet 訂單查詢與回報系統

這是一個基於 **Google Apps Script (GAS)** 開發的輕量化訂單管理系統，連結 Google Sheets 作為後端資料庫，提供使用者查詢訂單狀態、查看物流資訊，以及回報匯款末五碼或託運單號的功能。

## 🚀 核心功能
- **多模式查詢**：支援多種訂單模式（匯款、集運、賣場跳轉等）。
- **雙按鈕操作**：在歷史紀錄頁面提供「查看資訊」與「快速回報」獨立按鈕，優化使用者體驗。
- **智慧記憶**：自動記錄使用者輸入過的 Email、聯絡帳號，減少重複填寫。
- **防呆機制**：
  - 匯款末五碼必須為 5 碼數字。
  - 韓國託運單號必須為 12 碼數字。
- **UI 優化**：
  - 採用 Tailwind CSS 打造精美卡片式介面。
  - 支援 Modal 層級優化（Z-index），確保視窗彈出時不會互相遮擋。

## 🛠️ 技術棧
- **Frontend**: HTML5, Tailwind CSS, JavaScript (Vanilla JS)
- **Backend**: Google Apps Script
- **Database**: Google Sheets

## 📂 檔案架構
- `Code.gs`: 處理後端邏輯、Google Sheets 讀寫及 Email 發送。
- `Index.html`: 系統主介面、CSS 樣式及前端互動邏輯（JavaScript）。

## ⚙️ 安裝與部署
1. 將 `Code.gs` 與 `Index.html` 貼入 Google Apps Script 專案。
2. 於 `Code.gs` 中設定 `ADMIN_EMAIL` 與 `API_KEY`。
3. 部署為 Web App，並將網址提供給使用者。

---

## 🀄 台灣麻將記帳系統

`mahjong-scorekeeper/index.html` 是一個獨立的靜態網頁小工具，用於朋友間打麻將時記帳算錢。

- **設定區**：自訂 4 位玩家名稱、底 / 台金額。
- **記分操作區**：大按鈕選擇莊家（含連莊）、贏家、自摸／放槍（放槍者）、台數，即時預覽金額。
- **計算邏輯**：放槍由放槍者賠贏家；自摸由其餘三家均攤賠給贏家；莊家連莊 n 則在台數上額外 +2n。
- **即時戰況表**：四人目前總輸贏（綠色贏／紅色輸），下方為歷史對戰紀錄，可「刪除上一筆」修正記錯。
- 純 HTML/CSS/JavaScript，無需安裝，直接用瀏覽器開啟 `mahjong-scorekeeper/index.html` 即可使用，適合手機觸控操作。
- **雲端同步**：透過發布為 Claude Artifact 開啟時，每次記錄/刪除/設定變更都會即時同步到雲端，任何裝置打開同一個網址都會看到最新戰況；若不是透過該網址開啟（例如直接下載此檔案在本機打開），則自動退回只存在瀏覽器 `localStorage`（僅該裝置本機保存）。

---
*本專案僅供個人或內部團隊使用，請務必將 GitHub 倉庫設為 **Private** 以保護 API Key 與資料安全。*
