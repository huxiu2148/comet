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
*本專案僅供個人或內部團隊使用，請務必將 GitHub 倉庫設為 **Private** 以保護 API Key 與資料安全。*
