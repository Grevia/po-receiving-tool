# Google Apps Script 設定指南

## 問題說明
如果您在到貨確認頁面上傳資料後，Google Sheets 沒有資料寫入，這是因為 Google Apps Script Web App 尚未正確設定。

## 解決步驟

### 步驟 1：創建 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 將專案命名為「開箱工具資料上傳」

### 步驟 2：複製程式碼

1. 刪除預設的程式碼
2. 複製 `google-apps-script.js` 檔案中的程式碼
3. 貼上到 Google Apps Script 編輯器中

### 步驟 3：設定權限

1. 點擊「執行」按鈕
2. 選擇 `testUpload` 函數
3. 授權 Google Apps Script 存取您的 Google Sheets
4. 點擊「允許」

### 步驟 4：部署為 Web App

1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定：
   - 執行身分：自己
   - 存取權限：任何人
4. 點擊「部署」
5. **複製產生的 Web App URL**（重要！）

### 步驟 5：更新 JavaScript 檔案

1. 開啟 `receiving-confirm.js`
2. 找到第 502 行：
   ```javascript
   const webAppUrl = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL';
   ```
3. 將 `YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL` 替換為您的 Web App URL

### 步驟 6：測試功能

1. 重新啟動伺服器
2. 測試序號掃描和上傳功能
3. 檢查 Google Sheets 是否收到資料

## 常見問題

### Q: 為什麼會顯示「模擬上傳成功」？
A: 這是因為 `receiving-confirm.js` 中的 `webAppUrl` 還是預設值。請按照步驟 5 更新 URL。

### Q: 如何檢查 Web App URL 是否正確？
A: 在瀏覽器中開啟 Web App URL，應該會看到一個 JSON 回應。

### Q: 權限錯誤怎麼辦？
A: 確保 Google Apps Script 專案有權限存取您的 Google Sheets，並且 Web App 的存取權限設為「任何人」。

### Q: 如何確認資料是否成功寫入？
A: 檢查 Google Sheets 的 `receiving_confirm` 工作表，應該會看到新的資料行。

## 測試方法

1. 在到貨確認頁面輸入測試資料
2. 點擊「上傳資料」
3. 檢查瀏覽器控制台是否有錯誤訊息
4. 檢查 Google Sheets 是否有新資料

## 聯絡支援

如果仍然有問題，請檢查：
- 瀏覽器控制台的錯誤訊息
- Google Apps Script 的執行記錄
- Google Sheets 的權限設定
