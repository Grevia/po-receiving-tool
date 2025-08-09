# Google Apps Script 設定說明

## 步驟 1：創建 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 將專案命名為「開箱工具資料上傳」

## 步驟 2：複製程式碼

1. 刪除預設的程式碼
2. 複製 `google-apps-script.js` 檔案中的程式碼
3. 貼上到 Google Apps Script 編輯器中

## 步驟 3：設定權限

1. 點擊「執行」按鈕
2. 選擇 `testUpload` 函數
3. 授權 Google Apps Script 存取您的 Google Sheets
4. 點擊「允許」

## 步驟 4：部署為 Web App

1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定：
   - 執行身分：自己
   - 存取權限：任何人
4. 點擊「部署」
5. 複製產生的 Web App URL

## 步驟 5：更新 JavaScript 檔案

1. 開啟 `receiving-confirm.js`
2. 找到第 456 行：
   ```javascript
   const webAppUrl = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL';
   ```
3. 將 `YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL` 替換為您的 Web App URL

## 步驟 6：測試功能

1. 重新啟動伺服器
2. 測試序號掃描和上傳功能
3. 檢查 Google Sheets 是否收到資料

## 注意事項

- 確保 Google Sheets 的 `receiving_confirm` 工作表存在
- 確保工作表有適當的欄位標題
- 測試時建議先使用少量資料
- 如果遇到權限問題，請檢查 Google Apps Script 的執行權限
