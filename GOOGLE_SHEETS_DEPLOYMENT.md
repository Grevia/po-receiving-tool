# Google Sheets 部署說明

## 📋 部署步驟

### 步驟 1：創建 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 專案名稱輸入：「開箱工具系統」
4. 刪除預設的程式碼

### 步驟 2：複製程式碼

1. 將 `google-sheets-upload.js` 的內容複製到 Google Apps Script 編輯器
2. 點擊「儲存」按鈕

### 步驟 3：設定權限

1. 點擊「執行」按鈕
2. 選擇 `testFunctions` 函數
3. 點擊「執行」
4. 授權存取您的 Google 帳戶
5. 選擇「進階」→「前往開箱工具系統（不安全）」

### 步驟 4：部署為網頁應用程式

1. 點擊「部署」→「新增部署」
2. 類型選擇：「網頁應用程式」
3. 執行身分選擇：「我」
4. 存取權限選擇：「任何人」
5. 點擊「部署」
6. 複製網頁應用程式 URL

### 步驟 5：更新前端設定

1. ✅ 已更新 `google-sheets-client.js` 中的 `GOOGLE_APPS_SCRIPT_URL` 為：
   ```
   https://script.google.com/macros/s/AKfycbwZK7QwE-7vKW81wyrLTNOcMGtN7EVACYa9uIZ3hwrAjdE9CD8rpIfj_Esg-eAfzzogBA/exec
   ```
2. 重新部署前端應用程式

## 🔧 設定說明

### 試算表 ID
在 `google-sheets-upload.js` 中設定您的 Google Sheets ID：
```javascript
const SPREADSHEET_ID = '您的試算表ID';
```

### 工作表名稱
確保您的 Google Sheets 包含以下工作表：
- `員工名冊`
- `po_header`
- `supplier_contacts`
- `receiving_confirm`

## 📊 資料格式

### 員工名冊工作表
| A 欄 | B 欄 | C 欄 | D 欄 | E 欄 | F 欄 |
|------|------|------|------|------|------|
| 員工編號 | 姓名 | 電話 | Email | 註冊日期 | 狀態 |

### po_header 工作表
| A 欄 | B 欄 | C 欄 | T 欄 |
|------|------|------|------|
| 採購單號 | 採購日期 | 採購對象 | 進貨總數 |

### receiving_confirm 工作表
| A 欄 | B 欄 | C 欄 | D 欄 | E 欄 | F 欄 | G 欄 | H 欄 | I 欄 |
|------|------|------|------|------|------|------|------|------|
| 採購單號 | 開箱人員 | 開箱日期 | 批號匹配 | 商品分類 | 商品名稱 | 商品序號 | 數量 | 備註 |

## 🚀 測試

### 測試函數
在 Google Apps Script 中執行 `testFunctions()` 函數來測試所有功能。

### 測試資料上傳
使用前端應用程式測試資料上傳功能。

## ⚠️ 注意事項

1. **權限設定**：確保 Google Apps Script 有權限存取您的 Google Sheets
2. **試算表 ID**：正確設定試算表 ID，否則會出現錯誤
3. **工作表名稱**：工作表名稱必須完全匹配
4. **資料格式**：確保資料格式符合預期

## 🔍 故障排除

### 常見錯誤
- **權限錯誤**：重新授權 Google Apps Script
- **試算表找不到**：檢查試算表 ID 是否正確
- **工作表找不到**：檢查工作表名稱是否正確

### 除錯方法
1. 查看 Google Apps Script 執行記錄
2. 檢查瀏覽器控制台錯誤訊息
3. 使用 `testFunctions()` 函數測試各個功能
