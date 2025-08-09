# 開箱工具系統

這是一個現代化的網頁應用程式，用於管理採購單和到貨確認流程，包含序號掃描、批號生成等功能。

## 功能特色

- 🎨 現代化響應式設計
- 📱 支援手機和桌面裝置
- 🔄 自動從 Google Sheets 讀取資料
- 📋 採購單管理和查詢
- 📦 序號掃描和批號生成
- ⚡ 快速載入和流暢的使用者體驗
- 🛡️ 錯誤處理和重複檢查

## 檔案結構

```
employee-selector-app/
├── index.html              # 主頁面（開箱人員選擇）
├── function-menu.html      # 功能選單頁面
├── po-header.html          # 採購單管理頁面
├── receiving-confirm.html  # 到貨確認頁面
├── styles.css              # CSS 樣式檔案
├── script.js               # 主頁面 JavaScript
├── function-menu.js        # 功能選單 JavaScript
├── po-header.js            # 採購單管理 JavaScript
├── receiving-confirm.js    # 到貨確認 JavaScript
└── README.md               # 說明文件
```

## 使用方法

1. **開啟應用程式**
   - 在瀏覽器中開啟 `index.html` 檔案
   - 或使用本地伺服器（推薦）

2. **使用本地伺服器（推薦）**
   ```bash
   # 使用 Python 3
   python -m http.server 8000
   
   # 或使用 Node.js
   npx http-server
   ```

3. **選擇開箱人員**
   - 從下拉選單中選擇開箱人員
   - 點擊「開箱」按鈕進入功能選單

4. **功能選單**
   - 採購單管理：查看和管理採購單資料
   - 到貨確認：記錄序號和比對到貨數量

5. **採購單管理功能**
   - 採購單號查詢
   - 顯示採購日期、採購對象、進貨總數
   - 採購單列表顯示

6. **到貨確認功能**
   - 採購單號驗證
   - 商品分類選擇
   - 序號掃描記錄
   - 批號自動生成
   - 重複序號檢查
   - 資料上傳功能

## Google Sheets 設定

應用程式會自動讀取以下 Google Sheets：
- **試算表 ID**: `1KWIWFtGa362uoIbGbgVmPFAFkHj0GOM7--ixdydCixI`
- **工作表名稱**: 
  - `員工名冊` - 開箱人員資料
  - `po_header` - 採購單資料
  - `supplier_contacts` - 供應商資料
  - `receiving_confirm` - 到貨確認資料

### 試算表格式要求

#### 員工名冊
| A 欄 | B 欄 | C 欄 | D 欄 | E 欄 | F 欄 |
|------|------|------|------|------|------|
| 員工編號 | 姓名 | 電話 | Email | 註冊日期 | 狀態 |

#### po_header（採購單）
| A 欄 | B 欄 | C 欄 | T 欄 |
|------|------|------|------|
| 採購單號 | 採購日期 | 採購對象 | 進貨總數 |

#### supplier_contacts（供應商）
| A 欄 | B 欄 |
|------|------|
| 供應商代碼 | 供應商名稱 |

#### receiving_confirm（到貨確認）
| A 欄 | B 欄 | C 欄 | D 欄 | E 欄 | F 欄 | G 欄 |
|------|------|------|------|------|------|------|
| 採購單號 | 開箱人員 | 開箱日期 | 批號匹配 | 商品分類 | 商品名稱 | 商品序號 |

## 技術規格

- **前端**: HTML5, CSS3, JavaScript (ES6+)
- **字體**: Noto Sans TC（繁體中文支援）
- **響應式設計**: 支援桌面、平板、手機
- **瀏覽器支援**: Chrome, Firefox, Safari, Edge

## 功能說明

### 主要功能
- 自動載入 Google Sheets 資料
- 下拉式選單選擇員工
- 顯示員工詳細資訊
- 錯誤處理和備用資料

### 備用功能
當無法連接到 Google Sheets 時，系統會自動使用模擬資料，確保應用程式正常運作。

## 自訂設定

### 修改 Google Sheets 來源
在 `script.js` 檔案中修改以下常數：
```javascript
const SHEET_ID = '你的試算表ID';
const SHEET_NAME = '你的工作表名稱';
```

### 自訂樣式
修改 `styles.css` 檔案來自訂外觀和佈局。

## 故障排除

### 常見問題

1. **無法載入員工資料**
   - 檢查網路連線
   - 確認 Google Sheets 已設為公開
   - 檢查瀏覽器控制台錯誤訊息

2. **下拉選單空白**
   - 確認試算表中有資料
   - 檢查工作表名稱是否正確
   - 確認 B 欄位有員工姓名

3. **樣式顯示異常**
   - 確認 CSS 檔案路徑正確
   - 檢查瀏覽器是否支援 CSS 功能

## 授權

此專案僅供學習和內部使用。

## 聯絡資訊

如有問題或建議，請聯絡開發團隊。
