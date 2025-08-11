/**
 * 開箱工具系統 - Google Sheets 資料處理
 * 用於處理採購單、到貨確認等資料的上傳和管理
 */

// 設定常數
const SPREADSHEET_ID = '1tRyu-k-XwIeXL1m7iz3dQoFVCm1pOTUGC374LBzaXZQ';
const SHEET_NAMES = {
  EMPLOYEES: '員工名冊',
  PO_HEADER: 'po_header',
  SUPPLIER_CONTACTS: 'supplier_contacts',
  RECEIVING_CONFIRM: 'receiving_confirm'
};

/**
 * 主要函數：處理到貨確認資料上傳
 * @param {Object} data - 上傳的資料
 * @returns {Object} 處理結果
 */
function processReceivingConfirm(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.RECEIVING_CONFIRM);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + SHEET_NAMES.RECEIVING_CONFIRM);
    }
    
    // 驗證資料格式
    const validationResult = validateReceivingData(data);
    if (!validationResult.isValid) {
      return {
        success: false,
        error: validationResult.error,
        message: '資料格式驗證失敗'
      };
    }
    
    // 檢查重複序號
    const duplicateCheck = checkDuplicateSerialNumber(sheet, data.serialNumber);
    if (duplicateCheck.isDuplicate) {
      return {
        success: false,
        error: 'DUPLICATE_SERIAL',
        message: '序號已存在：' + data.serialNumber,
        existingData: duplicateCheck.existingData
      };
    }
    
    // 準備插入資料
    const rowData = [
      data.poNumber,           // 採購單號
      data.employeeName,       // 開箱人員
      new Date(),              // 開箱日期
      generateBatchNumber(),    // 批號匹配
      data.productCategory,    // 商品分類
      data.productName,        // 商品名稱
      data.serialNumber,       // 商品序號
      data.quantity,           // 數量
      data.notes || ''         // 備註
    ];
    
    // 插入資料到工作表
    sheet.appendRow(rowData);
    
    // 記錄操作日誌
    logOperation('RECEIVING_CONFIRM_ADD', data);
    
    return {
      success: true,
      message: '資料上傳成功',
      batchNumber: rowData[3],
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('處理到貨確認資料時發生錯誤：', error);
    return {
      success: false,
      error: 'PROCESSING_ERROR',
      message: '處理資料時發生錯誤：' + error.message
    };
  }
}

/**
 * 驗證到貨確認資料格式
 * @param {Object} data - 要驗證的資料
 * @returns {Object} 驗證結果
 */
function validateReceivingData(data) {
  const requiredFields = ['poNumber', 'employeeName', 'productCategory', 'productName', 'serialNumber'];
  const missingFields = [];
  
  // 檢查必填欄位
  requiredFields.forEach(field => {
    if (!data[field] || data[field].toString().trim() === '') {
      missingFields.push(field);
    }
  });
  
  if (missingFields.length > 0) {
    return {
      isValid: false,
      error: 'MISSING_FIELDS',
      message: '缺少必填欄位：' + missingFields.join(', ')
    };
  }
  
  // 驗證採購單號格式
  if (!/^PO\d{8}$/.test(data.poNumber)) {
    return {
      isValid: false,
      error: 'INVALID_PO_NUMBER',
      message: '採購單號格式不正確，應為 PO + 8位數字'
    };
  }
  
  // 驗證序號格式
  if (data.serialNumber.length < 5) {
    return {
      isValid: false,
      error: 'INVALID_SERIAL_NUMBER',
      message: '序號長度不足，至少需要5個字元'
    };
  }
  
  return { isValid: true };
}

/**
 * 檢查序號是否重複
 * @param {Sheet} sheet - 工作表物件
 * @param {string} serialNumber - 要檢查的序號
 * @returns {Object} 檢查結果
 */
function checkDuplicateSerialNumber(sheet, serialNumber) {
  const data = sheet.getDataRange().getValues();
  const serialColumnIndex = 6; // G欄是序號欄位
  
  for (let i = 1; i < data.length; i++) { // 跳過標題行
    if (data[i][serialColumnIndex] === serialNumber) {
      return {
        isDuplicate: true,
        existingData: {
          row: i + 1,
          poNumber: data[i][0],
          employeeName: data[i][1],
          date: data[i][2],
          productName: data[i][5]
        }
      };
    }
  }
  
  return { isDuplicate: false };
}

/**
 * 生成批號
 * @returns {string} 批號
 */
function generateBatchNumber() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hour = String(now.getHours()).padStart(2, '0');
  const minute = String(now.getMinutes()).padStart(2, '0');
  
  return `B${year}${month}${day}${hour}${minute}`;
}

/**
 * 記錄操作日誌
 * @param {string} operation - 操作類型
 * @param {Object} data - 相關資料
 */
function logOperation(operation, data) {
  try {
    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('操作日誌');
    if (!logSheet) {
      // 如果沒有日誌工作表，創建一個
      createLogSheet();
    }
    
    const logData = [
      new Date(),           // 時間戳
      operation,            // 操作類型
      data.employeeName || 'SYSTEM', // 操作人員
      JSON.stringify(data), // 操作資料
      Session.getActiveUser().getEmail() || 'UNKNOWN' // 使用者
    ];
    
    logSheet.appendRow(logData);
  } catch (error) {
    console.error('記錄操作日誌失敗：', error);
  }
}

/**
 * 創建操作日誌工作表
 */
function createLogSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = spreadsheet.insertSheet('操作日誌');
  
  // 設定標題行
  const headers = ['時間戳', '操作類型', '操作人員', '操作資料', '使用者'];
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 設定格式
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
  logSheet.getRange(1, 1, 1, headers.length).setFontColor('white');
  
  // 自動調整欄寬
  logSheet.autoResizeColumns(1, headers.length);
}

/**
 * 查詢採購單資料
 * @param {string} poNumber - 採購單號
 * @returns {Object} 查詢結果
 */
function queryPurchaseOrder(poNumber) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.PO_HEADER);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + SHEET_NAMES.PO_HEADER);
    }
    
    const data = sheet.getDataRange().getValues();
    const poColumnIndex = 0; // A欄是採購單號
    
    for (let i = 1; i < data.length; i++) { // 跳過標題行
      if (data[i][poColumnIndex] === poNumber) {
        return {
          success: true,
          data: {
            poNumber: data[i][0],
            purchaseDate: data[i][1],
            supplier: data[i][2],
            totalQuantity: data[i][19] // T欄是進貨總數
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'PO_NOT_FOUND',
      message: '找不到採購單：' + poNumber
    };
    
  } catch (error) {
    console.error('查詢採購單時發生錯誤：', error);
    return {
      success: false,
      error: 'QUERY_ERROR',
      message: '查詢時發生錯誤：' + error.message
    };
  }
}

/**
 * 查詢員工資料
 * @param {string} employeeId - 員工編號
 * @returns {Object} 查詢結果
 */
function queryEmployee(employeeId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + SHEET_NAMES.EMPLOYEES);
    }
    
    const data = sheet.getDataRange().getValues();
    const idColumnIndex = 0; // A欄是員工編號
    
    for (let i = 1; i < data.length; i++) { // 跳過標題行
      if (data[i][idColumnIndex] === employeeId) {
        return {
          success: true,
          data: {
            employeeId: data[i][0],
            name: data[i][1],
            phone: data[i][2],
            email: data[i][3],
            registerDate: data[i][4],
            status: data[i][5]
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'EMPLOYEE_NOT_FOUND',
      message: '找不到員工：' + employeeId
    };
    
  } catch (error) {
    console.error('查詢員工資料時發生錯誤：', error);
    return {
      success: false,
      error: 'QUERY_ERROR',
      message: '查詢時發生錯誤：' + error.message
    };
  }
}

/**
 * 取得所有員工列表
 * @returns {Object} 員工列表
 */
function getAllEmployees() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + SHEET_NAMES.EMPLOYEES);
    }
    
    const data = sheet.getDataRange().getValues();
    const employees = [];
    
    // 跳過標題行，從第二行開始
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) { // 確保員工編號和姓名存在
        employees.push({
          employeeId: data[i][0],
          name: data[i][1],
          phone: data[i][2] || '',
          email: data[i][3] || '',
          registerDate: data[i][4] || '',
          status: data[i][5] || 'active'
        });
      }
    }
    
    return {
      success: true,
      data: employees,
      count: employees.length
    };
    
  } catch (error) {
    console.error('取得員工列表時發生錯誤：', error);
    return {
      success: false,
      error: 'QUERY_ERROR',
      message: '取得員工列表時發生錯誤：' + error.message
    };
  }
}

/**
 * 測試函數：用於測試 Google Apps Script 功能
 */
function testFunctions() {
  console.log('開始測試 Google Apps Script 功能...');
  
  // 測試取得員工列表
  const employees = getAllEmployees();
  console.log('員工列表測試結果：', employees);
  
  // 測試查詢採購單
  const poResult = queryPurchaseOrder('PO20240801');
  console.log('採購單查詢測試結果：', poResult);
  
  // 測試資料驗證
  const testData = {
    poNumber: 'PO20240801',
    employeeName: '測試員工',
    productCategory: '電子產品',
    productName: '測試商品',
    serialNumber: 'TEST123456'
  };
  const validation = validateReceivingData(testData);
  console.log('資料驗證測試結果：', validation);
  
  console.log('測試完成！');
}

/**
 * 建立網頁應用程式的 doGet 函數
 * @returns {HtmlOutput} HTML 輸出
 */
function doGet() {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta charset="UTF-8">
      <title>開箱工具系統 - Google Apps Script</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .container { max-width: 800px; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 30px; }
        .function-list { list-style: none; padding: 0; }
        .function-list li { margin: 10px 0; padding: 10px; border: 1px solid #ddd; border-radius: 5px; }
        .function-list li:hover { background-color: #f5f5f5; }
        .test-button { background-color: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }
        .test-button:hover { background-color: #3367d6; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📦 開箱工具系統</h1>
          <p>Google Apps Script 後端服務</p>
        </div>
        
        <h2>可用功能</h2>
        <ul class="function-list">
          <li><strong>processReceivingConfirm(data)</strong> - 處理到貨確認資料上傳</li>
          <li><strong>queryPurchaseOrder(poNumber)</strong> - 查詢採購單資料</li>
          <li><strong>queryEmployee(employeeId)</strong> - 查詢員工資料</li>
          <li><strong>getAllEmployees()</strong> - 取得所有員工列表</li>
          <li><strong>validateReceivingData(data)</strong> - 驗證到貨確認資料</li>
        </ul>
        
        <button class="test-button" onclick="testBackend()">測試後端功能</button>
        
        <div id="test-results" style="margin-top: 20px; padding: 10px; background-color: #f0f0f0; border-radius: 5px; display: none;">
          <h3>測試結果</h3>
          <pre id="test-output"></pre>
        </div>
      </div>
      
      <script>
        function testBackend() {
          const resultsDiv = document.getElementById('test-results');
          const outputDiv = document.getElementById('test-output');
          
          resultsDiv.style.display = 'block';
          outputDiv.textContent = '測試中...';
          
          // 這裡可以調用 Google Apps Script 函數
          // 注意：需要設定適當的權限和部署設定
          outputDiv.textContent = '測試完成！請查看 Google Apps Script 控制台輸出。';
        }
      </script>
    </body>
    </html>
  `);
}
