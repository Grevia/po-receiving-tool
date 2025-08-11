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
    // 記錄接收到的資料，方便除錯
    Logger.log('接收到的資料：' + JSON.stringify(data));
    
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
  
  // 驗證採購單號格式 - 接受兩種格式：11位數字 或 PO+8位數字
  if (!(/^\d{11}$/.test(data.poNumber) || /^PO\d{8}$/.test(data.poNumber))) {
    return {
      isValid: false,
      error: 'INVALID_PO_NUMBER',
      message: '採購單號格式不正確，應為11位數字或PO+8位數字'
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
  // 只讀取 G 欄（序號欄位）和相關欄位，提高效能
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { isDuplicate: false }; // 只有標題行或空表
  
  const serialColumnIndex = 6; // G欄是序號欄位（0-based index）
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // 從第2行開始，讀取7欄
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][serialColumnIndex] === serialNumber) {
      return {
        isDuplicate: true,
        existingData: {
          row: i + 2, // 實際行號（跳過標題行）
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('操作日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('操作日誌');
      const headers = ['時間戳', '操作類型', '操作人員', '操作資料', '使用者'];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
              .setBackground('#4285f4').setFontColor('white');
      logSheet.autoResizeColumns(1, headers.length);
    }
    const logData = [
      new Date(),
      operation,
      (data && data.employeeName) || 'SYSTEM',
      JSON.stringify(data || {}),
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'UNKNOWN'
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
 * 支援 action 參數，回傳 JSON 資料
 * @param {Object} e - 請求物件
 * @returns {TextOutput} JSON 格式的回應或 HTML 頁面
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action || '').trim();
  if (!action) {
    // 沒 action 時回一個健康檢查頁，避免手動打開看到空白
    return HtmlService.createHtmlOutput('OK');
  }
  switch (action) {
    case 'getOpeners':
      return json(getOpenersForDropdown());
    case 'getPOHeaders':
      return json(getPOHeadersLite());
    case 'getPOInfo':
      return json(queryPurchaseOrder(String(e.parameter.po || '')));
    default:
      return json({ success: false, error: 'UNKNOWN_ACTION', action }, 400);
  }
}

/**
 * 回傳 JSON 格式的輔助函數
 * @param {Object} obj - 要回傳的物件
 * @param {number} code - HTTP 狀態碼（可選）
 * @returns {TextOutput} JSON 格式的回應
 */
function json(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out; // Apps Script Web App 不支援隨意設 header，但這樣即可跨網域讀 JSON
}

/**
 * 取得開箱人員下拉選單資料
 * 前端友善的函數，回傳員工姓名列表
 * @returns {Object} 包含開箱人員列表的物件
 */
function getOpenersForDropdown() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (!sh) return { openers: [] };
  const last = sh.getLastRow();
  const names = last >= 2 ? sh.getRange(2, 2, last - 1, 1).getValues().flat().filter(Boolean) : [];
  return { openers: names };
}

/**
 * 取得採購單標題簡化資料
 * 前端友善的函數，回傳採購單基本資訊
 * @returns {Object} 包含採購單列表的物件
 */
function getPOHeadersLite() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAMES.PO_HEADER);
  if (!sh) return { items: [] };
  const last = sh.getLastRow();
  const rows = last >= 2 ? sh.getRange(2, 1, last - 1, 20).getValues() : [];
  const items = rows.map(r => ({
    poNo: r[0],
    date: r[1],
    vendor: r[2],
    qty: Number(r[19] || 0)
  })).filter(x => x.poNo);
  return { items };
}

/**
 * 處理 POST 請求的函數
 * 用於接收前端傳送的資料並處理到貨確認
 * @param {Object} e - 請求物件
 * @returns {TextOutput} JSON 格式的回應
 */
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const result = processReceivingConfirm(body);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const res = { success: false, error: 'BAD_REQUEST', message: String(err) };
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 測試上傳功能的函數
 * 在 Google Apps Script 編輯器中直接執行此函數進行測試
 */
function testUpload() {
  const sample = {
    poNumber: '20250721001',
    employeeName: '測試員工',
    productCategory: '手機',
    productName: '測試機',
    serialNumber: 'TEST-' + Math.floor(Math.random() * 1e6),
    quantity: 1,
    notes: 'from test'
  };
  const res = processReceivingConfirm(sample);
  Logger.log(res);
}
