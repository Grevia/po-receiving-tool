// Google Apps Script 程式碼
// 請將此程式碼複製到您的 Google Apps Script 專案中

function doPost(e) {
  try {
    // 解析請求資料
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const data = requestData.data;
    const sheetId = requestData.sheetId;
    
    // 根據動作執行相應功能
    switch(action) {
      case 'uploadReceivingData':
        return uploadReceivingData(data, sheetId);
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: '未知的動作'
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 上傳 receiving_confirm 資料
function uploadReceivingData(data, sheetId) {
  try {
    // 開啟試算表
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getSheetByName('receiving_confirm');
    
    if (!sheet) {
      throw new Error('找不到 receiving_confirm 工作表');
    }
    
    // 準備要寫入的資料
    const rowsToAdd = data.map(item => [
      item.poNumber,      // A 欄位：採購單號
      item.employee,      // B 欄位：開箱人員
      item.date,          // C 欄位：開箱日期
      item.batchNumber,   // D 欄位：商品批號
      item.category,      // E 欄位：商品分類
      item.productName,   // F 欄位：商品名稱
      item.serial         // G 欄位：商品序號
    ]);
    
    // 取得最後一行的位置
    const lastRow = sheet.getLastRow();
    
    // 寫入資料到最後一行之後
    if (rowsToAdd.length > 0) {
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: `成功上傳 ${rowsToAdd.length} 筆資料`,
      uploadedCount: rowsToAdd.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 測試函數
function testUpload() {
  const testData = [
    {
      poNumber: '20250721001',
      employee: '測試人員',
      date: '2025-08-08',
      batchNumber: '003000001',
      category: '手機',
      productName: '',
      serial: 'TEST123456'
    }
  ];
  
  const result = uploadReceivingData(testData, '1KWIWFtGa362uoIbGbgVmPFAFkHj0GOM7--ixdydCixI');
  Logger.log(result);
}
