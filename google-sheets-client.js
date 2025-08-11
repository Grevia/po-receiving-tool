/**
 * 開箱工具系統 - Google Sheets 客戶端
 * 用於前端調用 Google Apps Script 函數
 */

// Google Apps Script 網頁應用程式 URL
// 已部署的實際 URL
const GOOGLE_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwZK7QwE-7vKW81wyrLTNOcMGtN7EVACYa9uIZ3hwrAjdE9CD8rpIfj_Esg-eAfzzogBA/exec';

/**
 * Google Sheets 資料處理類別
 */
class GoogleSheetsManager {
  constructor(scriptUrl) {
    this.scriptUrl = scriptUrl;
    this.isInitialized = false;
  }

  /**
   * 初始化 Google Sheets 管理器
   * @returns {Promise<boolean>} 初始化結果
   */
  async initialize() {
    try {
      // 測試連線
      const testResult = await this.testConnection();
      if (testResult.success) {
        this.isInitialized = true;
        console.log('Google Sheets 管理器初始化成功');
        return true;
      } else {
        console.error('Google Sheets 管理器初始化失敗：', testResult.error);
        return false;
      }
    } catch (error) {
      console.error('初始化 Google Sheets 管理器時發生錯誤：', error);
      return false;
    }
  }

  /**
   * 測試與 Google Apps Script 的連線
   * @returns {Promise<Object>} 測試結果
   */
  async testConnection() {
    try {
      const response = await fetch(this.scriptUrl, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
        }
      });

      if (response.ok) {
        return { success: true, message: '連線成功' };
      } else {
        return { 
          success: false, 
          error: 'HTTP_ERROR', 
          message: `HTTP ${response.status}: ${response.statusText}` 
        };
      }
    } catch (error) {
      return { 
        success: false, 
        error: 'NETWORK_ERROR', 
        message: '網路連線錯誤：' + error.message 
      };
    }
  }

  /**
   * 上傳到貨確認資料
   * @param {Object} data - 到貨確認資料
   * @returns {Promise<Object>} 上傳結果
   */
  async uploadReceivingConfirm(data) {
    if (!this.isInitialized) {
      throw new Error('Google Sheets 管理器尚未初始化');
    }

    try {
      const response = await fetch(this.scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: 'processReceivingConfirm',
          data: data
        })
      });

      if (response.ok) {
        const result = await response.json();
        return result;
      } else {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error('上傳到貨確認資料時發生錯誤：', error);
      return {
        success: false,
        error: 'UPLOAD_ERROR',
        message: '上傳失敗：' + error.message
      };
    }
  }

  /**
   * 查詢採購單資料
   * @param {string} poNumber - 採購單號
   * @returns {Promise<Object>} 查詢結果
   */
  async queryPurchaseOrder(poNumber) {
    if (!this.isInitialized) {
      throw new Error('Google Sheets 管理器尚未初始化');
    }

    try {
      const response = await fetch(this.scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: 'queryPurchaseOrder',
          poNumber: poNumber
        })
      });

      if (response.ok) {
        const result = await response.json();
        return result;
      } else {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error('查詢採購單時發生錯誤：', error);
      return {
        success: false,
        error: 'QUERY_ERROR',
        message: '查詢失敗：' + error.message
      };
      }
  }

  /**
   * 取得所有員工列表
   * @returns {Promise<Object>} 員工列表
   */
  async getAllEmployees() {
    if (!this.isInitialized) {
      throw new Error('Google Sheets 管理器尚未初始化');
    }

    try {
      const response = await fetch(this.scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: 'getAllEmployees'
        })
      });

      if (response.ok) {
        const result = await response.json();
        return result;
      } else {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error('取得員工列表時發生錯誤：', error);
      return {
        success: false,
        error: 'QUERY_ERROR',
        message: '取得員工列表失敗：' + error.message
      };
    }
  }

  /**
   * 查詢員工資料
   * @param {string} employeeId - 員工編號
   * @returns {Promise<Object>} 查詢結果
   */
  async queryEmployee(employeeId) {
    if (!this.isInitialized) {
      throw new Error('Google Sheets 管理器尚未初始化');
    }

    try {
      const response = await fetch(this.scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: 'queryEmployee',
          employeeId: employeeId
        })
      });

      if (response.ok) {
        const result = await response.json();
        return result;
      } else {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error('查詢員工資料時發生錯誤：', error);
      return {
        success: false,
        error: 'QUERY_ERROR',
        message: '查詢員工資料失敗：' + error.message
      };
    }
  }
}

/**
 * 使用範例
 */
async function initializeGoogleSheets() {
  try {
    // 創建 Google Sheets 管理器實例
    const sheetsManager = new GoogleSheetsManager(GOOGLE_APPS_SCRIPT_URL);
    
    // 初始化
    const initialized = await sheetsManager.initialize();
    if (initialized) {
      console.log('Google Sheets 管理器已準備就緒');
      
      // 取得員工列表
      const employees = await sheetsManager.getAllEmployees();
      if (employees.success) {
        console.log('員工列表：', employees.data);
        return sheetsManager;
      } else {
        console.error('取得員工列表失敗：', employees.error);
        return null;
      }
    } else {
      console.error('Google Sheets 管理器初始化失敗');
      return null;
    }
  } catch (error) {
    console.error('初始化 Google Sheets 時發生錯誤：', error);
    return null;
  }
}

/**
 * 上傳到貨確認資料範例
 */
async function uploadReceivingData(sheetsManager, data) {
  try {
    const result = await sheetsManager.uploadReceivingConfirm(data);
    if (result.success) {
      console.log('資料上傳成功：', result);
      return result;
    } else {
      console.error('資料上傳失敗：', result.error);
      return result;
    }
  } catch (error) {
    console.error('上傳資料時發生錯誤：', error);
    return {
      success: false,
      error: 'UPLOAD_ERROR',
      message: '上傳失敗：' + error.message
    };
  }
}

// 匯出類別和函數
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { GoogleSheetsManager, initializeGoogleSheets, uploadReceivingData };
}
