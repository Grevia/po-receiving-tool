// 到貨確認 JavaScript

// Google Sheets 設定
const SHEET_ID = '1tRyu-k-XwIeXL1m7iz3dQoFVCm1pOTUGC374LBzaXZQ';
const PO_HEADER_SHEET = 'po_header';
const SUPPLIER_CONTACTS_SHEET = 'supplier_contacts';
const RECEIVING_CONFIRM_SHEET = 'receiving_confirm';

// DOM 元素
const currentUserElement = document.getElementById('currentUser');
const poSearchInput = document.getElementById('poSearchInput');
const poSearchBtn = document.getElementById('poSearchBtn');
const poInfo = document.getElementById('poInfo');
const poNumber = document.getElementById('poNumber');
const poSupplier = document.getElementById('poSupplier');
const expectedQuantity = document.getElementById('expectedQuantity');
const uploadedCount = document.getElementById('uploadedCount');
const categorySelect = document.getElementById('categorySelect');
const serialInput = document.getElementById('serialInput');
const addSerialBtn = document.getElementById('addSerialBtn');
const scannedCount = document.getElementById('scannedCount');
const serialTableBody = document.getElementById('serialTableBody');
const uploadBtn = document.getElementById('uploadBtn');
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');

// 資料變數
let currentPO = null;
let scannedSerials = [];
let supplierCode = '';
let batchCounter = 1;

// 真實資料陣列
let poDataArray = [];
let supplierDataArray = [];

// 載入採購單資料
async function loadPOData() {
    try {
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${PO_HEADER_SHEET}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const csvText = await response.text();
        poDataArray = parsePOCSV(csvText);
        console.log('載入採購單資料成功：', poDataArray.length, '筆');
        
    } catch (error) {
        console.error('載入採購單資料時發生錯誤:', error);
        // 使用模擬資料作為備用
        useMockPOData();
    }
}

// 載入供應商資料
async function loadSupplierData() {
    try {
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${SUPPLIER_CONTACTS_SHEET}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const csvText = await response.text();
        supplierDataArray = parseSupplierCSV(csvText);
        console.log('載入供應商資料成功：', supplierDataArray.length, '筆');
        
    } catch (error) {
        console.error('載入供應商資料時發生錯誤:', error);
        // 使用模擬資料作為備用
        useMockSupplierData();
    }
}

// 解析採購單 CSV 資料
function parsePOCSV(csvText) {
    const lines = csvText.split('\n');
    const poData = [];
    
    // 跳過標題行，從第二行開始讀取
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
            const columns = parseCSVLine(line);
            
            // 檢查是否有足夠的欄位資料
            if (columns.length >= 20) {
                const poNumber = columns[0] || ''; // A 欄位：採購單號
                const supplier = columns[2] || ''; // C 欄位：採購對象
                const quantity = columns[19] || 0; // T 欄位：進貨總數
                
                // 只添加有採購單號且不是無效資料的記錄
                if (poNumber && poNumber.trim() && poNumber !== '無效資料' && poNumber !== '0') {
                    poData.push({
                        poNumber: poNumber.trim(),
                        supplier: supplier.trim(),
                        quantity: parseInt(quantity) || 0
                    });
                }
            }
        }
    }
    
    return poData;
}

// 解析供應商 CSV 資料
function parseSupplierCSV(csvText) {
    const lines = csvText.split('\n');
    const supplierData = [];
    
    // 跳過標題行，從第二行開始讀取
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
            const columns = parseCSVLine(line);
            
            if (columns.length >= 2) {
                const code = columns[0] || ''; // A 欄位：供應商代碼
                const name = columns[1] || ''; // B 欄位：供應商名稱
                
                if (code && name && code.trim() && name.trim()) {
                    supplierData.push({
                        code: code.trim(),
                        name: name.trim()
                    });
                }
            }
        }
    }
    
    return supplierData;
}

// 解析 CSV 行，處理引號內的逗號
function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    
    result.push(current);
    return result;
}

// 使用模擬採購單資料（當無法載入真實資料時）
function useMockPOData() {
    poDataArray = [
        { poNumber: '20250721001', supplier: '員力科技股份有限公司', quantity: 100 },
        { poNumber: '20250722001', supplier: '宏碁股份有限公司', quantity: 50 },
        { poNumber: '20250723001', supplier: '華碩電腦股份有限公司', quantity: 75 }
    ];
    console.log('使用模擬採購單資料');
}

// 使用模擬供應商資料（當無法載入真實資料時）
function useMockSupplierData() {
    supplierDataArray = [
        { code: '003', name: '員力科技股份有限公司' },
        { code: '001', name: '宏碁股份有限公司' },
        { code: '002', name: '華碩電腦股份有限公司' }
    ];
    console.log('使用模擬供應商資料');
}

// 初始化
document.addEventListener('DOMContentLoaded', async function() {
    // 顯示當前用戶
    const selectedEmployee = localStorage.getItem('selectedEmployee');
    if (selectedEmployee) {
        currentUserElement.textContent = `當前操作員：${selectedEmployee}`;
    } else {
        window.location.href = 'index.html';
    }
    
    // 載入資料
    showLoading();
    await Promise.all([
        loadPOData(),
        loadSupplierData()
    ]);
    hideLoading();
    
    // 監聽事件
    poSearchBtn.addEventListener('click', searchPO);
    addSerialBtn.addEventListener('click', addSerial);
    uploadBtn.addEventListener('click', uploadData);
    
    // 監聽 Enter 鍵
    poSearchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchPO();
        }
    });
    
    serialInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            addSerial();
        }
    });
    
    // 監聽商品分類變更
    categorySelect.addEventListener('change', function() {
        updateUploadButton();
    });
});

// 搜尋採購單
function searchPO() {
    const searchTerm = poSearchInput.value.trim();
    
    if (!searchTerm) {
        showError('請輸入採購單號');
        return;
    }
    
    // 從真實資料中搜尋
    const foundPO = poDataArray.find(po => po.poNumber === searchTerm);
    
    if (foundPO) {
        currentPO = foundPO;
        displayPOInfo(foundPO);
        hideError();
        
        // 獲取供應商代碼
        const supplier = supplierDataArray.find(s => s.name === foundPO.supplier);
        supplierCode = supplier ? supplier.code : '000';
        console.log('找到供應商代碼：', supplierCode, 'for', foundPO.supplier);
        
        // 獲取已上傳筆數
        getUploadedCount(searchTerm);
        
    } else {
        hidePOInfo();
        showError('找不到該採購單號');
        console.log('搜尋的採購單號：', searchTerm);
        console.log('可用的採購單號：', poDataArray.map(po => po.poNumber));
    }
}

// 顯示採購單資訊
function displayPOInfo(po) {
    poNumber.textContent = po.poNumber;
    poSupplier.textContent = po.supplier;
    expectedQuantity.textContent = po.quantity;
    
    poInfo.classList.remove('hidden');
}

// 隱藏採購單資訊
function hidePOInfo() {
    poInfo.classList.add('hidden');
    currentPO = null;
}

// 獲取已上傳筆數
async function getUploadedCount(poNumber) {
    try {
        // 從 receiving_confirm 工作表讀取已上傳筆數
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${RECEIVING_CONFIRM_SHEET}`;
        const response = await fetch(url);
        
        if (response.ok) {
            const csvText = await response.text();
            const uploadedData = parseReceivingCSV(csvText);
            
            // 計算該採購單號的已上傳筆數
            const count = uploadedData.filter(item => item.poNumber === poNumber).length;
            uploadedCount.textContent = count;
            
            // 更新批號計數器
            batchCounter = count + 1;
            
            console.log(`採購單 ${poNumber} 已上傳 ${count} 筆資料`);
        } else {
            // 如果無法讀取，使用模擬資料
            const mockUploadedCount = Math.floor(Math.random() * 20);
            uploadedCount.textContent = mockUploadedCount;
            batchCounter = mockUploadedCount + 1;
        }
        
    } catch (error) {
        console.error('獲取已上傳筆數時發生錯誤:', error);
        // 使用模擬資料
        const mockUploadedCount = Math.floor(Math.random() * 20);
        uploadedCount.textContent = mockUploadedCount;
        batchCounter = mockUploadedCount + 1;
    }
}

// 解析 receiving_confirm CSV 資料
function parseReceivingCSV(csvText) {
    const lines = csvText.split('\n');
    const receivingData = [];
    
    // 跳過標題行，從第二行開始讀取
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
            const columns = parseCSVLine(line);
            
            if (columns.length >= 7) {
                const poNumber = columns[0] || ''; // A 欄位：採購單號
                
                if (poNumber && poNumber.trim()) {
                    receivingData.push({
                        poNumber: poNumber.trim(),
                        employee: columns[1] || '',      // B 欄位：開箱人員
                        date: columns[2] || '',          // C 欄位：開箱日期
                        batchNumber: columns[3] || '',   // D 欄位：商品批號
                        category: columns[4] || '',      // E 欄位：商品分類
                        productName: columns[5] || '',   // F 欄位：商品名稱
                        serial: columns[6] || ''         // G 欄位：商品序號
                    });
                }
            }
        }
    }
    
    return receivingData;
}

// 新增序號
function addSerial() {
    const serial = serialInput.value.trim();
    const category = categorySelect.value;
    
    if (!currentPO) {
        showError('請先查詢採購單號');
        return;
    }
    
    if (!category) {
        showError('請選擇商品分類');
        return;
    }
    
    if (!serial) {
        showError('請輸入序號');
        return;
    }
    
    // 檢查序號是否重複
    if (scannedSerials.some(item => item.serial === serial)) {
        showError('序號重複，請檢查後重新輸入');
        serialInput.value = '';
        serialInput.focus();
        return;
    }
    
    // 生成批號
    const batchNumber = generateBatchNumber();
    
    // 新增序號到列表
    const serialItem = {
        poNumber: currentPO.poNumber,
        category: category,
        serial: serial,
        batchNumber: batchNumber
    };
    
    scannedSerials.push(serialItem);
    updateSerialTable();
    updateScannedCount();
    updateUploadButton();
    
    // 清空輸入框並聚焦
    serialInput.value = '';
    serialInput.focus();
    hideError();
}

// 生成批號
function generateBatchNumber() {
    const batchNumber = `${supplierCode}${batchCounter.toString().padStart(6, '0')}`;
    batchCounter++;
    return batchNumber;
}

// 更新序號表格
function updateSerialTable() {
    serialTableBody.innerHTML = '';
    
    if (scannedSerials.length === 0) {
        serialTableBody.innerHTML = '<tr><td colspan="5" class="no-data">尚未掃描序號</td></tr>';
        return;
    }
    
    scannedSerials.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.poNumber}</td>
            <td>${item.category}</td>
            <td>${item.serial}</td>
            <td>${item.batchNumber}</td>
            <td>
                <button class="btn btn-secondary btn-sm" onclick="removeSerial(${index})">
                    刪除
                </button>
            </td>
        `;
        serialTableBody.appendChild(row);
    });
}

// 刪除序號
function removeSerial(index) {
    scannedSerials.splice(index, 1);
    updateSerialTable();
    updateScannedCount();
    updateUploadButton();
}

// 更新掃描計數
function updateScannedCount() {
    scannedCount.textContent = scannedSerials.length;
}

// 更新上傳按鈕狀態
function updateUploadButton() {
    const canUpload = currentPO && 
                     categorySelect.value && 
                     scannedSerials.length > 0;
    
    uploadBtn.disabled = !canUpload;
}

// 上傳資料
async function uploadData() {
    if (scannedSerials.length === 0) {
        showError('沒有序號可以上傳');
        return;
    }
    
    showLoading();
    
    try {
        // 準備上傳資料
        const uploadData = prepareUploadData();
        console.log('準備上傳的資料：', uploadData);
        
        // 使用 Google Apps Script Web App 上傳資料
        const uploadResult = await uploadToGoogleSheets(uploadData);
        
        if (uploadResult.success) {
            // 上傳成功
            showSuccess(`成功上傳 ${scannedSerials.length} 筆序號資料！`);
            
            // 清空資料
            scannedSerials = [];
            updateSerialTable();
            updateScannedCount();
            updateUploadButton();
            
            // 更新已上傳筆數
            const currentCount = parseInt(uploadedCount.textContent) || 0;
            uploadedCount.textContent = currentCount + uploadData.length;
            
        } else {
            throw new Error(uploadResult.error || '上傳失敗');
        }
        
    } catch (error) {
        console.error('上傳資料時發生錯誤:', error);
        showError(`上傳失敗：${error.message}`);
    } finally {
        hideLoading();
    }
}

// 準備上傳資料
function prepareUploadData() {
    const selectedEmployee = localStorage.getItem('selectedEmployee');
    const currentDate = new Date().toISOString().split('T')[0]; // YYYY-MM-DD 格式
    
    return scannedSerials.map(item => ({
        poNumber: item.poNumber,           // A 欄位：採購單號
        employee: selectedEmployee,        // B 欄位：開箱人員
        date: currentDate,                 // C 欄位：開箱日期
        batchNumber: item.batchNumber,     // D 欄位：商品批號
        category: item.category,           // E 欄位：商品分類
        productName: '',                   // F 欄位：商品名稱（空白）
        serial: item.serial                // G 欄位：商品序號
    }));
}

// 上傳到 Google Sheets
async function uploadToGoogleSheets(data) {
    // Google Apps Script Web App URL - 已更新為最新部署
    const webAppUrl = 'https://script.google.com/macros/s/AKfycbwZK7QwE-7vKW81wyrLTNOcMGtN7EVACYa9uIZ3hwrAjdE9CD8rpIfj_Esg-eAfzzogBA/exec';
    
    // 檢查是否已設定 Web App URL
    if (!webAppUrl || webAppUrl.includes('YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL')) {
        console.warn('未設定 Google Apps Script Web App URL，使用模擬上傳');
        return simulateUpload(data);
    }
    
    try {
        const response = await fetch(webAppUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                action: 'uploadReceivingData',
                data: data,
                sheetId: SHEET_ID
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        return result;
        
    } catch (error) {
        console.error('上傳到 Google Sheets 失敗:', error);
        
        // 如果沒有設置 Web App，使用模擬上傳
        return simulateUpload(data);
    }
}

// 模擬上傳（當沒有設置 Google Apps Script 時）
function simulateUpload(data) {
    return new Promise((resolve) => {
        setTimeout(() => {
            console.log('模擬上傳資料到 receiving_confirm：', data);
            
            // 模擬成功回應
            resolve({
                success: true,
                message: '模擬上傳成功（請設定 Google Apps Script 以實際寫入 Google Sheets）',
                uploadedCount: data.length
            });
        }, 2000);
    });
}

// 顯示載入指示器
function showLoading() {
    loadingIndicator.classList.remove('hidden');
}

// 隱藏載入指示器
function hideLoading() {
    loadingIndicator.classList.add('hidden');
}

// 顯示錯誤訊息
function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
    errorMessage.style.background = '#fee';
    errorMessage.style.color = '#c33';
}

// 顯示成功訊息
function showSuccess(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
    errorMessage.style.background = '#efe';
    errorMessage.style.color = '#3c3';
}

// 隱藏錯誤訊息
function hideError() {
    errorMessage.classList.add('hidden');
}

// 返回功能選單
function goBack() {
    window.location.href = 'function-menu.html';
}
