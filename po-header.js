// 採購單管理 JavaScript

// Google Sheets 設定
const SHEET_ID = '1KWIWFtGa362uoIbGbgVmPFAFkHj0GOM7--ixdydCixI';
const SHEET_NAME = 'po_header';

// DOM 元素
const currentUserElement = document.getElementById('currentUser');
const poSearchInput = document.getElementById('poSearch');
const searchBtn = document.getElementById('searchBtn');
const poData = document.getElementById('poData');
const poNumber = document.getElementById('poNumber');
const poDate = document.getElementById('poDate');
const poSupplier = document.getElementById('poSupplier');
const poQuantity = document.getElementById('poQuantity');
const poTableBody = document.getElementById('poTableBody');
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');

// 採購單資料陣列
let poDataArray = [];

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    // 顯示當前用戶
    const selectedEmployee = localStorage.getItem('selectedEmployee');
    if (selectedEmployee) {
        currentUserElement.textContent = `當前操作員：${selectedEmployee}`;
    } else {
        window.location.href = 'index.html';
    }
    
    // 載入採購單資料
    loadPOData();
    
    // 監聽搜尋按鈕
    searchBtn.addEventListener('click', searchPO);
    
    // 監聽 Enter 鍵
    poSearchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchPO();
        }
    });
});

// 載入採購單資料
async function loadPOData() {
    showLoading();
    hideError();
    
    try {
        // 從 Google Sheets 讀取 po_header 資料
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${SHEET_NAME}`;
        
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const csvText = await response.text();
        const poData = parsePOCSV(csvText);
        
        poDataArray = poData;
        populateTable(poData);
        hideLoading();
        
    } catch (error) {
        console.error('載入採購單資料時發生錯誤:', error);
        showError('無法載入採購單資料，請檢查網路連線或稍後再試。');
        hideLoading();
        
        // 如果無法載入真實資料，使用模擬資料
        useMockData();
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
            // 處理 CSV 格式，考慮引號內的逗號
            const columns = parseCSVLine(line);
            
            // 檢查是否有足夠的欄位資料
            if (columns.length >= 20) { // 確保有 T 欄位（第20欄）
                const poNumber = columns[0] || ''; // A 欄位：採購單號
                const poDate = columns[1] || '';   // B 欄位：採購日期
                const supplier = columns[2] || ''; // C 欄位：採購對象
                const quantity = columns[19] || 0; // T 欄位：進貨總數（索引19，因為從0開始）
                
                // 只添加有採購單號的資料
                if (poNumber && poNumber.trim()) {
                    poData.push({
                        poNumber: poNumber.trim(),
                        poDate: poDate.trim(),
                        supplier: supplier.trim(),
                        quantity: parseInt(quantity) || 0
                    });
                }
            }
        }
    }
    
    return poData;
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

// 使用模擬資料（當無法載入真實資料時）
function useMockData() {
    const mockData = [
        { poNumber: '20250721001', poDate: '2025-07-21', supplier: '員力科技股份有限公司', quantity: 100 },
        { poNumber: '20250722001', poDate: '2025-07-22', supplier: '宏碁股份有限公司', quantity: 50 },
        { poNumber: '20250723001', poDate: '2025-07-23', supplier: '華碩電腦股份有限公司', quantity: 75 },
        { poNumber: '20250724001', poDate: '2025-07-24', supplier: '聯想集團', quantity: 120 },
        { poNumber: '20250725001', poDate: '2025-07-25', supplier: '戴爾科技', quantity: 80 }
    ];
    
    poDataArray = mockData;
    populateTable(mockData);
    showError('使用模擬資料（無法連接到 Google Sheets）');
}

// 搜尋採購單
function searchPO() {
    const searchTerm = poSearchInput.value.trim();
    
    if (!searchTerm) {
        showError('請輸入採購單號');
        return;
    }
    
    const foundPO = poDataArray.find(po => po.poNumber === searchTerm);
    
    if (foundPO) {
        displayPODetails(foundPO);
        hideError();
    } else {
        hidePODetails();
        showError('找不到該採購單號');
    }
}

// 顯示採購單詳細資訊
function displayPODetails(po) {
    poNumber.textContent = po.poNumber;
    poDate.textContent = po.poDate;
    poSupplier.textContent = po.supplier;
    poQuantity.textContent = po.quantity;
    
    poData.classList.remove('hidden');
}

// 隱藏採購單詳細資訊
function hidePODetails() {
    poData.classList.add('hidden');
}

// 填充表格
function populateTable(data) {
    poTableBody.innerHTML = '';
    
    if (data.length === 0) {
        poTableBody.innerHTML = '<tr><td colspan="4" class="no-data">無資料</td></tr>';
        return;
    }
    
    data.forEach(po => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${po.poNumber}</td>
            <td>${po.poDate}</td>
            <td>${po.supplier}</td>
            <td>${po.quantity}</td>
        `;
        poTableBody.appendChild(row);
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
}

// 隱藏錯誤訊息
function hideError() {
    errorMessage.classList.add('hidden');
}

// 返回功能選單
function goBack() {
    window.location.href = 'function-menu.html';
}
