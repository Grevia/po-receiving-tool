// Google Sheets 設定
const SHEET_ID = '1KWIWFtGa362uoIbGbgVmPFAFkHj0GOM7--ixdydCixI';
const SHEET_NAME = '員工名冊';
const COLUMN_B = 'B'; // B 欄位

// DOM 元素
const employeeSelect = document.getElementById('employeeSelect');
const selectedInfo = document.getElementById('selectedInfo');
const employeeDetails = document.getElementById('employeeDetails');
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');
const openBoxBtn = document.getElementById('openBoxBtn');

// 員工資料陣列
let employees = [];

// 初始化應用程式
document.addEventListener('DOMContentLoaded', function() {
    loadEmployees();
    
    // 監聽下拉選單變更
    employeeSelect.addEventListener('change', function() {
        const selectedEmployee = this.value;
        if (selectedEmployee) {
            showEmployeeInfo(selectedEmployee);
            openBoxBtn.disabled = false;
        } else {
            hideEmployeeInfo();
            openBoxBtn.disabled = true;
        }
    });
    
    // 監聽開箱按鈕
    openBoxBtn.addEventListener('click', function() {
        const selectedEmployee = employeeSelect.value;
        if (selectedEmployee) {
            // 儲存選擇的員工到 localStorage
            localStorage.setItem('selectedEmployee', selectedEmployee);
            // 跳轉到功能選單頁面
            window.location.href = 'function-menu.html';
        }
    });
});

// 載入員工資料
async function loadEmployees() {
    showLoading();
    hideError();
    
    try {
        // 使用 Google Sheets API 的公開 URL 格式
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${SHEET_NAME}`;
        
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const csvText = await response.text();
        const employees = parseCSV(csvText);
        
        populateSelect(employees);
        hideLoading();
        
    } catch (error) {
        console.error('載入員工資料時發生錯誤:', error);
        showError('無法載入員工資料，請檢查網路連線或稍後再試。');
        hideLoading();
        
        // 如果無法載入真實資料，使用模擬資料
        useMockData();
    }
}

// 解析 CSV 資料
function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const employees = [];
    
    // 跳過標題行，從第二行開始讀取
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
            // 處理 CSV 格式，考慮引號內的逗號
            const columns = parseCSVLine(line);
            if (columns.length > 1 && columns[1] && columns[1].trim()) {
                employees.push({
                    id: columns[0] || `員工${i}`,
                    name: columns[1].trim(),
                    phone: columns[2] || '',
                    email: columns[3] || '',
                    registerDate: columns[4] || '',
                    status: columns[5] || ''
                });
            }
        }
    }
    
    return employees;
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

// 填充下拉選單
function populateSelect(employees) {
    // 清空現有選項
    employeeSelect.innerHTML = '';
    
    // 添加預設選項
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '請選擇員工';
    employeeSelect.appendChild(defaultOption);
    
    // 添加員工選項
    employees.forEach(employee => {
        const option = document.createElement('option');
        option.value = employee.name;
        option.textContent = employee.name;
        option.dataset.employee = JSON.stringify(employee);
        employeeSelect.appendChild(option);
    });
    
    // 儲存員工資料
    window.employees = employees;
}

// 顯示員工資訊
function showEmployeeInfo(employeeName) {
    const employee = window.employees.find(emp => emp.name === employeeName);
    
    if (employee) {
        employeeDetails.innerHTML = `
            <p><strong>員工編號：</strong>${employee.id || '未設定'}</p>
            <p><strong>姓名：</strong>${employee.name}</p>
            <p><strong>電話：</strong>${employee.phone || '未設定'}</p>
            <p><strong>Email：</strong>${employee.email || '未設定'}</p>
            <p><strong>註冊日期：</strong>${employee.registerDate || '未設定'}</p>
            <p><strong>狀態：</strong>${employee.status || '未設定'}</p>
        `;
        
        selectedInfo.classList.remove('hidden');
    }
}

// 隱藏員工資訊
function hideEmployeeInfo() {
    selectedInfo.classList.add('hidden');
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

// 使用模擬資料（當無法載入真實資料時）
function useMockData() {
    const mockEmployees = [
        { id: '001', name: '廖堅順', phone: '983663408', email: 'liao@example.com', registerDate: '2024-01-15', status: '在職' },
        { id: '002', name: '小明', phone: '0912345678', email: 'ming@example.com', registerDate: '2024-02-01', status: '在職' },
        { id: '003', name: '小花', phone: '0987654321', email: 'hua@example.com', registerDate: '2024-02-15', status: '在職' },
        { id: '004', name: '張三', phone: '0923456789', email: 'zhang@example.com', registerDate: '2024-03-01', status: '在職' },
        { id: '005', name: '李四', phone: '0934567890', email: 'li@example.com', registerDate: '2024-03-15', status: '在職' }
    ];
    
    populateSelect(mockEmployees);
    showError('使用模擬資料（無法連接到 Google Sheets）');
}

// 重新載入按鈕功能（可選）
function reloadData() {
    loadEmployees();
}

// 匯出選中的員工資料（可選）
function exportSelectedEmployee() {
    const selectedEmployee = employeeSelect.value;
    if (selectedEmployee) {
        const employee = window.employees.find(emp => emp.name === selectedEmployee);
        if (employee) {
            const dataStr = JSON.stringify(employee, null, 2);
            const dataBlob = new Blob([dataStr], {type: 'application/json'});
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `${employee.name}_資料.json`;
            link.click();
            URL.revokeObjectURL(url);
        }
    }
}
