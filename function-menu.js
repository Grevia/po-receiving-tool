// 功能選單 JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // 顯示當前選擇的員工
    const selectedEmployee = localStorage.getItem('selectedEmployee');
    const currentUserElement = document.getElementById('currentUser');
    
    if (selectedEmployee) {
        currentUserElement.textContent = `當前操作員：${selectedEmployee}`;
    } else {
        // 如果沒有選擇員工，返回主頁
        window.location.href = 'index.html';
    }
});

// 跳轉到指定功能
function goToFunction(functionName) {
    switch(functionName) {
        case 'po-header':
            window.location.href = 'po-header.html';
            break;
        case 'receiving-confirm':
            window.location.href = 'receiving-confirm.html';
            break;
        default:
            console.error('未知功能：', functionName);
    }
}

// 返回主頁
function goBack() {
    window.location.href = 'index.html';
}
