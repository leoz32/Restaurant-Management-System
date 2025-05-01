// Margaret 餐厅管理系统 通用JavaScript函数

// 格式化货币的函数
function formatCurrency(value) {
    return new Intl.NumberFormat('zh-CN', {
        style: 'currency',
        currency: 'CNY'
    }).format(value);
}

// 格式化日期的函数
function formatDate(dateString) {
    if (!dateString) return '';
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return new Date(dateString).toLocaleDateString('zh-CN', options);
}

// 格式化日期时间的函数
function formatDateTime(dateTimeString) {
    if (!dateTimeString) return '';
    const options = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    return new Date(dateTimeString).toLocaleDateString('zh-CN', options);
}

// 表单验证
function validateForm(formId) {
    const form = document.getElementById(formId);
    if (!form) return true;
    
    return form.checkValidity();
}

// 导出表格为Excel
function exportTableToExcel(tableID, filename = '') {
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    var tableSelect = document.getElementById(tableID);
    var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
    
    // 指定文件名
    filename = filename ? filename + '.xls' : 'excel_data.xls';
    
    // 创建下载链接元素
    downloadLink = document.createElement("a");
    
    document.body.appendChild(downloadLink);
    
    if(navigator.msSaveOrOpenBlob){
        var blob = new Blob(['\ufeff', tableHTML], {
            type: dataType
        });
        navigator.msSaveOrOpenBlob(blob, filename);
    } else {
        // 创建一个指向blob的链接
        downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
    
        // 设置下载文件名
        downloadLink.download = filename;
        
        //触发点击
        downloadLink.click();
    }
}

// 动态计算订单金额
function calculateOrderTotal() {
    const menuItemSelect = document.getElementById('menu_item_id');
    const quantityInput = document.getElementById('quantity');
    const discountInput = document.getElementById('discount');
    const totalDisplay = document.getElementById('order_total');
    
    if (!menuItemSelect || !quantityInput || !discountInput || !totalDisplay) return;
    
    const updateTotal = () => {
        const selectedOption = menuItemSelect.options[menuItemSelect.selectedIndex];
        const price = parseFloat(selectedOption.getAttribute('data-price') || 0);
        const quantity = parseInt(quantityInput.value) || 1;
        const discount = parseFloat(discountInput.value) || 0;
        
        const total = price * quantity * (1 - discount / 100);
        totalDisplay.value = total.toFixed(2);
        
        // 如果有金额输入字段，同时更新它
        const amountInput = document.getElementById('amount');
        if (amountInput) {
            amountInput.value = total.toFixed(2);
        }
    };
    
    menuItemSelect.addEventListener('change', updateTotal);
    quantityInput.addEventListener('input', updateTotal);
    discountInput.addEventListener('input', updateTotal);
    
    // 初始计算
    updateTotal();
}

// 初始化日期选择器和相关组件
function initDatePickers() {
    // 查找所有date类型的输入
    const dateInputs = document.querySelectorAll('input[type="date"]');
    
    // 为每个日期输入设置默认值（如果没有值）
    dateInputs.forEach(input => {
        if (!input.value) {
            const today = new Date();
            const dd = String(today.getDate()).padStart(2, '0');
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            const yyyy = today.getFullYear();
            input.value = `${yyyy}-${mm}-${dd}`;
        }
    });
}

// 页面加载时自动执行
document.addEventListener('DOMContentLoaded', function() {
    // 初始化日期选择器
    initDatePickers();
    
    // 如果是订单创建或编辑页面，初始化计算器
    if (document.getElementById('order_total')) {
        calculateOrderTotal();
    }
    
    // 初始化任何tooltip
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl)
    });
});