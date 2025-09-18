let excelData = [];
let selectedItems = [];
let itemCounter = 1;
let dailyExports = []; // 儲存當天的所有匯出記錄

document.getElementById('excelFile').addEventListener('change', handleFile);
document.getElementById('productSearch').addEventListener('input', searchProducts);
document.getElementById('searchBtn').addEventListener('click', searchProducts);
document.getElementById('exportQuote').addEventListener('click', exportQuote);
document.getElementById('viewDailyExports').addEventListener('click', viewDailyExports);

// 頁面載入時初始化單據編號顯示
document.addEventListener('DOMContentLoaded', function() {
    updateDocumentNumber();
});

// 更新單據編號顯示
function updateDocumentNumber() {
    const documentNumberElement = document.getElementById('documentNumber');
    if (documentNumberElement) {
        const nextDocumentNumber = generateNextDocumentNumber();
        documentNumberElement.textContent = nextDocumentNumber;
    }
}

// 生成下一個單據編號
function generateNextDocumentNumber() {
    const today = new Date().toISOString().slice(0, 10);
    const lastExportDate = localStorage.getItem('lastExportDate');
    
    // 如果是新的一天，重置計數器
    if (lastExportDate !== today) {
        return generateDocumentNumber(1);
    }
    
    // 取得當前匯出次數並生成下一個編號
    const exportCount = parseInt(localStorage.getItem('dailyExportCount') || '0') + 1;
    return generateDocumentNumber(exportCount);
}

// 生成單據編號（格式：YYYYMMDD-XX）
function generateDocumentNumber(serialNumber) {
    const now = new Date();
    const year = now.getFullYear().toString();
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const serial = serialNumber.toString().padStart(2, '0');
    return `${year}${month}${day}-${serial}`;
}

// 檢查產品是否已被加入選擇清單
function isProductAlreadySelected(dwg, nc, poNumber) {
    return selectedItems.some(item => 
        item.dwg === dwg && 
        item.nc === nc && 
        item.poNumber === poNumber
    );
}

// 刷新搜索結果顯示狀態
function refreshSearchResults() {
    const searchTerm = document.getElementById('productSearch').value.trim();
    if (searchTerm && excelData.length > 0) {
        searchProducts(); // 重新執行搜索以更新狀態
    }
}

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // 顯示載入的工作表名稱
        console.log('載入的工作表:', workbook.SheetNames[0]);
        
        // 轉換資料格式，假設第一列是標題
        excelData = jsonData.slice(1).map(row => ({
            orderDate: row[0] || '',     // 訂購日期 (A欄)
            nc: row[1] || '',            // NC (B欄)
            dwg: row[2] || '',           // DWG (C欄)
            quantity: parseInt(row[3]) || 0,      // 數量 (D欄)
            unitPrice: parseFloat(row[4]) || 0,   // 單價 (E欄)
            subtotal: parseFloat(row[5]) || 0,    // 小計 (F欄)
            poNumber: row[6] || '',      // 採購單號 (G欄)
            deliveryDate: row[7] || '',  // 預計交期 (H欄)
            shipmentNumber: row[8] || '', // 出貨單據編號 (I欄)
            remark: row[13] || ''        // 備註 (N欄)
        })).filter(item => item.dwg && !item.shipmentNumber && item.remark !== '結案'); // 過濾空的DWG、只保留出貨單據編號為空且備註非結案的資料
        
        const totalData = jsonData.slice(1).filter(row => row[2]).length; // 總共有DWG的資料
        const filteredData = excelData.length;
        const excludedCount = totalData - filteredData;
        
        console.log('Excel資料載入完成:', excelData.length, '筆資料');
        alert(`Excel檔案載入成功！\n總資料：${totalData} 筆\n可搜尋資料：${filteredData} 筆\n已排除資料：${excludedCount} 筆（已出貨或結案）`);
    };
    reader.readAsArrayBuffer(file);
}

// 修正 smartSearch 函數 - 完整字串只顯示100%匹配，純數字使用模糊搜尋
function smartSearch(data, searchTerm) {
    const results = [];
    const cleanSearchTerm = searchTerm.trim().toLowerCase();
    
    // 檢查是否為純數字
    const isNumberOnly = /^\d+$/.test(searchTerm.trim());
    
    for (let item of data) {
        // 只搜尋出貨單據編號為空且備註非結案的資料
        if ((item.shipmentNumber && item.shipmentNumber.trim() !== '') || item.remark === '結案') {
            continue;
        }
        
        const cleanDwg = item.dwg.toLowerCase();
        let score = 0;
        let matchType = '';
        
        if (isNumberOnly) {
            // 純數字搜尋 - 使用模糊搜尋
            // 1. 完全匹配數字
            if (cleanDwg === cleanSearchTerm) {
                score = 100;
                matchType = '完全匹配';
            }
            // 2. 開頭匹配
            else if (cleanDwg.startsWith(cleanSearchTerm)) {
                score = 95;
                matchType = '開頭匹配';
            }
            // 3. 包含匹配
            else if (cleanDwg.includes(cleanSearchTerm)) {
                score = 85;
                matchType = '包含匹配';
            }
            // 4. 數字部分匹配
            else {
                const dwgNumbers = item.dwg.match(/\d+/g);
                if (dwgNumbers) {
                    for (let num of dwgNumbers) {
                        if (num.includes(searchTerm)) {
                            score = Math.max(score, 75);
                            matchType = '數字匹配';
                        }
                    }
                }
            }
            
            // 數字搜尋時，顯示60分以上的結果
            if (score >= 60) {
                results.push({
                    ...item,
                    similarity: Math.round(score),
                    matchType: matchType
                });
            }
        } else {
            // 非純數字搜尋 - 只顯示完全匹配
            if (cleanDwg === cleanSearchTerm) {
                results.push({
                    ...item,
                    similarity: 100,
                    matchType: '完全匹配'
                });
            }
        }
    }
    
    // 按分數排序，相同分數按DWG字母順序排序
    return results.sort((a, b) => {
        if (b.similarity === a.similarity) {
            return a.dwg.localeCompare(b.dwg);
        }
        return b.similarity - a.similarity;
    });
}

// 新的相似度計算函數 - 使用Levenshtein距離
function calculateLevenshteinSimilarity(str1, str2) {
    const len1 = str1.length;
    const len2 = str2.length;
    
    // 如果長度差異太大，直接返回低分
    if (Math.abs(len1 - len2) > Math.max(len1, len2) * 0.7) {
        return 0;
    }
    
    const matrix = Array(len1 + 1).fill(null).map(() => Array(len2 + 1).fill(null));
    
    for (let i = 0; i <= len1; i++) {
        matrix[i][0] = i;
    }
    
    for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
    }
    
    for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
            const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
            matrix[i][j] = Math.min(
                matrix[i - 1][j] + 1,     // 刪除
                matrix[i][j - 1] + 1,     // 插入
                matrix[i - 1][j - 1] + cost // 替換
            );
        }
    }
    
    const distance = matrix[len1][len2];
    const maxLen = Math.max(len1, len2);
    return Math.round((1 - distance / maxLen) * 100);
}

function searchProducts() {
    console.log('searchProducts函數被調用');
    const searchTerm = document.getElementById('productSearch').value.trim();
    const resultsDiv = document.getElementById('searchResults');
    
    console.log('搜尋關鍵字:', searchTerm);
    console.log('Excel資料長度:', excelData.length);
    
    if (!searchTerm) {
        resultsDiv.innerHTML = '';
        return;
    }
    
    if (excelData.length === 0) {
        resultsDiv.innerHTML = '<div style="padding: 10px; color: red;">請先上傳Excel檔案</div>';
        return;
    }
    
    // 使用新的智能搜尋
    const filteredData = smartSearch(excelData, searchTerm);
    console.log('搜尋結果數量:', filteredData.length);
    
    if (filteredData.length === 0) {
        resultsDiv.innerHTML = '<div style="padding: 10px; color: gray;">找不到相符的產品</div>';
        return;
    }
    
    // 清空並重新建立結果
    resultsDiv.innerHTML = '';
    
    filteredData.forEach((item, index) => {
        const itemDiv = document.createElement('div');
        const isSelected = isProductAlreadySelected(item.dwg, item.nc, item.poNumber);
        
        itemDiv.className = isSelected ? 'search-item selected-item' : 'search-item';
        
        const buttonText = isSelected ? '已加入' : '加入';
        const buttonDisabled = isSelected ? 'disabled' : '';
        const buttonClass = isSelected ? 'add-btn disabled-btn' : 'add-btn';
        
        itemDiv.innerHTML = `
            <div>
                <strong>${item.dwg}</strong>
                <span style="color: #666; font-size: 12px;">
                    (${item.matchType}: ${item.similarity}%)
                </span><br>
                NC: ${item.nc} | 數量: ${item.quantity} | 單價: $${item.unitPrice} | 採購單號: ${item.poNumber}
            </div>
            <button class="${buttonClass}" ${buttonDisabled}>${buttonText}</button>
        `;
        
        // 只有未選擇的項目才添加點擊事件
        if (!isSelected) {
            // 為整個項目添加點擊事件
            itemDiv.addEventListener('click', () => {
                addToSelected(item.dwg, item.unitPrice, item.nc, item.poNumber, item.quantity);
            });
            
            // 為按鈕添加獨立的點擊事件
            const addBtn = itemDiv.querySelector('.add-btn');
            addBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                addToSelected(item.dwg, item.unitPrice, item.nc, item.poNumber, item.quantity);
            });
        }
        
        resultsDiv.appendChild(itemDiv);
    });
    
    console.log('搜尋結果已顯示');
}

function escapeHtml(text) {
    if (!text) return '';
    
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML
        .replace(/'/g, '&#39;')   // 轉義單引號
        .replace(/"/g, '&quot;'); // 轉義雙引號
}

function addToSelected(dwg, unitPrice, nc, poNumber, originalQuantity = 1) {
    console.log('addToSelected 被調用:', { dwg, unitPrice, nc, poNumber, originalQuantity });
    
    // 修改檢查邏輯：檢查是否為完全相同的產品（所有關鍵欄位都相同）
    const existingItem = selectedItems.find(item => 
        item.dwg === dwg && 
        item.nc === nc && 
        item.poNumber === poNumber
    );
    
    if (existingItem) {
        // 如果是完全相同的產品，詢問是否要增加數量
        const confirmIncrease = confirm(`此產品已在清單中！\n\n產品：${dwg}\nNC：${nc}\n採購單號：${poNumber}\n目前數量：${existingItem.quantity}\n\n是否要將數量增加 ${originalQuantity}？`);
        if (confirmIncrease) {
            existingItem.quantity += originalQuantity;
            existingItem.subtotal = existingItem.unitPrice * existingItem.quantity;
            updateSelectedTable();
            refreshSearchResults(); // 刷新搜索結果顯示狀態
            alert(`產品數量已增加！目前數量：${existingItem.quantity}`);
        }
        return;
    }
    
    const newItem = {
        no: String(itemCounter++).padStart(2, '0'),
        nc: nc,
        dwg: dwg,
        quantity: originalQuantity, // 使用Excel中的原始數量
        unitPrice: unitPrice,
        subtotal: unitPrice * originalQuantity, // 小計也要相應調整
        poNumber: poNumber
    };
    
    selectedItems.push(newItem);
    console.log('已添加產品:', newItem);
    console.log('目前選擇的產品:', selectedItems);
    
    updateSelectedTable();
    updateDocumentNumber(); // 更新單據編號顯示
    refreshSearchResults(); // 刷新搜索結果顯示狀態
}

function updateSelectedTable() {
    const tbody = document.getElementById('selectedItems');
    console.log('updateSelectedTable 被調用, selectedItems:', selectedItems);
    
    if (!tbody) {
        console.error('找不到 selectedItems 元素');
        return;
    }
    
    tbody.innerHTML = selectedItems.map((item, index) => `
        <tr>
            <td>${item.no}</td>
            <td>${item.dwg}</td>
            <td>
                <input type="number" class="quantity-input" value="${item.quantity}" 
                       onchange="updateQuantity(${index}, this.value)" min="1">
            </td>
            <td>$${item.unitPrice.toFixed(2)}</td>
            <td>$${item.subtotal.toFixed(2)}</td>
            <td>${item.poNumber}</td>
            <td>
                <button class="remove-btn" onclick="removeItem(${index})">移除</button>
            </td>
        </tr>
    `).join('');
    
    console.log('表格已更新');
}

function updateQuantity(index, newQuantity) {
    const quantity = parseInt(newQuantity) || 1;
    selectedItems[index].quantity = quantity;
    selectedItems[index].subtotal = selectedItems[index].unitPrice * quantity;
    updateSelectedTable();
}

function removeItem(index) {
    selectedItems.splice(index, 1);
    updateSelectedTable();
    updateDocumentNumber(); // 更新單據編號顯示
    refreshSearchResults(); // 刷新搜索結果顯示狀態
}

function exportQuote() {
    if (selectedItems.length === 0) {
        alert('請先選擇產品');
        return;
    }
    
    const today = new Date().toISOString().slice(0, 10);
    const fileName = `出貨單_${today}.xlsx`;
    
    // 檢查是否為新的一天，如果是則清空之前的記錄
    const lastExportDate = localStorage.getItem('lastExportDate');
    if (lastExportDate !== today) {
        dailyExports = [];
        localStorage.setItem('lastExportDate', today);
        localStorage.setItem('dailyExportCount', '0');
    }
    
    // 取得當前匯出次數
    let exportCount = parseInt(localStorage.getItem('dailyExportCount') || '0') + 1;
    localStorage.setItem('dailyExportCount', exportCount.toString());
    
    // 產生西元年月日-兩碼流水號（如20250916-01）
    const now = new Date();
    const year = now.getFullYear().toString();
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const serial = exportCount.toString().padStart(2, '0');
    const deliveryNo = `${year}${month}${day}-${serial}`;
    
    // 將當前匯出加入到當天記錄中
    const exportRecord = {
        deliveryNo: deliveryNo,
        items: [...selectedItems], // 複製選中的項目
        timestamp: new Date().toLocaleTimeString(),
        date: today
    };
    
    dailyExports.push(exportRecord);
    
    // 創建或更新包含所有出貨單的單一工作表
    let wb;
    try {
        // 暫時禁用localStorage快取，強制創建新工作簿
        // const existingWorkbookData = localStorage.getItem(`workbook_${today}`);
        // if (existingWorkbookData && dailyExports.length > 1) {
        //     wb = XLSX.read(existingWorkbookData, { type: 'binary' });
        //     // 刪除舊的工作表，重新創建包含所有資料的工作表
        //     wb.SheetNames = [];
        //     wb.Sheets = {};
        // } else {
            wb = XLSX.utils.book_new();
        // }
    } catch (error) {
        console.log('無法讀取現有檔案，創建新的工作簿');
        wb = XLSX.utils.book_new();
    }
    
    // 創建包含所有出貨單的工作表資料
    const allDeliveryData = createCombinedDeliveryNoteTemplate(dailyExports);
    const ws = XLSX.utils.aoa_to_sheet(allDeliveryData);
    setupCombinedDeliveryNotePrintFormat(ws, dailyExports.length);
    
    // 額外強制設定第一個標題字體（最後一次嘗試）
    if (ws['A1']) {
        ws['A1'].v = '★★★  出  貨  單  ★★★';  // 強制更改文字內容
        ws['A1'].s = {
            font: { 
                name: 'Arial Black', 
                sz: 36, 
                bold: true, 
                color: { rgb: 'FF0000' } 
            },
            alignment: { horizontal: 'center', vertical: 'center' },
            fill: { fgColor: { rgb: 'FFFF00' } }
        };
        console.log('最終設定A1:', ws['A1'].v, ws['A1'].s);
    }
    
    // 添加工作表到工作簿
    const sheetName = `出貨單_${today.replace(/-/g, '')}`;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    // 將工作簿資料儲存到 localStorage 以便下次使用 (暫時禁用)
    // try {
    //     const workbookBinary = XLSX.write(wb, { type: 'binary', bookType: 'xlsx' });
    //     localStorage.setItem(`workbook_${today}`, workbookBinary);
    // } catch (error) {
    //     console.log('無法儲存工作簿到 localStorage:', error);
    // }
    
    // 匯出檔案 - 使用不同的寫入選項
    XLSX.writeFile(wb, fileName, {
        bookType: 'xlsx',
        cellStyles: true,
        cellNF: false,
        cellHTML: false
    });
    
    // 顯示成功訊息
    const isNewFile = exportCount === 1;
    const actionText = isNewFile ? '已創建新檔案' : '已更新現有檔案';
    alert(`出貨單匯出成功！\n${actionText}\n檔案名稱：${fileName}\n出貨單數量：${dailyExports.length}\n最新出貨單號：${deliveryNo}\n匯出時間：${exportRecord.timestamp}`);
    
    // 清空當前選擇的產品，準備下一次匯出
    selectedItems = [];
    itemCounter = 1;
    updateSelectedTable();
    updateDocumentNumber(); // 更新下一個單據編號
    refreshSearchResults(); // 刷新搜索結果顯示狀態
}

function createDeliveryNoteTemplate(items, date, serialNumber) {
    const data = [];
    
    // 第1列：出貨單標題 - 放在A1，其他欄位留空供合併
    data.push(['出    貨    單', '', '', '', '', '', '']);
    
    // 第2列：慶沅機械有限公司 + 客戶代號
    data.push(['慶沅機械有限公司', '', '', '', '客戶代號', 'FS', '']);
    
    // 第3列：電話 + 出貨日期
    data.push(['TEL：(07)375-6043', '', '', '', '出貨日期', date.replace(/-/g, '/'), '']);
    
    // 第4列：傳真 + 單號編號
    // 單號編號改為西元年月日-兩碼流水號
    const now = new Date();
    const year = now.getFullYear().toString();
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const serial = serialNumber.toString().padStart(2, '0');
    const deliveryNo = `${year}${month}${day}-${serial}`;
    data.push(['FAX：(07)373-2742', '', '', '', '單號編號', deliveryNo, '']);
    
    // 第5列：表格標題
    data.push(['NO.', '品名', '數量', '單價', '小計', '備註', '']);
    
    // 第6-13列：產品明細行（8行）
    for (let i = 0; i < 8; i++) {
        if (i < items.length) {
            const item = items[i];
            data.push([
                item.no,
                item.dwg,
                item.quantity,
                item.unitPrice,
                item.subtotal,
                item.poNumber,
                ''
            ]);
        } else {
            // 空行
            data.push(['', '', '', '', '', '', '']);
        }
    }
    
    // 第14-16列：空行（3行）
    data.push(['', '', '', '', '', '', '']);
    data.push(['', '', '', '', '', '', '']);
    data.push(['', '', '', '', '', '', '']);
    
    // 第17列：合計
    const totalAmount = items.reduce((sum, item) => sum + item.subtotal, 0);
    data.push(['', '', '', '合計', `$${totalAmount}`, '', '']);
    
    // 第18列：打單和簽收
    data.push(['打單', 'Liya', '', '', '', '簽收', '']);
    
    return data;
}

function setupDeliveryNotePrintFormat(ws) {
    // 設定欄寬 - 按照指定的精確寬度
    ws['!cols'] = [
        { wch: 8 },      // A欄
        { wch: 32.12 },  // B欄
        { wch: 6.13 },   // C欄
        { wch: 10.13 },  // D欄
        { wch: 10.13 },  // E欄
        { wch: 18.25 },  // F欄
        { wch: 10 }      // G欄
    ];
    
    // 設定列高
    ws['!rows'] = [];
    
    // 第1列高度設為36
    ws['!rows'][0] = { hpt: 36 };
    
    // 第2到18列高度設為21
    for (let i = 1; i <= 17; i++) { // 索引1-17對應列2-18
        ws['!rows'][i] = { hpt: 21 };
    }
    
    // 設定儲存格合併
    ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, // A1:G1 - 出貨單標題
        { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }, // A2:B2 - 慶沅機械有限公司
        { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, // A3:B3 - 電話
        { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }  // A4:B4 - 傳真
    ];
    
    // 設定頁面佈局為 A4 直印
    ws['!printSettings'] = {
        paperSize: 9, // A4
        orientation: 'portrait',
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        blackAndWhite: false,
        draft: false,
        cellComments: 'None',
        useFirstPageNumber: true,
        horizontalDpi: 300,
        verticalDpi: 300
    };
    
    // 設定邊距 (英寸)
    ws['!margins'] = {
        left: 0.5,
        right: 0.5,
        top: 0.5,
        bottom: 0.5,
        header: 0.3,
        footer: 0.3
    };
    
    // 預設字型樣式
    const defaultFont = {
        name: 'Noto Sans HK Black',   //Cambria
        sz: 12,
        bold: false,
        color: { rgb: '000000' }
    };
    
    const titleFont = {
        name: 'Microsoft JhengHei',
        sz: 50,
        bold: true,
        color: { rgb: 'FF0000' }
    };
    
    // 設定所有儲存格的字型為 Cambria
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            
            // 確保每個儲存格都存在
            if (!ws[cellAddress]) {
                ws[cellAddress] = { t: 's', v: '' };
            }
            
            // 初始化樣式物件
            ws[cellAddress].s = ws[cellAddress].s || {};
            
            // 根據儲存格位置設定字型
            if (cellAddress === 'A1') {
                // A1標題使用大字型
                ws[cellAddress].s = {
                    font: titleFont,
                    alignment: {
                        horizontal: 'center',
                        vertical: 'center',
                        wrapText: false
                    }
                };
            } else {
                // 其他儲存格使用預設字型
                ws[cellAddress].s = {
                    font: defaultFont,
                    alignment: {
                        horizontal: row >= 5 && row <= 12 ? 'left' : 'center',
                        vertical: 'center'
                    }
                };
            }
        }
    }
    
    // 特別針對第6-13列（產品明細行）進行額外設定
    for (let row = 5; row <= 12; row++) { // 第6-13列（索引5-12）
        for (let col = 0; col <= 6; col++) { // A到G欄
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            
            if (!ws[cellAddress]) {
                ws[cellAddress] = { t: 's', v: '' };
            }
            
            // 強制重新設定產品明細行的字型
            ws[cellAddress].s = {
                font: {
                    name: 'Noto Sans HK Black',
                    sz: 12,
                    bold: false,
                    color: { rgb: '000000' }
                },
                alignment: {
                    horizontal: 'left',
                    vertical: 'center'
                },
                border: {
                    top: { style: 'thin', color: { rgb: '000000' } },
                    bottom: { style: 'thin', color: { rgb: '000000' } },
                    left: { style: 'thin', color: { rgb: '000000' } },
                    right: { style: 'thin', color: { rgb: '000000' } }
                }
            };
        }
    }
}

function viewDailyExports() {
    const today = new Date().toISOString().slice(0, 10);
    const lastExportDate = localStorage.getItem('lastExportDate');
    
    if (lastExportDate !== today || dailyExports.length === 0) {
        alert('今天還沒有匯出記錄');
        return;
    }
    
    let message = `今天 (${today}) 的匯出記錄：\n\n`;
    let totalItems = 0;
    
    dailyExports.forEach((record, index) => {
        // 計算該出貨單的產品數量
        const itemCount = record.items.length;
        totalItems += itemCount;
        
        message += `${index + 1}. 出貨單號：${record.deliveryNo}\n`;
        message += `   匯出時間：${record.timestamp}\n`;
        message += `   產品數量：${itemCount} 項\n\n`;
    });
    
    message += `總計：${dailyExports.length} 張出貨單\n`;
    message += `總產品數：${totalItems} 項\n`;
    message += `檔案名稱：出貨單_${today}.xlsx\n`;
    message += `所有出貨單已合併在同一個工作表中`;
    
    alert(message);
}

// 更精確的完整輸入檢查
function isCompleteInput(searchTerm) {
    // 根據您的產品型號格式調整
    const patterns = [
        /^[A-Z]\d+[A-Z]\d+\s*版本\d+$/i,    // 例如：D5ZN10573 版本0
        /^[A-Z]\d+[A-Z]\d+$/i,              // 例如：D5ZN10573
        /^[A-Z]+\d{5,}\s*版本\d+$/i         // 字母開頭+5位以上數字+版本
    ];
    
    const trimmed = searchTerm.trim();
    const isPattern = patterns.some(pattern => pattern.test(trimmed));
    const isLongEnough = trimmed.length >= 10; // 完整型號通常較長
    
    return isPattern && isLongEnough;
}

// 創建包含所有出貨單的合併模板
function createCombinedDeliveryNoteTemplate(exports) {
    const data = [];
    const itemsPerDelivery = 8; // 每張出貨單的產品行數
    const linesPerDelivery = 18; // 每張出貨單的總行數
    
    exports.forEach((exportRecord, exportIndex) => {
        const isFirstDelivery = exportIndex === 0;
        
        // 如果不是第一張出貨單，加入分隔空行
        if (!isFirstDelivery) {
            data.push(['', '', '', '', '', '', '']); // 分隔線
        }
        
        // 第1列：出貨單標題 - 使用更顯眼的文字
        data.push(['★★★  出  貨  單  ★★★', '', '', '', '', '', '']);
        
        // 第2列：慶沅機械有限公司 + 客戶代號
        data.push(['慶沅機械有限公司', '', '', '', '客戶代號', 'FS', '']);
        
        // 第3列：電話 + 出貨日期
        data.push(['TEL：(07)375-6043', '', '', '', '出貨日期', exportRecord.date.replace(/-/g, '/'), '']);
        
        // 第4列：傳真 + 單號編號
        data.push(['FAX：(07)373-2742', '', '', '', '單號編號', exportRecord.deliveryNo, '']);
        
        // 第5列：表格標題
        data.push(['NO.', '品名', '數量', '單價', '小計', '備註', '']);
        
        // 第6-13列：產品明細行（8行）
        for (let i = 0; i < itemsPerDelivery; i++) {
            if (i < exportRecord.items.length) {
                const item = exportRecord.items[i];
                data.push([
                    item.no,
                    item.dwg,
                    item.quantity,
                    item.unitPrice,
                    item.subtotal,
                    item.poNumber,
                    ''
                ]);
            } else {
                // 空行
                data.push(['', '', '', '', '', '', '']);
            }
        }
        
        // 第14-16列：空行（3行）
        data.push(['', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '']);
        
        // 第17列：合計
        const totalAmount = exportRecord.items.reduce((sum, item) => sum + item.subtotal, 0);
        data.push(['', '', '', '合計', `$${totalAmount}`, '', '']);
        
        // 第18列：打單和簽收
        data.push(['打單', 'Liya', '', '', '', '簽收', '']);
    });
    
    return data;
}

// 設定合併出貨單的列印格式
function setupCombinedDeliveryNotePrintFormat(ws, deliveryCount) {
    // 設定欄寬 - 按照指定的精確寬度
    ws['!cols'] = [
        { wch: 8 },      // A欄
        { wch: 32.12 },  // B欄
        { wch: 6.13 },   // C欄
        { wch: 10.13 },  // D欄
        { wch: 10.13 },  // E欄
        { wch: 18.25 },  // F欄
        { wch: 10 }      // G欄
    ];
    
    // 設定列高
    ws['!rows'] = [];
    const linesPerDelivery = 18;
    const totalRows = deliveryCount * linesPerDelivery + (deliveryCount - 1); // 加上分隔行
    
    for (let deliveryIndex = 0; deliveryIndex < deliveryCount; deliveryIndex++) {
        const startRow = deliveryIndex * (linesPerDelivery + 1); // +1 for separator
        const separatorOffset = deliveryIndex > 0 ? 1 : 0; // 第一張沒有分隔行
        
        // 分隔行（如果不是第一張）
        if (deliveryIndex > 0) {
            ws['!rows'][startRow - 1] = { hpt: 21 };
        }
        
        // 第1列高度設為80（標題）- 加大讓標題更顯眼
        ws['!rows'][startRow + separatorOffset] = { hpt: 80 };
        
        // 第2到18列高度設為21
        for (let i = 1; i < linesPerDelivery; i++) {
            ws['!rows'][startRow + separatorOffset + i] = { hpt: 21 };
        }
    }
    
    // 設定儲存格合併和格式
    ws['!merges'] = [];
    
    for (let deliveryIndex = 0; deliveryIndex < deliveryCount; deliveryIndex++) {
        const baseRow = deliveryIndex * (linesPerDelivery + 1) + (deliveryIndex > 0 ? 1 : 0);
        
        // 合併儲存格
        ws['!merges'].push(
            { s: { r: baseRow, c: 0 }, e: { r: baseRow, c: 6 } }, // A1:G1 - 出貨單標題
            { s: { r: baseRow + 1, c: 0 }, e: { r: baseRow + 1, c: 1 } }, // A2:B2 - 慶沅機械有限公司
            { s: { r: baseRow + 2, c: 0 }, e: { r: baseRow + 2, c: 1 } }, // A3:B3 - 電話
            { s: { r: baseRow + 3, c: 0 }, e: { r: baseRow + 3, c: 1 } }  // A4:B4 - 傳真
        );
        
        // 設定標題格式 - 由於樣式不生效，改變文字內容本身
        const titleCellAddress = XLSX.utils.encode_cell({ r: baseRow, c: 0 });
        
        // 使用特殊符號和文字使標題更顯眼
        ws[titleCellAddress] = { 
            t: 's', 
            v: '★★★  出  貨  單  ★★★' 
        };
        
        // 仍然嘗試設定樣式（雖然可能無效）
        ws[titleCellAddress].s = {
            font: { 
                name: 'Arial Black',
                sz: 36,
                bold: true,
                color: { rgb: 'FF0000' }
            },
            alignment: { 
                horizontal: 'center', 
                vertical: 'center' 
            }
        };
        
        console.log('設定標題內容:', ws[titleCellAddress].v);
    }
    
    // 設定頁面佈局為 A4 直印
    ws['!printSettings'] = {
        paperSize: 9, // A4
        orientation: 'portrait',
        scale: 100,
        margins: {
            top: 0.75,
            bottom: 0.75,
            left: 0.7,
            right: 0.7,
            header: 0.3,
            footer: 0.3
        }
    };
}
