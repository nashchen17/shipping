let excelData = [];
let selectedItems = [];
let itemCounter = 1;
let dailyExports = []; // 儲存當天的所有匯出記錄

document.getElementById('excelFile').addEventListener('change', handleFile);
document.getElementById('productSearch').addEventListener('input', searchProducts);
document.getElementById('searchBtn').addEventListener('click', searchProducts);
document.getElementById('exportQuote').addEventListener('click', exportQuote);
document.getElementById('viewDailyExports').addEventListener('click', viewDailyExports);

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
        itemDiv.className = 'search-item';
        
        itemDiv.innerHTML = `
            <div>
                <strong>${item.dwg}</strong>
                <span style="color: #666; font-size: 12px;">
                    (${item.matchType}: ${item.similarity}%)
                </span><br>
                NC: ${item.nc} | 單價: $${item.unitPrice} | 採購單號: ${item.poNumber}
            </div>
            <button class="add-btn">加入</button>
        `;
        
        // 為整個項目添加點擊事件
        itemDiv.addEventListener('click', () => {
            addToSelected(item.dwg, item.unitPrice, item.nc, item.poNumber);
        });
        
        // 為按鈕添加獨立的點擊事件
        const addBtn = itemDiv.querySelector('.add-btn');
        addBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            addToSelected(item.dwg, item.unitPrice, item.nc, item.poNumber);
        });
        
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

function addToSelected(dwg, unitPrice, nc, poNumber) {
    console.log('addToSelected 被調用:', { dwg, unitPrice, nc, poNumber });
    
    // 修改檢查邏輯：檢查是否為完全相同的產品（所有關鍵欄位都相同）
    const existingItem = selectedItems.find(item => 
        item.dwg === dwg && 
        item.nc === nc && 
        item.poNumber === poNumber
    );
    
    if (existingItem) {
        // 如果是完全相同的產品，詢問是否要增加數量
        const confirmIncrease = confirm(`此產品已在清單中！\n\n產品：${dwg}\nNC：${nc}\n採購單號：${poNumber}\n\n是否要將數量增加 1？`);
        if (confirmIncrease) {
            existingItem.quantity += 1;
            existingItem.subtotal = existingItem.unitPrice * existingItem.quantity;
            updateSelectedTable();
            alert(`產品數量已增加！目前數量：${existingItem.quantity}`);
        }
        return;
    }
    
    const newItem = {
        no: String(itemCounter++).padStart(2, '0'),
        nc: nc,
        dwg: dwg,
        quantity: 1,
        unitPrice: unitPrice,
        subtotal: unitPrice,
        poNumber: poNumber
    };
    
    selectedItems.push(newItem);
    console.log('已添加產品:', newItem);
    console.log('目前選擇的產品:', selectedItems);
    
    updateSelectedTable();
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
    
    // 建立出貨單模板數據
    const deliveryNoteData = createDeliveryNoteTemplate(selectedItems, today, exportCount);
    
    // 將當前匯出加入到當天記錄中
    // 產生兩碼西元年月日+兩碼流水號（如25091601）
    const now = new Date();
    const year = (now.getFullYear() % 100).toString().padStart(2, '0');
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const serial = exportCount.toString().padStart(2, '0');
    const sheetName = `${year}${month}${day}${serial}`;
    const exportRecord = {
        sheetName: sheetName,
        data: deliveryNoteData,
        timestamp: new Date().toLocaleTimeString()
    };
    
    dailyExports.push(exportRecord);
    
    // 嘗試讀取現有的Excel檔案，如果不存在則創建新的
    let wb;
    try {
        // 檢查瀏覽器是否支援 File System Access API
        if ('showSaveFilePicker' in window) {
            // 如果是第一次匯出或檔案不存在，創建新工作簿
            if (dailyExports.length === 1) {
                wb = XLSX.utils.book_new();
            } else {
                // 嘗試從 localStorage 載入現有的工作簿資料
                const existingWorkbookData = localStorage.getItem(`workbook_${today}`);
                if (existingWorkbookData) {
                    wb = XLSX.read(existingWorkbookData, { type: 'binary' });
                } else {
                    wb = XLSX.utils.book_new();
                    // 重新添加之前的工作表
                    for (let i = 0; i < dailyExports.length - 1; i++) {
                        const record = dailyExports[i];
                        const ws = XLSX.utils.aoa_to_sheet(record.data);
                        setupDeliveryNotePrintFormat(ws);
                        XLSX.utils.book_append_sheet(wb, ws, record.sheetName);
                    }
                }
            }
        } else {
            // 對於不支援 File System Access API 的瀏覽器，重新創建包含所有工作表的工作簿
            wb = XLSX.utils.book_new();
            
            // 添加所有當天的匯出記錄
            dailyExports.forEach(record => {
                const ws = XLSX.utils.aoa_to_sheet(record.data);
                setupDeliveryNotePrintFormat(ws);
                
                // 特別設定產品明細行和標題
                for (let row = 5; row <= 12; row++) {
                    for (let col = 0; col <= 6; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        if (!ws[cellAddress]) {
                            ws[cellAddress] = { t: 's', v: '' };
                        }
                        ws[cellAddress].s = {
                            font: { name: 'Cambria', sz: 12, bold: false },
                            alignment: { horizontal: 'left', vertical: 'center' }
                        };
                    }
                }
                
                // 確保A1標題格式正確
                if (ws['A1']) {
                    ws['A1'].v = '出    貨    單';
                    ws['A1'].t = 's';
                    ws['A1'].s = {
                        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
                        font: { name: 'Cambria', bold: true, sz: 23 }
                    };
                }
                
                XLSX.utils.book_append_sheet(wb, ws, record.sheetName);
            });
        }
    } catch (error) {
        console.log('無法讀取現有檔案，創建新的工作簿');
        wb = XLSX.utils.book_new();
        
        // 添加所有當天的匯出記錄
        dailyExports.forEach(record => {
            const ws = XLSX.utils.aoa_to_sheet(record.data);
            setupDeliveryNotePrintFormat(ws);
            
            // 特別設定產品明細行和標題
            for (let row = 5; row <= 12; row++) {
                for (let col = 0; col <= 6; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    if (!ws[cellAddress]) {
                        ws[cellAddress] = { t: 's', v: '' };
                    }
                    ws[cellAddress].s = {
                        font: { name: 'Cambria', sz: 12, bold: false },
                        alignment: { horizontal: 'left', vertical: 'center' }
                    };
                }
            }
            
            // 確保A1標題格式正確
            if (ws['A1']) {
                ws['A1'].v = '出    貨    單';
                ws['A1'].t = 's';
                ws['A1'].s = {
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
                    font: { name: 'Cambria', bold: true, sz: 23 }
                };
            }
            
            XLSX.utils.book_append_sheet(wb, ws, record.sheetName);
        });
    }
    
    // 如果是新增工作表（不是重新創建整個工作簿）
    if (dailyExports.length > 1 && wb.SheetNames.length < dailyExports.length) {
        const currentRecord = dailyExports[dailyExports.length - 1];
        const ws = XLSX.utils.aoa_to_sheet(currentRecord.data);
        setupDeliveryNotePrintFormat(ws);
        
        // 特別設定產品明細行和標題
        for (let row = 5; row <= 12; row++) {
            for (let col = 0; col <= 6; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                if (!ws[cellAddress]) {
                    ws[cellAddress] = { t: 's', v: '' };
                }
                ws[cellAddress].s = {
                    font: { name: 'Cambria', sz: 12, bold: false },
                    alignment: { horizontal: 'left', vertical: 'center' }
                };
            }
        }
        
        // 確保A1標題格式正確
        if (ws['A1']) {
            ws['A1'].v = '出    貨    單';
            ws['A1'].t = 's';
            ws['A1'].s = {
                alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
                font: { name: 'Cambria', bold: true, sz: 23 }
            };
        }
        
        XLSX.utils.book_append_sheet(wb, ws, currentRecord.sheetName);
    }
    
    // 將工作簿資料儲存到 localStorage 以便下次使用
    try {
        const workbookBinary = XLSX.write(wb, { type: 'binary', bookType: 'xlsx' });
        localStorage.setItem(`workbook_${today}`, workbookBinary);
    } catch (error) {
        console.log('無法儲存工作簿到 localStorage:', error);
    }
    
    // 匯出檔案
    XLSX.writeFile(wb, fileName);
    
    // 顯示成功訊息
    const isNewFile = exportCount === 1;
    const actionText = isNewFile ? '已創建新檔案' : '已添加新工作表到現有檔案';
    alert(`出貨單匯出成功！\n${actionText}\n檔案名稱：${fileName}\n工作表總數：${wb.SheetNames.length}\n最新工作表：${exportRecord.sheetName}\n匯出時間：${exportRecord.timestamp}`);
    
    // 清空當前選擇的產品，準備下一次匯出
    selectedItems = [];
    itemCounter = 1;
    updateSelectedTable();
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
    // 單號編號也改為兩碼年+月+日+兩碼流水號
    const now = new Date();
    const year = (now.getFullYear() % 100).toString().padStart(2, '0');
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const day = now.getDate().toString().padStart(2, '0');
    const serial = serialNumber.toString().padStart(2, '0');
    const deliveryNo = `${year}${month}${day}${serial}`;
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
        name: 'Cambria',
        sz: 12,
        bold: false,
        color: { rgb: '000000' }
    };
    
    const titleFont = {
        name: 'Cambria',
        sz: 23,
        bold: true,
        color: { rgb: '000000' }
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
                    name: 'Cambria',
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
        // 計算該工作表的產品數量（扣除標題和格式行）
        const itemCount = record.data.slice(5, 13).filter(row => row[1] && row[1].trim() !== '').length;
        totalItems += itemCount;
        
        message += `${index + 1}. 工作表：${record.sheetName}\n`;
        message += `   匯出時間：${record.timestamp}\n`;
        message += `   產品數量：${itemCount} 項\n\n`;
    });
    
    message += `總計：${dailyExports.length} 個工作表\n`;
    message += `總產品數：${totalItems} 項\n`;
    message += `檔案名稱：出貨單_${today}.xlsx`;
    
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
