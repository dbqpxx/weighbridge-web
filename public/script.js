// ============================================================
// 全域變數與設定
// ============================================================

// [設定] Google Apps Script 部署網址 (Web App URL)
const API_URL = "https://script.google.com/macros/s/AKfycbxV0dPB2yhL-E817VCBMP2nchiu-eKMl24dsFvyI--vN3K4opyOYr6sAp1xlyPF2BMvlg/exec";

const APP = {
    currentPage: 'dashboard',
    config: null,
    charts: {},
    queryResult: [],
    pageSize: 50,
    currentPageNum: 1
};

// 廠區顏色
const PLANT_COLORS = {
    south: { bg: 'rgba(59, 130, 246, 0.8)', border: '#3b82f6' },
    renwu: { bg: 'rgba(34, 197, 94, 0.8)', border: '#22c55e' },
    gangshan: { bg: 'rgba(245, 158, 11, 0.8)', border: '#f59e0b' }
};

// 垃圾種類顏色
const WASTE_COLORS = [
    '#3b82f6', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6',
    '#06b6d4', '#ec4899', '#6366f1', '#84cc16', '#f97316'
];

// ============================================================
// API 通訊核心 (取代 google.script.run)
// ============================================================

/**
 * 呼叫 Google Apps Script API
 * @param {string} action - 後端函式名稱 (對應 doGet 的 action 參數)
 * @param {object} params - 傳遞給後端的參數物件
 * @returns {Promise} - 回傳 Promise 物件
 */
function callApi(action, params = {}) {
    // 顯示載入中 (可選)
    // document.body.style.cursor = 'wait';

    // 將參數轉換為 URL 查詢字串
    // 注意: 因為是 doGet，所有參數都要放在 URL 上
    // 為了避免中文亂碼，使用 encodeURIComponent
    const queryParams = new URLSearchParams();
    queryParams.append('action', action);

    // 將 params 物件展開到查詢字串
    for (const key in params) {
        if (params.hasOwnProperty(key)) {
            const value = params[key];
            if (typeof value === 'object') {
                queryParams.append(key, JSON.stringify(value));
            } else {
                queryParams.append(key, value);
            }
        }
    }

    const url = `${API_URL}?${queryParams.toString()}`;

    return fetch(url, {
        method: 'GET', // GAS Web App 最簡單是用 GET
        mode: 'cors',  // 跨域請求
        headers: {
            'Content-Type': 'text/plain;charset=utf-8',
        }
    })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            // document.body.style.cursor = 'default';
            return result;
        })
        .catch(error => {
            // document.body.style.cursor = 'default';
            console.error('API Call Error:', error);
            throw error;
        });
}

/**
 * 呼叫 Google Apps Script API (POST)
 * 用於上傳大量資料，避免 GET URL 長度限制
 */
function callApiPost(action, params = {}) {
    const url = API_URL; // POST 用同一個 URL (doPost)
    const payload = { ...params, action: action };

    return fetch(url, {
        method: 'POST',
        mode: 'cors',
        headers: {
            'Content-Type': 'text/plain;charset=utf-8',
        },
        body: JSON.stringify(payload)
    })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .catch(error => {
            console.error('API Post Error:', error);
            throw error;
        });
}

// ============================================================
// 初始化
// ============================================================

document.addEventListener('DOMContentLoaded', function () {
    initNavigation();
    initUpload();
    initQuery();
    initComparison();
    initReports();
    updateTime();
    setInterval(updateTime, 1000);

    // 載入設定
    loadConfig();

    // 載入儀表板資料
    loadDashboard();

    // 載入來源選項（區隊/廠商下拉選單）
    loadSourceOptions();
});

// ============================================================
// 導航功能
// ============================================================

function initNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    const menuToggle = document.getElementById('menuToggle');
    const sidebar = document.querySelector('.sidebar');

    navItems.forEach(item => {
        item.addEventListener('click', function () {
            const page = this.dataset.page;
            switchPage(page);

            // 更新導航狀態
            navItems.forEach(i => i.classList.remove('active'));
            this.classList.add('active');

            // 手機版關閉側邊欄
            if (window.innerWidth <= 768) {
                sidebar.classList.remove('open');
            }
        });
    });

    // 手機版選單切換
    menuToggle.addEventListener('click', function () {
        sidebar.classList.toggle('open');
    });
}

function switchPage(page) {
    const pages = document.querySelectorAll('.page');
    const titles = {
        dashboard: '總覽',
        upload: '資料上傳',
        query: '查詢分析',
        comparison: '跨廠比較',
        reports: '報表匯出'
    };

    pages.forEach(p => p.classList.remove('active'));
    const targetPage = document.getElementById('page-' + page);
    if (targetPage) targetPage.classList.add('active');

    const titleEl = document.getElementById('pageTitle');
    if (titleEl) titleEl.textContent = titles[page] || page;

    APP.currentPage = page;

    // 頁面切換時的特殊處理
    if (page === 'dashboard') {
        loadDashboard();
    }
}

function updateTime() {
    const now = new Date();
    const options = {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        weekday: 'short'
    };
    const timeEl = document.getElementById('currentTime');
    if (timeEl) timeEl.textContent = now.toLocaleString('zh-TW', options);
}

// ============================================================
// 載入設定
// ============================================================

function loadConfig() {
    callApi('getConfig')
        .then(result => {
            if (result.success) {
                APP.config = result.data;
                console.log('設定載入成功', APP.config);
            }
        })
        .catch(error => {
            console.error('設定載入失敗', error);
        });
}

// ============================================================
// 載入來源選項
// ============================================================

function loadSourceOptions() {
    callApi('getSourceList', {})
        .then(result => {
            if (result.success) {
                APP.sourceList = result.data;
                updateSourceDropdowns();
                console.log('來源清單載入成功', APP.sourceList);
            }
        })
        .catch(error => {
            console.error('來源清單載入失敗', error);
        });
}

function updateSourceDropdowns() {
    // 使用 datalist ID
    const districtList = document.getElementById('districtList');
    const vendorList = document.getElementById('vendorList');
    const wasteTypeList = document.getElementById('wasteTypeList');

    if (!APP.sourceList) {
        console.log('sourceList 尚未載入');
        return;
    }

    console.log('更新選單，區隊:', APP.sourceList.districts?.length, '廠商:', APP.sourceList.vendors?.length, '垃圾種類:', APP.sourceList.wasteTypes?.length);

    // 更新區隊清單
    if (districtList && APP.sourceList.districts) {
        districtList.innerHTML = '';
        APP.sourceList.districts.forEach(d => {
            const option = document.createElement('option');
            option.value = d.name;
            districtList.appendChild(option);
        });
    }

    // 更新廠商清單
    if (vendorList && APP.sourceList.vendors) {
        vendorList.innerHTML = '';
        APP.sourceList.vendors.forEach(v => {
            const option = document.createElement('option');
            option.value = v.name;
            vendorList.appendChild(option);
        });
    }

    // [NEW] 更新垃圾種類清單
    if (wasteTypeList && APP.sourceList.wasteTypes) {
        wasteTypeList.innerHTML = '';
        // Add default '全部' option if needed, but usually query logic handles empty
        const allOption = document.createElement('option');
        allOption.value = '全部';
        wasteTypeList.appendChild(allOption);

        APP.sourceList.wasteTypes.forEach(w => {
            const option = document.createElement('option');
            option.value = w.name;
            wasteTypeList.appendChild(option);
        });
    }
}

// ============================================================
// 儀表板功能
// ============================================================

function loadDashboard() {
    // 預設取得最近 60 天的資料（確保能看到剛上傳的上個月資料）
    const now = new Date();
    const startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1); // 上個月 1 號
    const endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);   // 本月底

    const params = {
        startDate: formatDate(startDate),
        endDate: formatDate(endDate)
    };

    // 載入摘要統計
    callApi('getSummary', params)
        .then(updateSummaryStats)
        .catch(handleError);

    // 載入垃圾種類統計
    callApi('getWasteTypeStats', params)
        .then(updateWasteTypeChart)
        .catch(handleError);

    // 載入每日趨勢
    callApi('getDailyTrend', params)
        .then(updateDailyTrendChart)
        .catch(handleError);
}

function updateSummaryStats(result) {
    if (!result.success) {
        console.error('載入統計失敗', result.error);
        return;
    }

    const data = result.data;

    document.getElementById('totalRecords').textContent =
        formatNumber(data.totalRecords);
    document.getElementById('totalWeight').textContent =
        formatNumber(data.totalNetWeightTon.toFixed(3));
    document.getElementById('totalAmount').textContent =
        formatNumber(data.totalAmount);
    document.getElementById('avgWeight').textContent =
        data.totalRecords > 0
            ? formatNumber(Math.round(data.totalNetWeight / data.totalRecords))
            : '-';

    // 更新廠區比較圖表
    updatePlantComparisonChart(data.plantStats);
}

function updatePlantComparisonChart(plantStats) {
    const ctx = document.getElementById('plantComparisonChart');

    if (APP.charts.plantComparison) {
        APP.charts.plantComparison.destroy();
    }

    const labels = plantStats.map(p => p.name);
    const data = plantStats.map(p => p.netWeightTon);
    const colors = plantStats.map(p => PLANT_COLORS[p.code]?.bg || '#6b7280');

    APP.charts.plantComparison = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '處理量 (噸)',
                data: data,
                backgroundColor: colors,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => formatNumber(value) + ' 噸'
                    }
                }
            }
        }
    });
}

function updateWasteTypeChart(result) {
    if (!result.success) return;
    const ctx = document.getElementById('wasteTypeChart');

    if (APP.charts.wasteType) {
        APP.charts.wasteType.destroy();
    }

    // 取前 8 種
    const topData = result.data.slice(0, 8);

    APP.charts.wasteType = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: topData.map(d => d.wasteType),
            datasets: [{
                data: topData.map(d => d.netWeightTon),
                backgroundColor: WASTE_COLORS,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        boxWidth: 12,
                        padding: 8,
                        font: {
                            size: 11
                        }
                    }
                }
            }
        }
    });
}

function updateDailyTrendChart(result) {
    if (!result.success) return;
    const ctx = document.getElementById('dailyTrendChart');

    if (APP.charts.dailyTrend) {
        APP.charts.dailyTrend.destroy();
    }

    const data = result.data;

    APP.charts.dailyTrend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: data.map(d => d.date.substring(5)), // MM-DD
            datasets: [{
                label: '處理量 (噸)',
                data: data.map(d => d.netWeightTon),
                borderColor: '#4f46e5',
                backgroundColor: 'rgba(79, 70, 229, 0.1)',
                fill: true,
                tension: 0.4,
                pointRadius: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => formatNumber(value) + ' 噸'
                    }
                }
            }
        }
    });
}

// ============================================================
// 上傳功能
// ============================================================

let uploadedData = null;

function initUpload() {
    const fileInput = document.getElementById('fileInput');
    const fileUploadArea = document.getElementById('fileUploadArea');
    const uploadBtn = document.getElementById('uploadBtn');

    if (!fileUploadArea) return;

    // 點擊上傳區域
    fileUploadArea.addEventListener('click', () => fileInput.click());

    // 拖曳上傳
    fileUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        fileUploadArea.classList.add('dragover');
    });

    fileUploadArea.addEventListener('dragleave', () => {
        fileUploadArea.classList.remove('dragover');
    });

    fileUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        fileUploadArea.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file);
    });

    // 選擇檔案
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) handleFile(file);
        });
    }

    // 上傳按鈕
    if (uploadBtn) {
        uploadBtn.addEventListener('click', doUpload);
    }
}

function handleFile(file) {
    const fileName = document.getElementById('fileName');
    const uploadBtn = document.getElementById('uploadBtn');
    fileName.textContent = file.name;

    // 解析檔案
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // 轉換為 JSON 陣列
            uploadedData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            uploadBtn.disabled = false;
            showToast('檔案載入成功，共 ' + (uploadedData.length - 1) + ' 筆資料', 'success');
        } catch (error) {
            showToast('檔案解析失敗: ' + error.message, 'error');
        }
    };
    reader.readAsBinaryString(file);
}

function doUpload() {
    const plant = document.getElementById('uploadPlant').value;
    const yearMonth = document.getElementById('uploadYearMonth').value;
    const uploadBtn = document.getElementById('uploadBtn');
    const progress = document.getElementById('uploadProgress');
    const result = document.getElementById('uploadResult');

    if (!plant) {
        showToast('請選擇廠區', 'error');
        return;
    }
    if (!yearMonth) {
        showToast('請輸入年月', 'error');
        return;
    }
    if (!uploadedData) {
        showToast('請選擇檔案', 'error');
        return;
    }

    // 顯示進度
    uploadBtn.disabled = true;
    progress.style.display = 'block';
    result.innerHTML = '';
    result.className = 'upload-result';

    // 模擬進度
    let progressValue = 0;
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');

    const interval = setInterval(() => {
        progressValue += 10;
        if (progressValue <= 90) {
            progressFill.style.width = progressValue + '%';
            progressText.textContent = '上傳中... ' + progressValue + '%';
        }
    }, 200);

    // 呼叫後端匯入 (使用 POST)
    callApiPost('importData', {
        plant: plant,
        yearMonth: yearMonth,
        data: uploadedData
    })
        .then(result => {
            clearInterval(interval);
            progressFill.style.width = '100%';
            progressText.textContent = '完成!';

            if (result.success) {
                const weightMsg = `成功匯入 ${result.count} 筆資料 (共 ${formatNumber(result.totalWeightKg)} KG)`;
                document.getElementById('uploadResult').innerHTML = '✓ ' + weightMsg;
                document.getElementById('uploadResult').classList.add('success');
                showToast(weightMsg, 'success');

                // 如果有新增來源，重新載入來源清單
                if (result.newSources > 0) {
                    loadSourceOptions();
                }

                // 重設表單
                setTimeout(() => {
                    document.getElementById('uploadPlant').value = '';
                    document.getElementById('uploadYearMonth').value = '';
                    document.getElementById('fileName').textContent = '';
                    uploadedData = null;
                    document.getElementById('uploadProgress').style.display = 'none';
                    document.getElementById('progressFill').style.width = '0%';
                    document.getElementById('fileInput').value = '';
                }, 2000);
            } else {
                throw new Error(result.error);
            }
            uploadBtn.disabled = false;
        })
        .catch(error => {
            clearInterval(interval);
            document.getElementById('uploadProgress').style.display = 'none';

            let errorMsg = error.message;
            if (errorMsg && errorMsg.includes('414')) {
                errorMsg = '資料量過大';
            }

            document.getElementById('uploadResult').innerHTML = '✗ 上傳失敗: ' + errorMsg;
            document.getElementById('uploadResult').classList.add('error');
            showToast('上傳失敗: ' + errorMsg, 'error');
            uploadBtn.disabled = false;
        });
}

// ============================================================
// 查詢功能
// ============================================================

function initQuery() {
    const queryBtn = document.getElementById('queryBtn');
    const resetBtn = document.getElementById('resetBtn');
    const exportBtn = document.getElementById('exportExcelBtn');
    const viewDetailBtn = document.getElementById('viewDetailBtn');
    const viewSummaryBtn = document.getElementById('viewSummaryBtn');

    if (!queryBtn) return;

    // 設定預設日期（當月）
    const now = new Date();
    const startDate = new Date(now.getFullYear(), now.getMonth(), 1);
    const endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);

    document.getElementById('queryStartDate').value = formatDate(startDate);
    document.getElementById('queryEndDate').value = formatDate(endDate);

    // 互斥邏輯：區隊與廠商只能選一個
    const districtSelect = document.getElementById('queryDistrict');
    const vendorSelect = document.getElementById('queryVendor');

    if (districtSelect && vendorSelect) {
        districtSelect.addEventListener('change', function () {
            if (this.value) vendorSelect.value = "";
        });

        vendorSelect.addEventListener('change', function () {
            if (this.value) districtSelect.value = "";
        });
    }

    queryBtn.addEventListener('click', doQuery);
    resetBtn.addEventListener('click', resetQuery);
    exportBtn.addEventListener('click', exportQueryResult);

    // 分頁
    document.getElementById('prevPage').addEventListener('click', () => changePage(-1));
    document.getElementById('nextPage').addEventListener('click', () => changePage(1));

    // 明細/匯總切換
    viewDetailBtn.addEventListener('click', () => switchView('detail'));
    viewSummaryBtn.addEventListener('click', () => switchView('summary'));
}

function switchView(view) {
    const detailView = document.getElementById('detailView');
    const summaryView = document.getElementById('summaryView');
    const viewDetailBtn = document.getElementById('viewDetailBtn');
    const viewSummaryBtn = document.getElementById('viewSummaryBtn');
    const pagination = document.getElementById('pagination');

    if (view === 'detail') {
        detailView.style.display = 'block';
        summaryView.style.display = 'none';
        viewDetailBtn.classList.add('active');
        viewSummaryBtn.classList.remove('active');
        // 明細模式顯示分頁
        const totalPages = Math.ceil(APP.queryResult.length / APP.pageSize);
        pagination.style.display = totalPages > 1 ? 'flex' : 'none';
    } else {
        detailView.style.display = 'none';
        summaryView.style.display = 'block';
        viewDetailBtn.classList.remove('active');
        viewSummaryBtn.classList.add('active');
        // 匯總模式隱藏分頁
        pagination.style.display = 'none';
        // 渲染匯總表格
        renderSummaryTable();
    }
}

function renderSummaryTable() {
    const summaryBody = document.getElementById('summaryBody');
    const summaryFooter = document.getElementById('summaryFooter');

    if (APP.queryResult.length === 0) {
        summaryBody.innerHTML = '<tr><td colspan="5" class="no-data">無符合條件的資料</td></tr>';
        summaryFooter.style.display = 'none';
        return;
    }

    // 依垃圾種類匯總
    const summary = {};
    let totalWeight = 0;
    let totalAmount = 0;
    let totalCount = 0;

    APP.queryResult.forEach(row => {
        const type = row.wasteType || '未分類';
        if (!summary[type]) {
            summary[type] = { count: 0, netWeight: 0, amount: 0 };
        }
        summary[type].count++;
        summary[type].netWeight += (row.netWeight || 0);
        summary[type].amount += (row.amount || 0);
        totalWeight += (row.netWeight || 0);
        totalAmount += (row.amount || 0);
        totalCount++;
    });

    // 轉換為陣列並排序
    const summaryList = Object.entries(summary)
        .map(([type, data]) => ({
            type,
            count: data.count,
            netWeight: data.netWeight,
            netWeightTon: data.netWeight / 1000,
            amount: data.amount,
            percentage: totalWeight > 0 ? ((data.netWeight / totalWeight) * 100).toFixed(1) : 0
        }))
        .sort((a, b) => b.netWeight - a.netWeight);

    // 渲染表格
    summaryBody.innerHTML = summaryList.map(row => `
            <tr>
                <td>${row.type}</td>
                <td>${formatNumber(row.count)}</td>
                <td>${formatNumber(row.netWeightTon.toFixed(3))}</td>
                <td>${formatNumber(row.amount)}</td>
                <td>${row.percentage}%</td>
            </tr>
        `).join('');

    // 更新合計列
    summaryFooter.style.display = '';
    document.getElementById('summaryTotalCount').textContent = formatNumber(totalCount);
    document.getElementById('summaryTotalWeight').textContent = formatNumber((totalWeight / 1000).toFixed(3));
    document.getElementById('summaryTotalAmount').textContent = formatNumber(totalAmount);
}

function doQuery() {
    const startDate = document.getElementById('queryStartDate').value;
    const endDate = document.getElementById('queryEndDate').value;
    const wasteType = document.getElementById('queryWasteType').value;
    const district = document.getElementById('queryDistrict').value;
    const vendor = document.getElementById('queryVendor').value;

    // 組合來源篩選（區隊或廠商）
    const source = district || vendor || '';

    // 取得選中的廠區
    const plants = [];
    document.querySelectorAll('input[name="queryPlant"]:checked').forEach(cb => {
        plants.push(cb.value);
    });

    const params = {
        startDate: startDate,
        endDate: endDate,
        plants: plants.join(','),
        wasteTypes: wasteType === '全部' ? '' : wasteType,
        source: source,
        limit: 2000 // 增加限制筆數
    };

    showToast('查詢中...', 'info');

    callApi('queryData', params)
        .then(displayQueryResult)
        .catch(handleError);
}

function displayQueryResult(result) {
    console.log('收到查詢結果 (Raw):', result);

    // 檢查 result 是否為 null 或 undefined
    if (!result) {
        showToast('❌ 查詢失敗：伺服器未回傳資料 (Result is null)', 'error');
        console.error('Critical Error: Result is null or undefined.');
        return;
    }

    if (!result.success) {
        showToast('查詢失敗: ' + (result.error || '未知錯誤'), 'error');
        console.error('Server reported failure:', result.error);
        return;
    }

    // 檢查 data 是否存在
    if (!Array.isArray(result.data)) {
        showToast('❌ 查詢失敗：返回數據格式錯誤 (Data is not array)', 'error');
        console.error('Invalid data format:', result);
        return;
    }

    APP.queryResult = result.data;
    APP.currentPageNum = 1;

    console.log('✅ 查詢成功，數據筆數:', result.data.length);

    // 更新統計
    const stats = document.getElementById('queryStats');
    stats.style.display = 'flex';

    const totalWeight = APP.queryResult.reduce((sum, r) => sum + (r.netWeight || 0), 0);
    const totalAmount = APP.queryResult.reduce((sum, r) => sum + (r.amount || 0), 0);

    document.getElementById('queryCount').textContent = formatNumber(result.total);
    document.getElementById('queryWeight').textContent = formatNumber((totalWeight / 1000).toFixed(3)) + ' 噸';
    document.getElementById('queryAmount').textContent = formatNumber(totalAmount) + ' 元';

    // 啟用匯出按鈕
    document.getElementById('exportExcelBtn').disabled = result.data.length === 0;

    // 渲染表格
    renderQueryTable();

    showToast('查詢完成，共 ' + result.total + ' 筆', 'success');
}

function renderQueryTable() {
    const tbody = document.getElementById('queryResultBody');
    const pagination = document.getElementById('pagination');

    if (APP.queryResult.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" class="no-data">無符合條件的資料</td></tr>';
        pagination.style.display = 'none';
        return;
    }

    const start = (APP.currentPageNum - 1) * APP.pageSize;
    const end = Math.min(start + APP.pageSize, APP.queryResult.length);
    const pageData = APP.queryResult.slice(start, end);

    tbody.innerHTML = pageData.map(row => `
    <tr>
      <td>${row.seqNo}</td>
      <td>${row.plantName}</td>
      <td>${formatDateTime(row.datetime)}</td>
      <td>${row.lane}</td>
      <td>${row.vehicleNo}</td>
      <td>${row.source}</td>
      <td>${row.wasteType}</td>
      <td class="text-right">${formatNumber(row.grossWeight)}</td>
      <td class="text-right">${formatNumber(row.tareWeight)}</td>
      <td class="text-right">${formatNumber(row.netWeight)}</td>
      <td class="text-right">${formatNumber(row.amount)}</td>
      <td>${row.remark || ''}</td>
    </tr>
  `).join('');

    // 更新分頁資訊
    const totalPages = Math.ceil(APP.queryResult.length / APP.pageSize);
    document.getElementById('pageInfo').textContent = `第 ${APP.currentPageNum} 頁 / 共 ${totalPages} 頁`;

    // 按鈕狀態
    document.getElementById('prevPage').disabled = APP.currentPageNum === 1;
    document.getElementById('nextPage').disabled = APP.currentPageNum === totalPages;
}

function changePage(delta) {
    const totalPages = Math.ceil(APP.queryResult.length / APP.pageSize);
    const newPage = APP.currentPageNum + delta;

    if (newPage >= 1 && newPage <= totalPages) {
        APP.currentPageNum = newPage;
        renderQueryTable();
        // 滾動到表格頂部
        document.querySelector('.result-header').scrollIntoView({ behavior: 'smooth' });
    }
}

function resetQuery() {
    document.getElementById('queryStartDate').value = formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    document.getElementById('queryEndDate').value = formatDate(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0));
    document.getElementById('queryDistrict').value = '';
    document.getElementById('queryVendor').value = '';
    document.getElementById('queryWasteType').value = '';
    document.querySelectorAll('input[name="queryPlant"]').forEach(cb => cb.checked = true);

    APP.queryResult = [];
    document.getElementById('queryStats').style.display = 'none';
    renderQueryTable();
}

function exportQueryResult() {
    if (APP.queryResult.length === 0) return;

    // 使用 SheetJS 匯出
    const ws = XLSX.utils.json_to_sheet(APP.queryResult.map(row => ({
        "日期時間": formatDateTime(row.datetime),
        "廠區": row.plantName,
        "車號": row.vehicleNo,
        "來源": row.source,
        "垃圾種類": row.wasteType,
        "毛重(kg)": row.grossWeight,
        "空重(kg)": row.tareWeight,
        "淨重(kg)": row.netWeight,
        "金額": row.amount,
        "備註": row.remark
    })));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "查詢結果");
    XLSX.writeFile(wb, `地磅查詢結果_${formatDate(new Date())}.xlsx`);
}

// ============================================================
// 跨廠比較功能
// ============================================================

function initComparison() {
    // 簡單實作 - 可以根據需求擴充
    const btn = document.getElementById('compQueryBtn');
    if (btn) {
        btn.addEventListener('click', () => {
            showToast('功能開發中...', 'info');
        });
    }
}

// ============================================================
// 報表匯出功能
// ============================================================

function initReports() {
    // 簡單實作 - 可以根據需求擴充
    const btn = document.getElementById('generatePdfBtn');
    if (btn) {
        btn.addEventListener('click', () => {
            showToast('PDF 下載請求已發送', 'info');
        });
    }
}

// ============================================================
// 工具函式
// ============================================================

function formatDate(date) {
    if (!date) return '';
    const d = new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDateTime(datetimeStr) {
    if (!datetimeStr) return '';
    // 如果是 ISO 字串，轉換格式
    const d = new Date(datetimeStr);
    return d.toLocaleString('zh-TW', { hour12: false }).replace(/\//g, '-');
}

function formatNumber(num) {
    if (num === null || num === undefined) return '0';
    return Number(num).toLocaleString('zh-TW');
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${message}</span>`; // 可以加 icon

    container.appendChild(toast);

    // 3秒後消失
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(20px)';
        setTimeout(() => {
            container.removeChild(toast);
        }, 300);
    }, 3000);
}

function handleError(error) {
    console.error('系統錯誤:', error);
    showToast('系統發生錯誤，請稍後再試', 'error');
}
