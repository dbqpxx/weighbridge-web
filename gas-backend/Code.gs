/**
 * 焚化廠地磅系統月報表查詢分析統計系統
 * Google Apps Script 後端程式碼
 * 
 * 功能模組：
 * 1. doGet/doPost - Web App 進入點
 * 2. importData - 資料匯入模組
 * 3. queryData - 查詢統計模組
 * 4. generateReport - 報表產生模組
 */

// ============================================================
// 全域設定
// ============================================================

const CONFIG = {
  // 試算表 ID
  SPREADSHEET_ID: '1ycqwRqLZl2IIx_X9K_kMhdbQNXpazJ3ax0IdyUqyFPw',
  
  // 工作表名稱
  SHEETS: {
    MASTER: 'Master',           // 主資料表
    CONFIG: 'Config',           // 設定表
    UPLOAD_LOG: 'UploadLog',    // 上傳記錄
    SOURCE_LIST: 'SourceList'   // 來源清單（區隊/廠商）
  },
  
  // 廠區定義
  PLANTS: {
    'south': '南區廠',
    'renwu': '仁武廠',
    'gangshan': '岡山廠'
  },
  
  // 車道編碼對應
  LANES: {
    '3': '車道1',
    '4': '車道2',
    '5': '車道3',
    '6': '車道4'
  },
  
  // 固定區隊清單（高雄市各區）
  DISTRICTS: [
    '前鎮', '小港', '鳳山', '苓雅', '新興', '前金', '鹽埕', '鼓山',
    '旗津', '三民', '左營', '楠梓', '大社', '仁武', '鳥松', '大樹',
    '岡山', '橋頭', '燕巢', '田寮', '阿蓮', '路竹', '湖內', '茄萣',
    '永安', '彌陀', '梓官', '旗山', '美濃', '六龜', '甲仙', '杉林',
    '內門', '茂林', '桃源', '那瑪夏', '林園', '大寮'
  ],
  
  // 廠商特徵關鍵字
  VENDOR_KEYWORDS: ['公司', '企業', '有限', '股份', '行', '社', '廠', '處理', '工程', '清潔', '環保'],
  
  // 垃圾種類清單
  WASTE_TYPES: [
    '一般廢棄物',
    '事業廢棄物',
    '廚餘',
    '廚餘焚化',
    '底渣',
    '底渣生垃圾回運',
    '水肥(轉入)',
    '水肥轉運(轉出)',
    '固化物',
    '環保局巨大垃圾',
    '區隊廚餘運入',
    '公務單位一般事業廢棄物(免費)'
  ]
};

// ============================================================
// Web App 進入點
// ============================================================

/**
 * GET 請求處理 - 返回 HTML 頁面或 JSON 資料
 */
function doGet(e) {
  // 檢查參數是否存在（從編輯器直接執行時 e 會是 undefined）
  const action = e && e.parameter ? e.parameter.action : null;
  
  if (action) {
    // API 模式：返回 JSON 資料
    return handleApiRequest(e);
  }
  
  // 網頁模式：返回 HTML 頁面
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('焚化廠地磅查詢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * POST 請求處理 - 資料上傳
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch (action) {
      case 'import':
        return jsonResponse(importData(data));
      case 'query':
        return jsonResponse(queryData(data));
      default:
        return jsonResponse({ success: false, error: '未知的操作' });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

/**
 * API 請求處理
 */
function handleApiRequest(e) {
  const action = e.parameter.action;
  
  try {
    // 最簡單的測試 - 不需要任何權限
    if (action === 'simpleTest') {
      return jsonResponse(simpleTest());
    }
    
    // 簡單的 ping 測試
    if (action === 'ping') {
      return jsonResponse({
        success: true,
        message: 'pong',
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        timestamp: new Date().toISOString()
      });
    }
    
    if (action === 'getConfig') {
      return jsonResponse(getConfig());
    }
    
    if (action === 'getSummary') {
      return jsonResponse(getSummary(e.parameter));
    }
    
    if (action === 'queryData') {
      return jsonResponse(queryData(e.parameter));
    }
    
    if (action === 'getPlantComparison') {
      return jsonResponse(getPlantComparison(e.parameter));
    }
    
    if (action === 'getWasteTypeStats') {
      return jsonResponse(getWasteTypeStats(e.parameter));
    }
    
    if (action === 'getSourceRanking') {
      return jsonResponse(getSourceRanking(e.parameter));
    }
    
    if (action === 'getDailyTrend') {
      return jsonResponse(getDailyTrend(e.parameter));
    }
    
    if (action === 'getSourceList') {
      return jsonResponse(getSourceList(e.parameter));
    }
    
    return jsonResponse({ 
      success: false, 
      error: '未知的操作: ' + action 
    });
    
  } catch (error) {
    Logger.log('handleApiRequest 錯誤: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return jsonResponse({ 
      success: false, 
      error: error.message,
      stack: error.stack
    });
  }
}

/**
 * JSON 回應包裝
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 載入 HTML 片段（用於模板引入）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 設定模組
// ============================================================

/**
 * 取得系統設定
 */
function getConfig() {
  try {
    return {
      success: true,
      data: {
        plants: CONFIG.PLANTS,
        lanes: CONFIG.LANES,
        districts: CONFIG.DISTRICTS,
        wasteTypes: CONFIG.WASTE_TYPES,
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        // 診斷信息
        canAccessSpreadsheet: testSpreadsheetAccess()
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * 測試試算表存取權限
 */
function testSpreadsheetAccess() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    return {
      success: true,
      name: ss.getName(),
      id: ss.getId()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================
// 資料匯入模組
// ============================================================

/**
 * 匯入月報表資料
 * @param {Object} params - 匯入參數
 * @param {string} params.plant - 廠區代碼
 * @param {string} params.yearMonth - 年月 (例如: 115-01)
 * @param {Array} params.data - CSV 資料陣列
 */
function importData(params) {
  const { plant, yearMonth, data } = params;
  
  if (!plant || !yearMonth || !data || !data.length) {
    return { success: false, error: '缺少必要參數' };
  }
  
  if (!CONFIG.PLANTS[plant]) {
    return { success: false, error: '無效的廠區代碼' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let masterSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    // 如果 Master 表不存在，建立它
    if (!masterSheet) {
      masterSheet = createMasterSheet(ss);
    }
    
    // 刪除該廠區該月份的舊資料
    deleteExistingData(masterSheet, plant, yearMonth);
    
    // 準備新資料
    const rows = [];
    const importTime = new Date();
    
    for (let i = 1; i < data.length; i++) { // 跳過標題列
      const row = data[i];
      if (!row || row.length < 8 || !row[0]) continue; // 跳過空行
      
      const seqNo = String(row[0]).trim();
      const laneCode = seqNo.charAt(0);
      
      rows.push([
        plant,                           // 廠區
        yearMonth,                       // 年月
        seqNo,                          // 序號
        CONFIG.LANES[laneCode] || '',   // 車道
        String(row[1] || '').trim(),    // 車號
        parseDateTime(row[2]),          // 時間
        String(row[3] || '').trim(),    // 區隊/廠商
        String(row[4] || '').trim(),    // 垃圾種類
        parseNumber(row[5]),            // 總重
        parseNumber(row[6]),            // 空重
        parseNumber(row[7]),            // 淨重
        parseNumber(row[8]),            // 金額
        String(row[9] || '').trim(),    // 備註
        importTime                      // 匯入時間
      ]);
    }
    
    if (rows.length === 0) {
      return { success: false, error: '沒有有效的資料列' };
    }
    
    // 寫入資料
    const lastRow = masterSheet.getLastRow();
    masterSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    // 記錄上傳日誌
    logUpload(ss, plant, yearMonth, rows.length, importTime);
    
    // 更新來源清單（擷取所有區隊/廠商）
    const sources = rows.map(r => r[6]); // 第 7 欄是區隊/廠商
    const newSourceCount = updateSourceList(sources, plant);
    
    // 計算本次匯入的總重量用於日誌除錯
    const totalImportWeight = rows.reduce((sum, r) => sum + r[10], 0);
    Logger.log(`[IMPORT] ${plant} ${yearMonth} 成功匯入 ${rows.length} 筆, 總淨重: ${totalImportWeight} KG`);
    
    return {
      success: true,
      message: `成功匯入 ${rows.length} 筆資料` + (newSourceCount > 0 ? `，新增 ${newSourceCount} 個來源` : ''),
      count: rows.length,
      totalWeightKg: totalImportWeight,
      newSources: newSourceCount
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 建立 Master 資料表
 */
function createMasterSheet(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEETS.MASTER);
  
  // 設定標題列
  const headers = [
    '廠區', '年月', '序號', '車道', '車號', '時間',
    '區隊/廠商', '垃圾種類', '總重', '空重', '淨重', '金額', '備註', '匯入時間'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // 凍結標題列
  sheet.setFrozenRows(1);
  
  return sheet;
}

/**
 * 刪除既有資料（同廠區同月份）- 高效能批次版
 */
function deleteExistingData(sheet, plant, yearMonth) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetKey = String(plant).trim() + "_" + String(yearMonth).trim();
  
  // 使用 Filter 過濾不需要刪除的資料
  const retainedData = data.filter((row, index) => {
    if (index === 0) return false; // 移除標題（另外寫入）
    const rowKey = String(row[0]).trim() + "_" + String(row[1]).trim();
    return rowKey !== targetKey;
  });
  
  const deleteCount = (data.length - 1) - retainedData.length;
  Logger.log(`批次刪除: 共 ${data.length-1} 筆, 刪除 ${deleteCount} 筆, 剩餘 ${retainedData.length} 筆`);
  
  if (deleteCount > 0) {
    // 清除內容並重寫內容
    sheet.clearContents();
    const finalData = [headers].concat(retainedData);
    sheet.getRange(1, 1, finalData.length, headers.length).setValues(finalData);
    
    // 重新設定標題格式（clearContents 會清除，但不會清除樣式，保險起見可用 clearContents 僅清內容）
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }
}

/**
 * 記錄上傳日誌
 */
function logUpload(ss, plant, yearMonth, count, time) {
  let logSheet = ss.getSheetByName(CONFIG.SHEETS.UPLOAD_LOG);
  
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.UPLOAD_LOG);
    logSheet.getRange(1, 1, 1, 5).setValues([['時間', '廠區', '年月', '筆數', '操作者']]);
    logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  const user = Session.getActiveUser().getEmail() || '系統';
  logSheet.appendRow([time, CONFIG.PLANTS[plant], yearMonth, count, user]);
}

/**
 * 解析日期時間
 */
function parseDateTime(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  
  let str = String(value).trim();
  
  // 處理民國年格式: 114/01/01 或 114-01-01
  const rocMatch = str.match(/^(\d{2,3})([\\/\\-])(\d{1,2})([\\/\\-])(\d{1,2})(.*)/);
  if (rocMatch) {
    const rocYear = parseInt(rocMatch[1]);
    if (rocYear < 200) { // 民國年特徵
      const adYear = rocYear + 1911;
      str = adYear + rocMatch[2] + rocMatch[3] + rocMatch[4] + rocMatch[5] + rocMatch[6];
    }
  }
  
  // 1. 嘗試解析 YYYY/MM/DD HH:mm:ss 或 YYYY/MM/DD HH:mm 格式
  const match1 = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(\d{1,2}):(\d{2})(:(\d{2}))?/);
  if (match1) {
    return new Date(
      parseInt(match1[1]),
      parseInt(match1[2]) - 1,
      parseInt(match1[3]),
      parseInt(match1[4]),
      parseInt(match1[5]),
      match1[7] ? parseInt(match1[7]) : 0
    );
  }
  
  // 2. 嘗試解析 YYYY/MM/DD 或 YYYY-MM-DD (不含時間)
  const match2 = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (match2) {
    return new Date(parseInt(match2[1]), parseInt(match2[2]) - 1, parseInt(match2[3]));
  }
  
  // 3. 基本 Date 解析
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  
  return str;
}

/**
 * 解析數字（移除千分位逗號）
 */
function parseNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  
  // 移除逗號及括號
  const cleanStr = String(value).replace(/,/g, '').replace(/[\(\)]/g, '');
  const num = parseFloat(cleanStr);
  return isNaN(num) ? 0 : num;
}

/**
 * 確保為 Date 物件（用於比對）- 時區安全版
 */
function ensureDate(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  // 處理 YYYY-MM-DD 格式
  if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return parseLocalDate(value);
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 解析 YYYY-MM-DD 為本地時區的 Date 物件 (00:00:00)
 */
function parseLocalDate(dateStr) {
  if (!dateStr) return null;
  
  // 支援民國年: 114/01/01 或 114-01-01
  const rocMatch = String(dateStr).match(/^(\d{2,3})([\\/\\-])(\d{1,2})([\\/\\-])(\d{1,2})$/);
  if (rocMatch) {
    const rocYear = parseInt(rocMatch[1]);
    const y = rocYear < 200 ? rocYear + 1911 : rocYear;
    const m = parseInt(rocMatch[3]) - 1;
    const d = parseInt(rocMatch[5]);
    return new Date(y, m, d);
  }

  const parts = String(dateStr).split(/[-\/]/);
  if (parts.length === 3) {
    const y = parseInt(parts[0]);
    const m = parseInt(parts[1]) - 1;
    const d = parseInt(parts[2]);
    return new Date(y, m, d);
  }
  return new Date(dateStr);
}

// ============================================================
// 查詢統計模組
// ============================================================

/**
 * 偵錯用：檢查 Master 表的實際內容
 * 請在 Apps Script 編輯器中選擇執行此函數，並查看執行記錄
 */
function debugMaster() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    if (!sheet) {
      Logger.log("錯誤: 找不到 Master 工作表");
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1) {
      Logger.log("Master 表是空的");
      return;
    }
    
    // 讀取前 5 筆資料（含標題）
    const limit = Math.min(lastRow, 5);
    const data = sheet.getRange(1, 1, limit, lastCol).getValues();
    
    Logger.log("--- Master 表格結構檢查 ---");
    Logger.log("標題列: " + JSON.stringify(data[0]));
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`第 ${i+1} 列資料: ` + JSON.stringify(row));
      Logger.log(`欄位 0 (廠區): 内容="${row[0]}", 類型=${typeof row[0]}`);
      Logger.log(`欄位 5 (時間): 内容="${row[5]}", 類型=${row[5] instanceof Date ? 'Date' : typeof row[5]}`);
      Logger.log(`欄位 10 (淨重): 内容="${row[10]}", 類型=${typeof row[10]}`);
    }
    
    return { success: true, message: "請查看 Apps Script 執行記錄" };
  } catch (e) {
    Logger.log("執行偵錯失敗: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 查詢資料
 * @param {Object} params - 查詢參數
 */
/**
 * 簡單測試 - 不需要任何權限
 */
function simpleTest() {
  return {
    success: true,
    message: "Web App 通訊正常！",
    timestamp: new Date().toISOString(),
    config: {
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      hasPlants: !!CONFIG.PLANTS
    }
  };
}

/**
 * 測試查詢 - 返回固定數據
 */
function queryDataTest() {
  return {
    success: true,
    data: [
      {
        plant: 'south',
        plantName: '南區廠',
        yearMonth: '115-01',
        seqNo: 1,
        lane: '車道1',
        vehicleNo: 'TEST-001',
        datetime: new Date('2026-01-01'),
        source: '測試來源',
        wasteType: '測試垃圾',
        grossWeight: 1000,
        tareWeight: 500,
        netWeight: 500,
        amount: 1000,
        remark: '測試數據'
      }
    ],
    total: 1
  };
}

/**
 * 查詢資料 - 參考 getSummary 的成功模式
 */
function queryData(params) {
  Logger.log('[queryData] Started with params: ' + JSON.stringify(params));
  const { startDate, endDate, plants, wasteTypes, source } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      Logger.log('[queryData] Error: Master sheet not found');
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    // Check data range size
    const lastRow = sheet.getLastRow();
    Logger.log('[queryData] Last row: ' + lastRow);
    
    if (lastRow < 2) {
       return { success: true, data: [], total: 0 };
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('[queryData] Data fetched, rows: ' + data.length);

    const plantList = parseList(plants);
    const wasteTypeList = parseList(wasteTypes);
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    const results = [];
    const limit = 2000; // Match frontend limit
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 廠區過濾
      if (plantList.length > 0 && !plantList.includes(String(row[0] || '').trim())) continue;
      
      // 日期過濾
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      // 垃圾種類過濾
      if (wasteTypeList.length > 0 && !wasteTypeList.includes(String(row[7] || '').trim())) continue;
      
      // 來源過濾
      if (source && String(row[6] || '').indexOf(source) === -1) continue;
      
      results.push({
        plant: row[0],
        plantName: CONFIG.PLANTS[row[0]] || row[0],
        yearMonth: row[1],
        seqNo: row[2],
        lane: row[3],
        vehicleNo: row[4],
        datetime: row[5],
        source: row[6],
        wasteType: row[7],
        grossWeight: row[8],
        tareWeight: row[9],
        netWeight: row[10],
        amount: row[11],
        remark: row[12]
      });
      
      if (results.length >= limit) break;
    }
    
    Logger.log('[queryData] Query completed, results found: ' + results.length);
    
    return JSON.stringify({
      success: true,
      data: results,
      total: results.length
    });
    
  } catch (error) {
    Logger.log('[queryData] Error: ' + error.stack);
    return JSON.stringify({ success: false, error: '系統錯誤: ' + error.message });
  }
}

/**
 * 取得統計摘要
 */
function getSummary(params) {
  const { startDate, endDate, plants } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    const data = sheet.getDataRange().getValues();
    const plantList = parseList(plants);
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    let totalRecords = 0;
    let totalNetWeight = 0;
    let totalAmount = 0;
    const plantStats = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (plantList.length > 0 && !plantList.includes(String(row[0] || '').trim())) continue;
      
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      const plant = row[0];
      const netWeight = row[10] || 0;
      const amount = row[11] || 0;
      
      totalRecords++;
      totalNetWeight += netWeight;
      totalAmount += amount;
      
      if (!plantStats[plant]) {
        plantStats[plant] = { records: 0, netWeight: 0, amount: 0 };
      }
      plantStats[plant].records++;
      plantStats[plant].netWeight += netWeight;
      plantStats[plant].amount += amount;
    }
    
    return {
      success: true,
      data: {
        totalRecords,
        totalNetWeight,
        totalNetWeightTon: totalNetWeight / 1000,
        totalAmount,
        plantStats: Object.entries(plantStats).map(([code, stats]) => ({
          code,
          name: CONFIG.PLANTS[code] || code,
          ...stats,
          netWeightTon: stats.netWeight / 1000
        }))
      }
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 取得跨廠比較資料
 */
function getPlantComparison(params) {
  const { startDate, endDate } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    const data = sheet.getDataRange().getValues();
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    const comparison = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      const plant = row[0];
      const wasteType = row[7];
      const netWeight = row[10] || 0;
      
      if (!comparison[plant]) {
        comparison[plant] = {};
      }
      if (!comparison[plant][wasteType]) {
        comparison[plant][wasteType] = 0;
      }
      comparison[plant][wasteType] += netWeight;
    }
    
    return {
      success: true,
      data: comparison
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 取得垃圾種類統計
 */
function getWasteTypeStats(params) {
  const { startDate, endDate, plants } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    const data = sheet.getDataRange().getValues();
    const plantList = parseList(plants);
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    const stats = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (plantList.length > 0 && !plantList.includes(row[0])) continue;
      
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      const wasteType = row[7];
      const netWeight = row[10] || 0;
      const amount = row[11] || 0;
      
      if (!stats[wasteType]) {
        stats[wasteType] = { count: 0, netWeight: 0, amount: 0 };
      }
      stats[wasteType].count++;
      stats[wasteType].netWeight += netWeight;
      stats[wasteType].amount += amount;
    }
    
    const result = Object.entries(stats)
      .map(([type, data]) => ({
        wasteType: type,
        count: data.count,
        netWeight: data.netWeight,
        netWeightTon: data.netWeight / 1000,
        amount: data.amount
      }))
      .sort((a, b) => b.netWeight - a.netWeight);
    
    return {
      success: true,
      data: result
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 取得區隊/廠商排名
 */
function getSourceRanking(params) {
  const { startDate, endDate, plants, wasteTypes, limit = 20 } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    const data = sheet.getDataRange().getValues();
    const plantList = parseList(plants);
    const wasteTypeList = parseList(wasteTypes);
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    const ranking = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (plantList.length > 0 && !plantList.includes(row[0])) continue;
      if (wasteTypeList.length > 0 && !wasteTypeList.includes(row[7])) continue;
      
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      const source = row[6];
      const netWeight = row[10] || 0;
      const amount = row[11] || 0;
      
      if (!ranking[source]) {
        ranking[source] = { count: 0, netWeight: 0, amount: 0 };
      }
      ranking[source].count++;
      ranking[source].netWeight += netWeight;
      ranking[source].amount += amount;
    }
    
    const result = Object.entries(ranking)
      .map(([source, data]) => ({
        source,
        count: data.count,
        netWeight: data.netWeight,
        netWeightTon: data.netWeight / 1000,
        amount: data.amount
      }))
      .sort((a, b) => b.netWeight - a.netWeight)
      .slice(0, limit);
    
    return {
      success: true,
      data: result
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 取得每日趨勢
 */
function getDailyTrend(params) {
  const { startDate, endDate, plants, wasteTypes } = params;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    
    if (!sheet) {
      return { success: false, error: 'Master 資料表不存在' };
    }
    
    const data = sheet.getDataRange().getValues();
    const plantList = parseList(plants);
    const wasteTypeList = parseList(wasteTypes);
    const startDateObj = parseLocalDate(startDate);
    let endDateObj = parseLocalDate(endDate);
    if (endDateObj) endDateObj.setHours(23, 59, 59, 999);
    
    const trend = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (plantList.length > 0 && !plantList.includes(row[0])) continue;
      if (wasteTypeList.length > 0 && !wasteTypeList.includes(row[7])) continue;
      
      const rowDate = ensureDate(row[5]);
      if (!rowDate) continue;
      
      if (startDateObj && rowDate < startDateObj) continue;
      if (endDateObj && rowDate > endDateObj) continue;
      
      // 取得日期字串 (YYYY-MM-DD)
      const dateKey = Utilities.formatDate(
        rowDate instanceof Date ? rowDate : new Date(rowDate),
        'Asia/Taipei',
        'yyyy-MM-dd'
      );
      
      const netWeight = row[10] || 0;
      const amount = row[11] || 0;
      
      if (!trend[dateKey]) {
        trend[dateKey] = { count: 0, netWeight: 0, amount: 0 };
      }
      trend[dateKey].count++;
      trend[dateKey].netWeight += netWeight;
      trend[dateKey].amount += amount;
    }
    
    const result = Object.entries(trend)
      .map(([date, data]) => ({
        date,
        count: data.count,
        netWeight: data.netWeight,
        netWeightTon: data.netWeight / 1000,
        amount: data.amount
      }))
      .sort((a, b) => a.date.localeCompare(b.date));
    
    return {
      success: true,
      data: result
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 解析清單（陣列或逗號分隔字串）
 */
function parseList(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.filter(s => s && s !== 'null' && s !== 'undefined');
  return String(value).split(',')
    .map(s => s.trim())
    .filter(s => s && s !== 'null' && s !== 'undefined');
}

// ============================================================
// 報表產生模組
// ============================================================

/**
 * 產生月報表 PDF
 */
function generateMonthlyReport(plant, yearMonth) {
  // TODO: 使用 Docs API 產生 PDF 報表
  return { success: false, error: '功能開發中' };
}

/**
 * 匯出 Excel
 */
function exportToExcel(params) {
  // TODO: 產生 Excel 下載連結
  return { success: false, error: '功能開發中' };
}

/**
 * 測試基本連接 - 直接在編輯器中執行此函數
 * 執行後請查看"執行記錄"(Ctrl+Enter 或 View > Logs)
 */
function testBasicConnection() {
  Logger.log("=== 開始測試 ===");
  
  // 1. 檢查 SPREADSHEET_ID
  Logger.log("SPREADSHEET_ID: " + CONFIG.SPREADSHEET_ID);
  
  if (CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    Logger.log("❌ 錯誤: SPREADSHEET_ID 尚未設定！");
    Logger.log("請在 Code.gs 的第 18 行設定您的 Google Sheets ID");
    return;
  }
  
  try {
    // 2. 嘗試開啟試算表
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log("✓ 成功開啟試算表: " + ss.getName());
    
    // 3. 檢查 Master 表
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    if (!sheet) {
      Logger.log("❌ 錯誤: 找不到 'Master' 工作表");
      Logger.log("現有工作表: " + ss.getSheets().map(s => s.getName()).join(", "));
      return;
    }
    
    Logger.log("✓ 找到 Master 工作表");
    
    // 4. 檢查數據
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log(`Master 表大小: ${lastRow} 列 x ${lastCol} 欄`);
    
    if (lastRow <= 1) {
      Logger.log("❌ Master 表只有標題列（或是空的）");
      return;
    }
    
    // 5. 讀取前 3 筆數據
    const data = sheet.getRange(1, 1, Math.min(4, lastRow), lastCol).getValues();
    Logger.log("\n標題列:");
    Logger.log(JSON.stringify(data[0]));
    
    Logger.log("\n前 3 筆數據:");
    for (let i = 1; i < data.length; i++) {
      Logger.log(`第 ${i} 筆: ${JSON.stringify(data[i])}`);
    }
    
    Logger.log("\n=== 測試完成 ===");
    Logger.log(`✓ 總共有 ${lastRow - 1} 筆數據`);
    
  } catch (e) {
    Logger.log("❌ 發生錯誤: " + e.message);
  }
}

/**
 * 測試 queryData 函數
 */
function testQueryData() {
  Logger.log("=== 測試 queryData ===");
  
  // 調用 queryData（傳入空參數，應該返回前 500 筆）
  const result = queryData({});
  
  Logger.log("返回結果:");
  Logger.log("success: " + result.success);
  
  if (result.success) {
    Logger.log("total: " + result.total);
    Logger.log("data.length: " + result.data.length);
    
    if (result.data.length > 0) {
      Logger.log("\n第一筆數據:");
      Logger.log(JSON.stringify(result.data[0], null, 2));
    }
    
    if (result.debug) {
      Logger.log("\ndebug 資訊:");
      Logger.log(JSON.stringify(result.debug, null, 2));
    }
  } else {
    Logger.log("❌ 錯誤: " + result.error);
  }
  
  Logger.log("\n=== 測試完成 ===");
}

// ============================================================
// 初始化函數
// ============================================================

/**
 * 初始化資料庫結構
 */
function initializeDatabase() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 建立 Master 資料表
  if (!ss.getSheetByName(CONFIG.SHEETS.MASTER)) {
    createMasterSheet(ss);
    Logger.log('已建立 Master 資料表');
  }
  
  // 建立 Config 設定表
  if (!ss.getSheetByName(CONFIG.SHEETS.CONFIG)) {
    const configSheet = ss.insertSheet(CONFIG.SHEETS.CONFIG);
    configSheet.getRange('A1').setValue('plants');
    configSheet.getRange('B1').setValue(JSON.stringify(CONFIG.PLANTS));
    configSheet.getRange('A2').setValue('wasteTypes');
    configSheet.getRange('B2').setValue(JSON.stringify(CONFIG.WASTE_TYPES));
    configSheet.getRange('A3').setValue('lanes');
    configSheet.getRange('B3').setValue(JSON.stringify(CONFIG.LANES));
    configSheet.getRange('A4').setValue('districts');
    configSheet.getRange('B4').setValue(JSON.stringify(CONFIG.DISTRICTS));
    Logger.log('已建立 Config 設定表');
  }
  
  // 建立 UploadLog 上傳記錄表
  if (!ss.getSheetByName(CONFIG.SHEETS.UPLOAD_LOG)) {
    const logSheet = ss.insertSheet(CONFIG.SHEETS.UPLOAD_LOG);
    logSheet.getRange(1, 1, 1, 5).setValues([['時間', '廠區', '年月', '筆數', '操作者']]);
    logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
    Logger.log('已建立 UploadLog 上傳記錄表');
  }
  
  // 建立 SourceList 來源清單表
  if (!ss.getSheetByName(CONFIG.SHEETS.SOURCE_LIST)) {
    createSourceListSheet(ss);
    Logger.log('已建立 SourceList 來源清單表');
  }
  
  Logger.log('資料庫初始化完成');
}

/**
 * 建立 SourceList 工作表並預載入區隊
 */
function createSourceListSheet(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEETS.SOURCE_LIST);
  
  // 設定標題列
  const headers = ['來源名稱', '類型', '廠區', '首次出現', '累計車次'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // 預載入固定區隊
  const now = new Date();
  const districtRows = CONFIG.DISTRICTS.map(d => [d, 'district', '', now, 0]);
  
  if (districtRows.length > 0) {
    sheet.getRange(2, 1, districtRows.length, 5).setValues(districtRows);
  }
  
  // 凍結標題列
  sheet.setFrozenRows(1);
  
  return sheet;
}

// ============================================================
// 來源管理模組
// ============================================================

/**
 * 取得來源清單
 * @param {Object} params - 參數
 * @param {string} params.type - 類型篩選：district（區隊）、vendor（廠商）、空白為全部
 * @param {string} params.plant - 廠區篩選
 */
function getSourceList(params) {
  const { type, plant } = params || {};
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEETS.SOURCE_LIST);
    
    // 如果 SourceList 不存在，建立它
    if (!sheet) {
      Logger.log('SourceList 不存在，正在建立...');
      sheet = createSourceListSheet(ss);
    }
    
    const data = sheet.getDataRange().getValues();
    const results = { districts: [], vendors: [] };
    
    Logger.log('SourceList 資料列數: ' + (data.length - 1));
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sourceName = row[0];
      const sourceType = row[1];
      const sourcePlant = row[2];
      
      if (!sourceName) continue; // 跳過空白列
      
      // 廠區篩選
      if (plant && sourcePlant && sourcePlant !== plant) continue;
      
      // 類型篩選
      if (type && sourceType !== type) continue;
      
      if (sourceType === 'district') {
        results.districts.push({
          name: sourceName,
          plant: sourcePlant || '',
          recordCount: row[4] || 0
        });
      } else {
        results.vendors.push({
          name: sourceName,
          plant: sourcePlant || '',
          recordCount: row[4] || 0
        });
      }
    }
    
    // 排序：按車次數量降序，車次為 0 的放後面
    results.districts.sort((a, b) => {
      if (a.recordCount === 0 && b.recordCount === 0) return a.name.localeCompare(b.name);
      return b.recordCount - a.recordCount;
    });
    results.vendors.sort((a, b) => b.recordCount - a.recordCount);
    
    Logger.log('返回區隊數: ' + results.districts.length + ', 廠商數: ' + results.vendors.length);
    
    return {
      success: true,
      data: results
    };
    
  } catch (error) {
    Logger.log('getSourceList 錯誤: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 分類來源（區隊或廠商）
 * @param {string} name - 來源名稱
 * @returns {string} 'district' 或 'vendor'
 */
function classifySource(name) {
  if (!name) return 'vendor';
  
  // 先移除後綴標註再判斷
  const cleanName = String(name).replace(/[\(（]車次[\)）]/g, '').trim();
  
  // 檢查是否在固定區隊清單中
  if (CONFIG.DISTRICTS.includes(cleanName)) {
    return 'district';
  }
  
  // 檢查是否包含廠商關鍵字
  for (const keyword of CONFIG.VENDOR_KEYWORDS) {
    if (cleanName.includes(keyword)) {
      return 'vendor';
    }
  }
  
  // 檢查名稱長度（區隊通常 2-4 字）
  if (cleanName.length <= 4 && !cleanName.includes('(') && !cleanName.includes('（')) {
    return 'district';
  }
  
  return 'vendor';
}

/**
 * 更新來源清單（匯入資料時呼叫）
 * @param {string[]} sources - 來源名稱陣列
 * @param {string} plant - 廠區代碼
 */
function updateSourceList(sources, plant) {
  if (!sources || sources.length === 0) return;
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SOURCE_LIST);
  
  if (!sheet) {
    sheet = createSourceListSheet(ss);
  }
  
  const data = sheet.getDataRange().getValues();
  const existingMap = new Map();
  
  // 建立現有來源的 Map
  for (let i = 1; i < data.length; i++) {
    existingMap.set(data[i][0], { row: i + 1, count: data[i][4] || 0 });
  }
  
  // 統計新來源的車次
  const sourceCount = {};
  sources.forEach(s => {
    const name = String(s).trim();
    if (name) {
      sourceCount[name] = (sourceCount[name] || 0) + 1;
    }
  });
  
  const now = new Date();
  const newRows = [];
  
  // 處理每個來源
  for (const [name, count] of Object.entries(sourceCount)) {
    if (existingMap.has(name)) {
      // 更新現有來源的車次
      const existing = existingMap.get(name);
      const newCount = existing.count + count;
      sheet.getRange(existing.row, 5).setValue(newCount);
      
      // 更新廠區（如果之前沒有）
      if (!data[existing.row - 1][2] && plant) {
        sheet.getRange(existing.row, 3).setValue(plant);
      }
    } else {
      // 新增來源
      const sourceType = classifySource(name);
      newRows.push([name, sourceType, plant, now, count]);
      existingMap.set(name, { row: data.length + newRows.length, count: count });
    }
  }
  
  // 批次寫入新來源
  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
  }
  
  return newRows.length;
}

// ============================================================
// 測試函數
// ============================================================

/**
 * 測試查詢功能
 */
function testQuery() {
  const result = getSummary({
    startDate: '2026-01-01',
    endDate: '2026-01-31'
  });
  Logger.log(JSON.stringify(result, null, 2));
}
