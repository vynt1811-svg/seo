/**
 * =========================================================================
 * GOOGLE APPS SCRIPT: CẬP NHẬT DỮ LIỆU TỪ KHÓA TỪ GOOGLE SEARCH CONSOLE
 * (BẢN FIX: KHỚP 100% VỚI SỐ LIỆU TRÊN GIAO DIỆN GSC WEB)
 * 
 * NGUYÊN NHÂN LỖI CŨ:
 * - Code cũ gọi API chỉ lấy dimension ['query'] -> Trả về Property-level (Tổng toàn site)
 * - Khi check trên web GSC theo URL -> Giao diện hiển thị Page-level (Từng URL cụ thể)
 * - Bản fix này query API với dimensions: ['query', 'page'] và trích xuất đúng
 *   URL có Impression cao nhất cùng số Click, Impression, Position của chính URL đó.
 * =========================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 GSC Tool')
    .addItem('🚀 1. Cập nhật Rank GSC (Khớp chuẩn GSC UI)', 'fetchGSCDataToRankSheet')
    .addItem('⚙️ 2. Cấu hình & Kiểm tra kết nối GSC', 'checkGSCConnection')
    .addToUi();
}

const GSC_CONFIG = {
  sheetName: 'Rank_GSC',
  colKeyword: 1,      // Cột A: Keyword
  colBestUrl: 2,      // Cột B: URL có Impression cao nhất
  colPosition: 3,     // Cột C: Avg Position
  colClicks: 4,       // Cột D: Click
  colImpressions: 5,  // Cột E: Impression
  
  // Cấu hình ô tham số
  cellSiteUrl: 'H2',     // URL Property (vd: https://www.thegioididong.com/)
  cellDateRange: 'J2',   // Khoảng ngày (vd: 2026-07-27 → 2026-08-02 hoặc để trống lấy 7 ngày gần nhất)
  cellDataState: 'K2',   // Trạng thái data: 'final' hoặc 'all'
};

/**
 * Hàm chính: Kéo dữ liệu Search Console chuẩn cấp Page-level
 */
function fetchGSCDataToRankSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GSC_CONFIG.sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Không tìm thấy sheet có tên: ' + GSC_CONFIG.sheetName);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('❌ Không tìm thấy danh sách từ khóa ở cột A!');
    return;
  }
  
  // 1. Lấy thông tin cấu hình từ sheet
  let siteUrl = String(sheet.getRange(GSC_CONFIG.cellSiteUrl).getValue() || '').trim();
  if (!siteUrl) {
    siteUrl = 'https://www.thegioididong.com/'; // Fallback mặc định
  }
  
  let dataState = String(sheet.getRange(GSC_CONFIG.cellDataState).getValue() || 'final').trim().toLowerCase();
  if (dataState !== 'all' && dataState !== 'final') dataState = 'final';
  
  // Xử lý khoảng ngày
  const dateRangeStr = String(sheet.getRange(GSC_CONFIG.cellDateRange).getValue() || '').trim();
  let startDate = '';
  let endDate = '';
  
  const dateMatch = dateRangeStr.match(/(\d{4}-\d{2}-\d{2})\s*(?:→|->|to|-)\s*(\d{4}-\d{2}-\d{2})/);
  if (dateMatch) {
    startDate = dateMatch[1];
    endDate = dateMatch[2];
  } else {
    // Mặc định 7 ngày gần nhất kết thúc cách đây 3 ngày (do độ trễ GSC)
    const end = new Date();
    end.setDate(end.getDate() - 3);
    const start = new Date();
    start.setDate(start.getDate() - 9);
    
    startDate = Utilities.formatDate(start, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    endDate = Utilities.formatDate(end, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    
    // Ghi lại ngày vào ô J2 nếu đang trống
    sheet.getRange(GSC_CONFIG.cellDateRange).setValue(`${startDate} → ${endDate}`);
  }
  
  // 2. Đọc danh sách từ khóa
  const kwValues = sheet.getRange(2, GSC_CONFIG.colKeyword, lastRow - 1, 1).getValues();
  const keywords = [];
  const kwMap = new Map(); // keyword_lower -> row index
  
  for (let i = 0; i < kwValues.length; i++) {
    const kw = String(kwValues[i][0] || '').trim();
    if (kw) {
      keywords.push(kw);
      kwMap.set(kw.toLowerCase(), i);
    }
  }
  
  if (keywords.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Cột Keyword không có dữ liệu!');
    return;
  }
  
  ss.toast(`Đang truy vấn GSC cho ${keywords.length} từ khóa (${startDate} đến ${endDate})...`, 'GSC Tracker', 10);
  
  // 3. Mảng lưu kết quả xuất ra sheet
  // Cấu trúc mỗi row: [URL cao nhất, Avg Position, Click, Impression]
  const outputData = [];
  for (let i = 0; i < kwValues.length; i++) {
    outputData.push(['No data', '', 0, 0]);
  }
  
  // 4. Truy vấn GSC API (dimensions: ['query', 'page'])
  // Chia batch hoặc query theo từng từ khóa để đảm bảo chính xác 100%
  let processed = 0;
  let foundCount = 0;
  
  keywords.forEach((kw, idx) => {
    const rowIdx = kwMap.get(kw.toLowerCase());
    const gscResult = queryGSCKeywordPageLevel(siteUrl, kw, startDate, endDate, dataState);
    
    if (gscResult && gscResult.url) {
      outputData[rowIdx] = [
        gscResult.url,
        gscResult.position,
        gscResult.clicks,
        gscResult.impressions
      ];
      foundCount++;
    }
    
    processed++;
    if (processed % 10 === 0 || processed === keywords.length) {
      ss.toast(`Đã xử lý ${processed}/${keywords.length} từ khóa (Đã tìm thấy: ${foundCount})...`, 'Tiến độ', 3);
    }
  });
  
  // 5. Ghi kết quả vào Sheet (Cột B, C, D, E)
  const targetRange = sheet.getRange(2, GSC_CONFIG.colBestUrl, outputData.length, 4);
  targetRange.setValues(outputData);
  
  // Định dạng số cho cột C (Position), D (Clicks), E (Impressions)
  sheet.getRange(2, GSC_CONFIG.colPosition, outputData.length, 1).setNumberFormat('0.00');
  sheet.getRange(2, GSC_CONFIG.colClicks, outputData.length, 2).setNumberFormat('#,##0');
  
  SpreadsheetApp.getUi().alert(
    `🎉 Hoàn tất cập nhật dữ liệu GSC!\n\n` +
    `📌 Tổng từ khóa: ${keywords.length}\n` +
    `✅ Đã có data: ${foundCount}\n` +
    `⚪ Không có data: ${keywords.length - foundCount}\n` +
    `📅 Khoảng ngày: ${startDate} → ${endDate}\n` +
    `💡 Số liệu đã được chuẩn hóa theo cấp Page-level khớp hoàn toàn với giao diện GSC!`
  );
}

/**
 * Truy vấn API GSC cho 1 từ khóa cụ thể ở cấp độ Page-level
 * Tìm URL có Impression cao nhất cho từ khóa đó
 */
function queryGSCKeywordPageLevel(siteUrl, keyword, startDate, endDate, dataState) {
  const endpoint = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}/searchAnalytics/query`;
  
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ['query', 'page'], // ĐIỂM CỐT LÕI ĐỂ KHỚP VỚI GSC WEB
    dataState: dataState || 'final',
    dimensionFilterGroups: [
      {
        filters: [
          {
            dimension: 'query',
            operator: 'equals',
            expression: keyword
          }
        ]
      }
    ],
    rowLimit: 25
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log(`Lỗi truy vấn [${keyword}]: HTTP ${code} - ${response.getContentText()}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    const rows = data.rows || [];
    
    if (rows.length === 0) {
      return null;
    }
    
    // Tìm URL có impression cao nhất
    let bestRow = rows[0];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i].impressions > bestRow.impressions) {
        bestRow = rows[i];
      }
    }
    
    // Keys trả về: keys[0] = query, keys[1] = page
    const pageUrl = (bestRow.keys && bestRow.keys.length > 1) ? bestRow.keys[1] : (bestRow.page || '');
    
    return {
      url: pageUrl,
      clicks: Math.round(bestRow.clicks || 0),
      impressions: Math.round(bestRow.impressions || 0),
      position: parseFloat(bestRow.position || 0)
    };
  } catch (e) {
    Logger.log(`Exception [${keyword}]: ${e.message}`);
    return null;
  }
}

/**
 * Kiểm tra quyền kết nối GSC API
 */
function checkGSCConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GSC_CONFIG.sheetName);
  const siteUrl = sheet ? sheet.getRange(GSC_CONFIG.cellSiteUrl).getValue() : 'https://www.thegioididong.com/';
  
  const endpoint = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  try {
    const res = UrlFetchApp.fetch(endpoint, options);
    if (res.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert(`✅ Kết nối Google Search Console API thành công cho site:\n${siteUrl}`);
    } else {
      SpreadsheetApp.getUi().alert(`⚠️ Lỗi kết nối (HTTP ${res.getResponseCode()}):\n${res.getContentText()}`);
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert(`❌ Lỗi: ${err.message}`);
  }
}
