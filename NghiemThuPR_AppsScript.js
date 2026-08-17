/**
 * =========================================================================
 * TOOL NGHIỆM THU LINK BÁO PR - GOOGLE APPS SCRIPT (BẢN TỰ ĐỘNG HÓA CAO CẤP)
 * - Tự động gom data từ 5 sheet nguồn và LỌC BỎ TRÙNG LẶP
 * - Tự động nhận diện bài mới (chưa check) để nghiệm thu
 * =========================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Nghiệm Thu Link PR')
    .addItem('📥 1. TỰ ĐỘNG GOM DATA MỚI TỪ 5 SHEET (Không trùng bài cũ)', 'syncNewDataFromSheets')
    .addSeparator()
    .addItem('▶️ 2. Chỉ kiểm tra các DÒNG MỚI (Chưa check / Lỗi)', 'checkPendingOrFailedLinks')
    .addItem('▶️ 3. Kiểm tra CÁC DÒNG ĐANG CHỌN', 'checkSelectedRows')
    .addItem('▶️ 4. Kiểm tra TOÀN BỘ danh sách', 'checkAllLinks')
    .addSeparator()
    .addItem('🧹 5. Lọc & Xóa các dòng bị trùng link bài báo', 'removeDuplicateLinks')
    .addItem('🔄 6. Xóa toàn bộ kết quả kiểm tra cũ để chạy lại', 'clearCheckResults')
    .addToUi();
}

/**
 * Cấu hình hệ thống
 */
const CONFIG = {
  targetSheetName: 'Nghiem-thu',
  sourceSheets: ['VTDĐ', 'Apple', 'Laptop', 'Phụ Kiện', 'Đồng Hồ'],
  colNhom: 1,       // Cột A: Nhóm NH
  colAnchor: 2,     // Cột B: AnchorText
  colUrlDich: 3,    // Cột C: URL đích
  colLinkBao: 4,    // Cột D: Link bài đăng trên báo
  colStatus: 5,     // Cột E: Trạng thái (PASS, CẢNH BÁO, LỖI)
  colDetail: 6,     // Cột F: Chi tiết nghiệm thu
  colTime: 7        // Cột G: Thời gian check
};

/**
 * TÍNH NĂNG 1: Tự động gom bài viết mới có link báo từ 5 sheet nguồn,
 * Tự động lọc trùng: Bài nào đã có trong sheet Nghiem-thu rồi thì KHÔNG thêm lại.
 */
function syncNewDataFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(CONFIG.targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(CONFIG.targetSheetName);
  }
  ensureHeaders(targetSheet);

  // Lấy danh sách link bài báo hiện có trong sheet Nghiem-thu để đối chiếu trùng
  const lastRow = targetSheet.getLastRow();
  const existingLinksSet = new Set();
  
  if (lastRow > 1) {
    const existingData = targetSheet.getRange(2, CONFIG.colLinkBao, lastRow - 1, 1).getValues();
    for (let i = 0; i < existingData.length; i++) {
      const link = cleanUrlString(String(existingData[i][0] || ''));
      if (link) {
        existingLinksSet.add(link.toLowerCase());
      }
    }
  }

  const newRowsToAdd = [];

  // Quét qua từng sheet nguồn
  CONFIG.sourceSheets.forEach(sheetName => {
    const sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet) return;

    const srcLastRow = sourceSheet.getLastRow();
    if (srcLastRow < 2) return;

    const values = sourceSheet.getDataRange().getValues();
    const header = values[0];

    // Xác định vị trí các cột
    let colNhomIdx = 0;
    let colAnchorIdx = -1;
    let colUrlDichIdx = -1;
    let colLinkBaoIdx = -1;

    for (let c = 0; c < header.length; c++) {
      const colTitle = String(header[c]).toLowerCase().trim();
      if (colTitle.includes('nhóm nh') || colTitle.includes('nhom nh')) colNhomIdx = c;
      else if (colTitle.includes('anchortext') || colTitle.includes('anchor')) colAnchorIdx = c;
      else if (colTitle.includes('url đích') || colTitle.includes('url dich') || colTitle.includes('url_đích')) colUrlDichIdx = c;
      else if (colTitle.includes('link bài đăng trên báo') || colTitle.includes('link bài đăng')) colLinkBaoIdx = c;
    }

    if (colLinkBaoIdx === -1) return;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const linkBaoRaw = String(row[colLinkBaoIdx] || '').trim();
      const cleanLink = cleanUrlString(linkBaoRaw);

      // Nếu dòng có link báo và CHƯA CÓ trong sheet Nghiem-thu
      if (cleanLink && !cleanLink.startsWith('#') && cleanLink !== '-' && !existingLinksSet.has(cleanLink.toLowerCase())) {
        const nhom = String(row[colNhomIdx] || '').trim() || sheetName;
        const anchor = colAnchorIdx !== -1 ? String(row[colAnchorIdx] || '').trim() : '';
        const urlDich = colUrlDichIdx !== -1 ? String(row[colUrlDichIdx] || '').trim() : '';

        newRowsToAdd.push([nhom, anchor, urlDich, cleanLink]);
        existingLinksSet.add(cleanLink.toLowerCase()); // Đánh dấu để tránh trùng chính nó
      }
    }
  });

  if (newRowsToAdd.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Dữ liệu đã là mới nhất! Không có link bài báo mới nào cần thêm.');
    return;
  }

  // Thêm các dòng mới vào cuối bảng
  const startAppendRow = targetSheet.getLastRow() + 1;
  targetSheet.getRange(startAppendRow, 1, newRowsToAdd.length, 4).setValues(newRowsToAdd);

  SpreadsheetApp.getUi().alert(
    `🎉 Đã gom thêm thành công ${newRowsToAdd.length} link bài báo mới vào cuối sheet!\n` +
    `👉 Các bài trùng đã được tự động loại bỏ.`
  );
}

/**
 * TÍNH NĂNG 2: Chỉ kiểm tra những dòng MỚI (chưa có kết quả check) hoặc bị LỖI/CẢNH BÁO
 */
function checkPendingOrFailedLinks() {
  runVerification({ mode: 'PENDING_OR_FAIL' });
}

/**
 * TÍNH NĂNG 3: Kiểm tra các dòng đang chọn bằng chuột
 */
function checkSelectedRows() {
  runVerification({ mode: 'SELECTED' });
}

/**
 * TÍNH NĂNG 4: Kiểm tra toàn bộ
 */
function checkAllLinks() {
  runVerification({ mode: 'ALL' });
}

/**
 * TÍNH NĂNG 5: Lọc & Xóa các dòng bị trùng link bài báo trong chính sheet Nghiem-thu
 */
function removeDuplicateLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.targetSheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const seenLinks = new Set();
  const rowsToKeep = [];
  let duplicateCount = 0;

  for (let i = 0; i < data.length; i++) {
    const link = cleanUrlString(String(data[i][CONFIG.colLinkBao - 1] || ''));
    if (!link) continue;

    const linkLower = link.toLowerCase();
    if (seenLinks.has(linkLower)) {
      duplicateCount++;
    } else {
      seenLinks.add(linkLower);
      rowsToKeep.push(data[i]);
    }
  }

  if (duplicateCount === 0) {
    SpreadsheetApp.getUi().alert('✅ Sheet không có link bài báo nào bị trùng!');
    return;
  }

  // Xóa và ghi lại
  sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  sheet.getRange(2, 1, rowsToKeep.length, 7).setValues(rowsToKeep);

  SpreadsheetApp.getUi().alert(`🧹 Đã xóa ${duplicateCount} dòng bị trùng link bài báo!`);
}

/**
 * TÍNH NĂNG 6: Xóa kết quả cột E, F, G
 */
function clearCheckResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.targetSheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, CONFIG.colStatus, lastRow - 1, 3).clearContent().setBackground(null);
    ss.toast('Đã xóa sạch kết quả kiểm tra cũ!', 'Thông báo', 3);
  }
}

/**
 * Hàm điều phối kiểm tra
 */
function runVerification(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.targetSheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Không tìm thấy sheet: ' + CONFIG.targetSheetName);
    return;
  }

  ensureHeaders(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Không có dữ liệu trong sheet!');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let targetRowIndices = [];

  if (options.mode === 'ALL') {
    for (let i = 0; i < data.length; i++) targetRowIndices.push(i);
  } else if (options.mode === 'SELECTED') {
    const selection = sheet.getActiveRange();
    if (!selection) {
      SpreadsheetApp.getUi().alert('Vui lòng chọn các dòng cần kiểm tra!');
      return;
    }
    const startRow = selection.getRow();
    const numRows = selection.getNumRows();
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r >= 2 && r <= lastRow) {
        targetRowIndices.push(r - 2);
      }
    }
  } else if (options.mode === 'PENDING_OR_FAIL') {
    // CHẾ ĐỘ QUAN TRỌNG: Chỉ check dòng nào CỘT TRẠNG THÁI ĐANG TRỐNG hoặc KHÔNG PHẢI LÀ PASS
    for (let i = 0; i < data.length; i++) {
      const status = String(data[i][CONFIG.colStatus - 1] || '').trim().toUpperCase();
      if (!status || status.indexOf('PASS') === -1) {
        targetRowIndices.push(i);
      }
    }
  }

  if (targetRowIndices.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Tất cả các dòng đều đã được check và đạt chuẩn (PASS)! Không có dòng mới cần kiểm tra.');
    return;
  }

  ss.toast('Bắt đầu kiểm tra ' + targetRowIndices.length + ' bài báo...', 'Đang xử lý', 5);

  let successCount = 0;
  let warningCount = 0;
  let errorCount = 0;

  for (let idx = 0; idx < targetRowIndices.length; idx++) {
    const rowIndex = targetRowIndices[idx];
    const sheetRowNum = rowIndex + 2;
    const row = data[rowIndex];

    const anchorTextRaw = String(row[CONFIG.colAnchor - 1] || '');
    const urlDichRaw = String(row[CONFIG.colUrlDich - 1] || '');
    const linkBaoRaw = String(row[CONFIG.colLinkBao - 1] || '').trim();

    if (!linkBaoRaw) continue;

    const result = verifyArticle(linkBaoRaw, anchorTextRaw, urlDichRaw);

    const statusRange = sheet.getRange(sheetRowNum, CONFIG.colStatus);
    const detailRange = sheet.getRange(sheetRowNum, CONFIG.colDetail);
    const timeRange = sheet.getRange(sheetRowNum, CONFIG.colTime);

    statusRange.setValue(result.status);
    detailRange.setValue(result.details.join('\n'));
    timeRange.setValue(new Date());

    if (result.status === 'PASS') {
      statusRange.setBackground('#D1FAE5').setFontColor('#065F46').setFontWeight('bold'); // Xanh lá
      successCount++;
    } else if (result.status === 'CẢNH BÁO') {
      statusRange.setBackground('#FEF3C7').setFontColor('#92400E').setFontWeight('bold'); // Vàng cam
      warningCount++;
    } else {
      statusRange.setBackground('#FEE2E2').setFontColor('#991B1B').setFontWeight('bold'); // Đỏ
      errorCount++;
    }

    if ((idx + 1) % 5 === 0 || idx === targetRowIndices.length - 1) {
      ss.toast(`Đã kiểm tra ${idx + 1}/${targetRowIndices.length} (PASS: ${successCount}, CẢNH BÁO: ${warningCount}, LỖI: ${errorCount})`, 'Tiến độ', 3);
      SpreadsheetApp.flush();
    }
  }

  SpreadsheetApp.getUi().alert(
    `🎉 Hoàn tất kiểm tra ${targetRowIndices.length} bài báo!\n\n` +
    `✅ PASS (Đúng & Đủ): ${successCount}\n` +
    `⚠️ CẢNH BÁO (Sai anchor / thiếu link): ${warningCount}\n` +
    `❌ LỖI (Không mở được link / không có link): ${errorCount}`
  );
}

function verifyArticle(articleUrl, anchorTextRaw, urlDichRaw) {
  const anchors = anchorTextRaw.split('\n').map(s => s.trim()).filter(s => s.length > 0);
  const targetUrls = urlDichRaw.split('\n').map(s => s.trim()).filter(s => s.length > 0);

  if (targetUrls.length === 0) {
    return { status: 'LỖI', details: ['❌ Không có URL đích cần kiểm tra'] };
  }

  let html = '';
  try {
    const cleanArticleUrl = cleanUrlString(articleUrl);
    const options = {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'vi,en-US;q=0.9,en;q=0.8'
      }
    };
    const response = UrlFetchApp.fetch(cleanArticleUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      return { status: 'LỖI HTTP', details: [`❌ Lỗi tải trang: HTTP ${responseCode}`] };
    }
    html = response.getContentText();
  } catch (err) {
    return { status: 'LỖI TẢI TRANG', details: [`❌ Không thể tải link: ${err.message}`] };
  }

  const pageLinks = extractLinksFromHtml(html);
  let allOk = true;
  let hasPartialError = false;
  const details = [];

  for (let i = 0; i < targetUrls.length; i++) {
    const expectedUrl = targetUrls[i];
    const expectedAnchor = anchors[i] || '';
    const normExpectedUrl = normalizeUrlForMatching(expectedUrl);

    // Khớp URL đích
    const matched = pageLinks.filter(l => l.normHref === normExpectedUrl || l.normHref.indexOf(normExpectedUrl) !== -1 || normExpectedUrl.indexOf(l.normHref) !== -1);

    if (matched.length === 0) {
      allOk = false;
      hasPartialError = true;
      details.push(`❌ THIẾU LINK: [${expectedAnchor}] -> ${expectedUrl}`);
    } else {
      let anchorMatch = false;
      const actualTexts = [];

      for (let m = 0; m < matched.length; m++) {
        const actualText = matched[m].text;
        actualTexts.push(actualText || '(trống/hình ảnh)');

        if (expectedAnchor) {
          const normExpAnchor = cleanStringForCompare(expectedAnchor);
          const normActAnchor = cleanStringForCompare(actualText);
          if (normActAnchor.indexOf(normExpAnchor) !== -1 || normExpAnchor.indexOf(normActAnchor) !== -1) {
            anchorMatch = true;
            break;
          }
        } else {
          anchorMatch = true;
          break;
        }
      }

      if (anchorMatch) {
        details.push(`✅ OK: "${expectedAnchor}" -> ${expectedUrl}`);
      } else {
        allOk = false;
        hasPartialError = true;
        details.push(`⚠️ SAI ANCHOR: Order "${expectedAnchor}" ➜ Thực tế: "${actualTexts.join(' | ')}"`);
      }
    }
  }

  let finalStatus = 'PASS';
  if (!allOk) {
    finalStatus = hasPartialError ? 'CẢNH BÁO' : 'LỖI';
  }

  return { status: finalStatus, details: details };
}

function extractLinksFromHtml(html) {
  const links = [];
  const regex = /<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = regex.exec(html)) !== null) {
    const href = match[1].trim();
    if (!href || href.startsWith('#') || href.startsWith('javascript:')) continue;

    let text = match[2].replace(/<[^>]+>/g, ' ');
    text = decodeHtmlEntities(text);
    text = text.replace(/\s+/g, ' ').trim();

    links.push({
      href: href,
      normHref: normalizeUrlForMatching(href),
      text: text
    });
  }
  return links;
}

function normalizeUrlForMatching(url) {
  if (!url) return '';
  let u = url.trim().toLowerCase();
  u = u.split('?')[0].split('#')[0];
  u = u.replace(/^https?:\/\//, '');
  u = u.replace(/^www\./, '');
  u = u.replace(/\/+$/, '');
  return u;
}

function cleanStringForCompare(str) {
  if (!str) return '';
  // Normalize Unicode (NFC), remove dots, replace hyphens/commas with space, lowercase, and collapse spaces
  let s = str.normalize('NFC').toLowerCase();
  s = s.replace(/\./g, '').replace(/[-–—,]/g, ' ');
  return s.replace(/\s+/g, ' ').trim();
}

function decodeHtmlEntities(text) {
  if (!text) return '';
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec));
}

function cleanUrlString(url) {
  if (!url) return '';
  const match = url.match(/https?:\/\/[^\s"'<>]+/);
  return match ? match[0] : url.trim();
}

function ensureHeaders(sheet) {
  const headers = ['Trạng thái', 'Chi tiết nghiệm thu', 'Thời gian check'];
  const headerRange = sheet.getRange(1, CONFIG.colStatus, 1, 3);
  const currentValues = headerRange.getValues()[0];

  if (!currentValues[0] || !currentValues[1] || !currentValues[2]) {
    headerRange.setValues([headers]);
    headerRange.setBackground('#1E40AF').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('CENTER');
    sheet.setColumnWidth(CONFIG.colStatus, 140);
    sheet.setColumnWidth(CONFIG.colDetail, 420);
    sheet.setColumnWidth(CONFIG.colTime, 160);
  }
}
