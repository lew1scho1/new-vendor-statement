/**
 * @OnlyCurrentDoc
 */

const UPS_TRACKING_CONFIG = {
  INPUT_SHEET: 'INPUT',
  TRACKING_SHEET: 'TRACKING',
  INPUT_VENDOR_COL: 1,  // A
  INPUT_UPS_LINK_COL: 14, // N
  TRACKING_HEADERS: ['Tracking #', 'Vendor', 'Input Row', 'Status Text', 'Delivery Date', 'Keep?', 'Paste Raw'],
  SPECIAL_ROWS: [501, 505],
  START_ROW: 507
};

/**
 * Backward-compatible entry point.
 * Use this for assigned buttons if needed.
 */
function updateUPSTracking() {
  refreshUPSTrackingList();
}

/**
 * Rebuilds TRACKING list from INPUT!N (rows 501, 505, and 507+).
 * No external API call is made.
 */
function refreshUPSTrackingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(UPS_TRACKING_CONFIG.INPUT_SHEET);
  const trackingSheet = getOrCreateTrackingSheet_(ss);
  const ui = SpreadsheetApp.getUi();

  if (!inputSheet) {
    ui.alert('Sheet "INPUT" not found.');
    return;
  }

  const existingMap = buildExistingTrackingMap_(trackingSheet);
  const lastRow = inputSheet.getLastRow();
  const candidateRows = buildCandidateRows_(lastRow);
  const nextRows = [];
  const seen = new Set();

  candidateRows.forEach(rowNum => {
    const vendor = String(inputSheet.getRange(rowNum, UPS_TRACKING_CONFIG.INPUT_VENDOR_COL).getValue() || '').trim();
    const rawLink = String(inputSheet.getRange(rowNum, UPS_TRACKING_CONFIG.INPUT_UPS_LINK_COL).getValue() || '').trim();
    if (!rawLink) return;

    const trackingNums = extractTrackingNumbers_(rawLink);
    if (trackingNums.length === 0) return;

    trackingNums.forEach(trackingNum => {
      const key = `${trackingNum}__${rowNum}`;
      if (seen.has(key)) return;
      seen.add(key);

      const existing = existingMap[key];
      const statusText = existing ? existing.statusText : '';
      const deliveryDate = existing ? existing.deliveryDate : '';
      const pasteRaw = existing ? existing.pasteRaw : '';
      const keep = computeKeepFlag_(statusText);

      nextRows.push([
        trackingNum,
        vendor,
        rowNum,
        statusText,
        deliveryDate,
        keep,
        pasteRaw
      ]);
    });
  });

  writeTrackingRows_(trackingSheet, nextRows);
  const trackingNumbers = nextRows.map(r => String(r[0] || '')).filter(Boolean);
  if (trackingNumbers.length > 0) {
    showUPSOpenAndCopyDialog_(trackingNumbers);
  }
  ss.toast(`TRACKING list refreshed: ${nextRows.length} item(s).`, 'UPS Tracking', 5);
}

/**
 * Parses user-pasted UPS full-page text in TRACKING!G and updates D/E/F.
 */
function parsePastedUPSResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackingSheet = getOrCreateTrackingSheet_(ss);
  const lastRow = trackingSheet.getLastRow();

  if (lastRow < 2) {
    ss.toast('No tracking rows to parse.', 'UPS Tracking', 5);
    return;
  }

  const rowCount = lastRow - 1;
  const values = trackingSheet.getRange(2, 1, rowCount, 7).getValues();
  const updates = [];
  let updated = 0;
  const globalRaw = values
    .map(row => String(row[6] || '').trim())
    .filter(Boolean)
    .join('\n');

  values.forEach(row => {
    const trackingNum = String(row[0] || '').trim();

    // Ignore noise rows created by multi-line paste (no tracking number in A).
    if (!trackingNum) {
      updates.push(['', '', '']);
      return;
    }

    const sourceRaw = globalRaw;
    if (!sourceRaw) {
      updates.push([row[3] || '', row[4] || '', computeKeepFlag_(row[3] || '')]);
      return;
    }

    const scopedRaw = sliceTrackingBlockByNumber_(sourceRaw, trackingNum);
    const parsedStatus = extractStatusText_(scopedRaw) || extractStatusText_(sourceRaw);
    const parsedDate = extractDeliveryDateText_(scopedRaw) || extractDeliveryDateText_(sourceRaw);
    const keep = computeKeepFlag_(parsedStatus);

    updates.push([parsedStatus, parsedDate, keep]);
    updated += 1;
  });

  trackingSheet.getRange(2, 4, updates.length, 3).setValues(updates);
  cleanupNoiseRows_(trackingSheet);
  ss.toast(`Parsed pasted results for ${updated} row(s).`, 'UPS Tracking', 5);
}

/**
 * Legacy menu compatibility.
 */
function updateTrackingInfo() {
  refreshUPSTrackingList();
}

/**
 * Legacy menu compatibility.
 */
function sendTrackingEmail() {
  SpreadsheetApp.getUi().alert(
    'UPS Tracking',
    'Email summary for tracking is deprecated in manual UPS mode. Use TRACKING sheet directly.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Legacy menu compatibility.
 */
function setupAfterShipCredentials() {
  SpreadsheetApp.getUi().alert(
    'UPS Tracking',
    'AfterShip/API setup is disabled. This file now works with manual paste parsing only.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Legacy menu compatibility.
 */
function setupDailyTrigger() {
  SpreadsheetApp.getUi().alert(
    'UPS Tracking',
    'Auto API tracking trigger is disabled. Use refreshUPSTrackingList() and parsePastedUPSResults() manually.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Legacy menu compatibility.
 */
function setupDailyTrackingWithEmail() {
  SpreadsheetApp.getUi().alert(
    'UPS Tracking',
    'Auto tracking + email mode is disabled in manual UPS mode.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getOrCreateTrackingSheet_(ss) {
  let sheet = ss.getSheetByName(UPS_TRACKING_CONFIG.TRACKING_SHEET);
  if (!sheet) sheet = ss.insertSheet(UPS_TRACKING_CONFIG.TRACKING_SHEET);
  ensureTrackingHeader_(sheet);
  return sheet;
}

function ensureTrackingHeader_(sheet) {
  sheet.getRange(1, 1, 1, UPS_TRACKING_CONFIG.TRACKING_HEADERS.length)
    .setValues([UPS_TRACKING_CONFIG.TRACKING_HEADERS])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 240);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 430);
}

function buildCandidateRows_(lastRow) {
  const rows = UPS_TRACKING_CONFIG.SPECIAL_ROWS.filter(r => r <= lastRow);
  for (let r = UPS_TRACKING_CONFIG.START_ROW; r <= lastRow; r++) rows.push(r);
  return rows;
}

function buildExistingTrackingMap_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  values.forEach(row => {
    const trackingNum = String(row[0] || '').trim();
    const inputRow = Number(row[2] || 0);
    if (!trackingNum || !inputRow) return;
    const key = `${trackingNum}__${inputRow}`;
    map[key] = {
      statusText: row[3] || '',
      deliveryDate: row[4] || '',
      pasteRaw: row[6] || ''
    };
  });

  return map;
}

function writeTrackingRows_(sheet, rows) {
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, 7).clearContent();
  if (rows.length === 0) return;
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);
}

function extractTrackingNumbers_(raw) {
  const text = String(raw || '').trim();
  if (!text) return [];

  const out = new Set();
  const oneZAll = text.match(/\b1Z[0-9A-Z]{16}\b/gi) || [];
  oneZAll.forEach(v => out.add(v.toUpperCase()));

  const paramsRegex = /(?:tracknum|trakNumber|InquiryNumber\d+)=([A-Za-z0-9]+)/gi;
  let m;
  while ((m = paramsRegex.exec(text)) !== null) {
    const candidate = String(m[1] || '').toUpperCase();
    if (/^1Z[0-9A-Z]{16}$/.test(candidate)) out.add(candidate);
  }

  return Array.from(out);
}

function extractStatusText_(rawText) {
  const lines = String(rawText || '')
    .split(/\r?\n/)
    .map(s => s.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  if (lines.length === 0) return '';

  const statusKeywords = [
    'On the Way',
    'In Transit',
    'Out for Delivery',
    'Delivered',
    'Exception',
    'Delay',
    'Pending',
    'Label Created'
  ];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (let j = 0; j < statusKeywords.length; j++) {
      if (line.toLowerCase().includes(statusKeywords[j].toLowerCase())) return line;
    }
  }

  return lines[0].substring(0, 220);
}

function extractDeliveryDateText_(rawText) {
  const lines = String(rawText || '')
    .split(/\r?\n/)
    .map(s => s.replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!/(estimated delivery|scheduled delivery|expected delivery|delivery date|arriving|arrives)/i.test(line)) {
      continue;
    }
    if (line.includes(':')) {
      const value = line.split(':').slice(1).join(':').trim();
      if (value) return value.substring(0, 120);
    }
    const nextLine = lines[i + 1] || '';
    if (nextLine) return nextLine.substring(0, 120);
  }

  const text = String(rawText || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';

  const phraseRegexes = [
    /(?:Estimated|Scheduled|Expected)\s+delivery\s*:\s*([^.;\n\r]{1,120})/i,
    /((?:Today|Tomorrow),\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:,\s*\d{4})?(?:\s+by\s+\d{1,2}:\d{2}\s*[AP]\.?M\.?)?)/i,
    /((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:,\s*\d{4})?(?:\s+by\s+\d{1,2}:\d{2}\s*[AP]\.?M\.?)?)/i,
    /(\d{1,2}\/\d{1,2}\/\d{2,4})/
  ];

  for (let i = 0; i < phraseRegexes.length; i++) {
    const m = text.match(phraseRegexes[i]);
    if (m && m[1]) return m[1].trim();
  }

  return '';
}

function sliceTrackingBlockByNumber_(rawText, trackingNum) {
  const raw = String(rawText || '');
  const num = String(trackingNum || '').trim().toUpperCase();
  if (!raw || !num) return raw;

  const upperRaw = raw.toUpperCase();
  const startIdx = upperRaw.indexOf(num);
  if (startIdx < 0) return raw;

  const afterStart = upperRaw.substring(startIdx + num.length);
  const nextTracking = afterStart.match(/\b1Z[0-9A-Z]{16}\b/);
  if (!nextTracking) return raw.substring(startIdx);

  const endIdx = startIdx + num.length + nextTracking.index;
  return raw.substring(startIdx, endIdx);
}

function computeKeepFlag_(statusText) {
  const s = String(statusText || '').toLowerCase();
  if (s.includes('delivered')) return 'N';
  return 'Y';
}

function cleanupNoiseRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const rowsToDelete = [];

  values.forEach((row, idx) => {
    const trackingNum = String(row[0] || '').trim();
    const vendor = String(row[1] || '').trim();
    const inputRow = String(row[2] || '').trim();
    if (!trackingNum && !vendor && !inputRow) rowsToDelete.push(idx + 2);
  });

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

function buildUPSUrl_(trackingNumbers) {
  return 'https://www.ups.com/track?loc=en_US&requester=ST/';
}

function showUPSOpenAndCopyDialog_(trackingNumbers) {
  const template = HtmlService.createTemplateFromFile('UPS_OpenAndCopy');
  template.upsUrl = buildUPSUrl_(trackingNumbers);
  template.clipboardText = trackingNumbers.join('\n');
  template.totalCount = trackingNumbers.length;

  const html = template
    .evaluate()
    .setWidth(560)
    .setHeight(420);

  SpreadsheetApp.getUi().showModelessDialog(html, 'UPS Open + Clipboard');
}
