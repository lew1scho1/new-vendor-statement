/**
 * @OnlyCurrentDoc
 *
 * Common.gs
 * 공통 상수, 설정, 유틸리티 함수
 */

// ========== 공통 상수 ==========

const SHEET_NAMES = {
  INPUT: 'INPUT',
  MONTHLY: 'MONTHLY',
  VENDOR: 'VENDOR',
  BASIC: 'BASIC',
  ETC_DETAILS: 'ETC 상세',
  LOG: 'LOG'
};

const COLUMN_INDICES = {
  INPUT: {
    VENDOR: 1,
    YEAR: 2,
    MONTH: 3,
    DAY: 4,         // D열: Day
    INVOICE: 5,     // E열: Invoice #
    AMOUNT: 6,
    PAY_YEAR: 7,    // G열: Payment Year
    PAY_MONTH: 8,   // H열: Payment Month
    PAY_DATE: 9,    // I열: Payment Date
    OUTSTANDING: 11, // K열: Outstanding (O/X/빈칸)
    CHECK_NUM: 12   // L열: Check Number
  },
  BASIC: {
    VENDOR: 1,
    CATEGORY: 2,
    PAYMENT_METHOD: 3,
    CUST_NUM: 4
  }
};

const HEADER_ROWS = {
  MONTHLY: {
    YEAR: 4,  // 1-based row index
    MONTH: 5  // 1-based row index
  },
  VENDOR: {
    YEAR: 3,  // 1-based row index
    MONTH: 4  // 1-based row index
  }
};

// VENDOR 탭 관련 상수
const VENDOR_CELLS_PER_MONTH = 8; // 한 달당 8칸 (4개 인보이스 × 2칸)
const VENDOR_MAX_INVOICES = 4;    // 최대 4개 인보이스 표시

const PROTECTED_LABELS = ['SUBTOTAL', 'MONTH', 'GM GRAND TOTAL', 'GM GRAND  TOTAL', 'GRAND TOTAL'];
const PROTECTED_COLUMNS = [15, 28]; // O, AB columns (1-based)

const SECTION_LABELS = ['HAIR', 'GM', 'HARWIN'];
const SUBTOTAL_LABEL = 'SUBTOTAL';
const MONTH_ROW_LABEL = 'MONTH';
const GM_GRAND_TOTAL_LABEL = 'GM GRAND TOTAL';
const GRAND_TOTAL_ROW = 1;

const MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

// ========== 공통 유틸리티 함수 ==========

/**
 * 월 이름을 숫자로 변환
 * @param {string} monthName - 월 이름 (JAN, FEB, etc.)
 * @return {number} 월 번호 (1-12), 유효하지 않으면 0
 */
function parseMonthNameToNumber(monthName) {
  const monthNameMap = {
    JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUNE: 6,
    JUL: 7, JULY: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
  };

  const upper = String(monthName).trim().toUpperCase();
  return monthNameMap[upper] || 0;
}

/**
 * 벤더 이름 정규화 (공백 제거, 대소문자 통일)
 * @param {string} vendorName - 원본 벤더 이름
 * @return {string} 정규화된 벤더 이름
 */
function normalizeVendorName(vendorName) {
  return String(vendorName || '').trim();
}

/**
 * 스프레드시트 가져오기
 * @return {Spreadsheet} 현재 활성 스프레드시트
 */
function getActiveSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 시트 가져오기
 * @param {string} sheetName - 시트 이름
 * @return {Sheet|null} 시트 객체 또는 null
 */
function getSheet(sheetName) {
  const ss = getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * 두 시트가 모두 존재하는지 확인
 * @param {string} sheetName1 - 첫 번째 시트 이름
 * @param {string} sheetName2 - 두 번째 시트 이름
 * @return {boolean} 두 시트 모두 존재하면 true
 */
function checkSheetsExist(sheetName1, sheetName2) {
  const sheet1 = getSheet(sheetName1);
  const sheet2 = getSheet(sheetName2);
  return sheet1 !== null && sheet2 !== null;
}

/**
 * 여러 시트가 모두 존재하는지 확인
 * @param {...string} sheetNames - 확인할 시트 이름들
 * @return {boolean} 모든 시트가 존재하면 true
 */
function checkMultipleSheetsExist(...sheetNames) {
  for (const sheetName of sheetNames) {
    if (!getSheet(sheetName)) {
      return false;
    }
  }
  return true;
}

/**
 * 날짜를 MM/DD 형식으로 포맷팅
 * @param {number} month - 월 (1-12)
 * @param {number} date - 일 (1-31)
 * @return {string} MM/DD 형식의 문자열
 */
function formatPaymentDate(month, date) {
  const mm = String(month).padStart(2, '0');
  const dd = String(date).padStart(2, '0');
  return `${mm}/${dd}`;
}

/**
 * LOG 시트에 메시지 기록
 * @param {string} source - 로그 출처 (예: 'MONTHLY', 'VENDOR')
 * @param {string} message - 로그 메시지
 */
function writeToLog(source, message) {
  // As per user request, skip logging for sources containing 'ETC' or 'HARWIN'
  // to prevent errors related to merged cells in the LOG sheet.
  if (source && (source.toUpperCase().includes('ETC') || source.toUpperCase().includes('HARWIN'))) {
    Logger.log(`[${source}] ${message} (Skipped writing to LOG sheet)`);
    return;
  }

  const ss = getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(SHEET_NAMES.LOG);

  // LOG 시트가 없으면 생성
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAMES.LOG);
    // 헤더 추가
    logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Source', 'Message']]);
    logSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
    logSheet.setFrozenRows(1);
  }

  // 새 로그 추가 (맨 아래에 추가)
  const timestamp = new Date();
  const nextRow = logSheet.getLastRow() + 1;
  breakApartMergedRangesAtRow(logSheet, nextRow);
  logSheet.getRange(nextRow, 1, 1, 3).setValues([[timestamp, source, message]]);

  // 타임스탬프 포맷 설정
  logSheet.getRange(nextRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  Logger.log(`[${source}] ${message}`);
}

/**
 * Breaks any merged ranges that include the target row.
 * @param {Sheet} sheet
 * @param {number} row
 */
function breakApartMergedRangesAtRow(sheet, row) {
  // Get a range representing the entire target row to check for intersections.
  const targetRowRange = sheet.getRange(row, 1, 1, sheet.getMaxColumns());

  // Get only the merged ranges that specifically overlap with our target row.
  const overlappingMergedRanges = targetRowRange.getMergedRanges();

  // Iterate through the intersecting merged ranges and break them apart.
  // This is safer than iterating over all merged ranges on the sheet.
  for (const mergedRange of overlappingMergedRanges) {
    mergedRange.breakApart();
  }
}

/**
 * Apply common borders for detail sheets.
 * - Medium vertical borders between units (and right edge)
 * - Thin borders around row 3 cells
 * @param {Sheet} sheet
 * @param {number} totalColumns
 * @param {number} unitColumns
 */
function applyDetailSheetBorders(sheet, totalColumns, unitColumns) {
  const maxRows = sheet.getMaxRows();

  // Row 3: thin borders around each cell
  sheet.getRange(3, 1, 1, totalColumns)
    .setBorder(true, true, true, true, true, true, SpreadsheetApp.BorderStyle.SOLID, null);

  // Vertical unit separators (including rightmost edge)
  for (let col = unitColumns; col <= totalColumns; col += unitColumns) {
    sheet.getRange(1, col, maxRows, 1)
      .setBorder(null, null, null, true, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM, null);
  }
}
