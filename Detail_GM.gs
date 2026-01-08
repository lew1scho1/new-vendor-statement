/**
 * @OnlyCurrentDoc
 *
 * Detail_GM.gs
 */

/**
 * Create and populate GM detail sheets (AMKO/GNS/JINNY/SHK).
 * Distributes GM vendors from BASIC into tabs, 4 per sheet.
 * @param {object} [preReadVendorData] - Optional pre-read vendor data (Phase 1 optimization)
 */
function createAndPopulateGmDetailSheets(preReadVendorData) {
  Logger.log('\n========== Auto Create and Populate GM Detail Sheets ==========');

  const gmVendors = getGmVendorsFromBasicRange();
  if (gmVendors.length === 0) {
    SpreadsheetApp.getUi().alert('BASIC sheet has no GM vendors in the AMKO..UNION range.');
    Logger.log('ERROR: No GM vendors found in AMKO..UNION range.');
    return;
  }

  const unitSize = 4;
  const groups = buildGmVendorGroups(gmVendors, unitSize);
  const targetSheetNames = groups.map(group => group.name);

  const props = PropertiesService.getDocumentProperties();
  const previous = props.getProperty('GM_DETAIL_SHEETS');
  const previousSheetNames = previous ? previous.split(',').map(s => s.trim()).filter(Boolean) : [];

  // Remove previously managed sheets that are no longer needed
  for (const oldName of previousSheetNames) {
    if (!targetSheetNames.includes(oldName)) {
      const oldSheet = getSheet(oldName);
      if (oldSheet) {
        getActiveSpreadsheet().deleteSheet(oldSheet);
      }
    }
  }

  for (const group of groups) {
    ensureSheetExists(group.name);
    prepareGmDetailSheetLayout(group.name, group.vendors, unitSize);
    populateGmDetailSheetData(group.name, group.vendors, unitSize, preReadVendorData);
  }

  props.setProperty('GM_DETAIL_SHEETS', targetSheetNames.join(','));

  Logger.log('GM detail sheets created and populated.');
  writeToLog('GM', 'GM detail sheets created and populated');
}

/**
 * Reads GM vendors from BASIC and returns the AMKO..UNION range (inclusive).
 * Uses BASIC order.
 * @return {string[]}
 */
function getGmVendorsFromBasicRange() {
  const basicSheet = getSheet(SHEET_NAMES.BASIC);
  if (!basicSheet) {
    Logger.log('ERROR: BASIC sheet not found');
    return [];
  }

  const basicData = basicSheet.getDataRange().getValues().slice(1);
  const gmVendors = [];

  for (let i = 0; i < basicData.length; i++) {
    const vendorName = normalizeVendorName(basicData[i][COLUMN_INDICES.BASIC.VENDOR - 1]);
    const category = String(basicData[i][COLUMN_INDICES.BASIC.CATEGORY - 1] || '').trim().toUpperCase();
    if (category === 'GM' && vendorName) {
      gmVendors.push(vendorName);
    }
  }

  const startIndex = gmVendors.indexOf('AMKO');
  const endIndex = gmVendors.lastIndexOf('UNION');
  if (startIndex !== -1 && endIndex !== -1 && endIndex >= startIndex) {
    return gmVendors.slice(startIndex, endIndex + 1);
  }

  return gmVendors;
}

/**
 * Distributes vendors across sheets in order, using a fixed unit size.
 * @param {string[]} vendors
 * @param {string[]} sheetNames
 * @param {number} unitSize
 * @return {Object<string, string[]>}
 */
function buildGmVendorGroups(vendors, unitSize) {
  const groups = [];
  for (let i = 0; i < vendors.length; i += unitSize) {
    const vendorsForSheet = vendors.slice(i, i + unitSize);
    let name = vendorsForSheet[0] || `GM_${Math.floor(i / unitSize) + 1}`;
    if (groups.some(group => group.name === name)) {
      name = `${name}_${Math.floor(i / unitSize) + 1}`;
    }
    groups.push({ name, vendors: vendorsForSheet });
  }
  return groups;
}

function ensureSheetExists(sheetName) {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Prepares GM detail sheet layout.
 * Row 1 and 3 are filled; row 2 is untouched.
 * @param {string} sheetName
 * @param {string[]} vendorsForSheet
 * @param {number} unitSize
 */
function prepareGmDetailSheetLayout(sheetName, vendorsForSheet, unitSize) {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_COLUMNS = UNIT_COLUMNS * unitSize;

  // Clear data area only (rows 5+). Rows 1-3 are user-managed.
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 4) {
    sheet.getRange(5, 1, maxRows - 4, maxCols).clearContent();
  }

  // Row 1: vendor names (do not touch row 2)
  for (let i = 0; i < unitSize; i++) {
    const name = vendorsForSheet[i] || '';
    const startCol = i * UNIT_COLUMNS + 1;
    sheet.getRange(1, startCol).setValue(name);
  }

  // Row 3: headers
  const headers = ['DATE', 'INVOICE', 'AMOUNT', 'PAY DATE', 'SOURCE'];
  for (let i = 0; i < unitSize; i++) {
    const startCol = i * UNIT_COLUMNS + 1;
    for (let j = 0; j < headers.length; j++) {
      sheet.getRange(3, startCol + j).setValue(headers[j]);
    }
  }

  // Row 4: year row (merged)
  const yearRange = sheet.getRange(4, 1, 1, TOTAL_COLUMNS);
  yearRange.breakApart();
  yearRange.merge();
  yearRange.setValue('2025');
  yearRange.setHorizontalAlignment('center');
  yearRange.setVerticalAlignment('middle');
  yearRange.setFontSize(18);
  yearRange.setFontWeight('bold');
  yearRange.setBackground('#7db3a6');

  applyDetailSheetBorders(sheet, TOTAL_COLUMNS, UNIT_COLUMNS);

  sheet.hideRows(1);
  sheet.setFrozenRows(4);
}
