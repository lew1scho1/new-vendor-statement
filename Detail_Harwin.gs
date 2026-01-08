/**
 * @OnlyCurrentDoc
 *
 * Detail_Harwin.gs
 */

/**
 * Create and populate HARWIN detail sheets (HARWIN-1, HARWIN-2, etc.).
 * Distributes HARWIN vendors from BASIC into tabs, 4 per sheet.
 * HARWIN vendors are those after UNION in BASIC sheet.
 */
function createAndPopulateHarwinDetailSheets() {
  Logger.log('\n========== Auto Create and Populate HARWIN Detail Sheets ==========');

  const harwinVendors = getHarwinVendorsFromBasicAfterUnion();
  if (harwinVendors.length === 0) {
    SpreadsheetApp.getUi().alert('BASIC sheet has no HARWIN vendors after UNION.');
    Logger.log('ERROR: No HARWIN vendors found after UNION.');
    return;
  }

  Logger.log(`Total HARWIN vendors from BASIC: ${harwinVendors.length}`);
  Logger.log(`HARWIN vendors: ${harwinVendors.join(', ')}`);

  // Map ETC vendors
  const etcVendors = getEtcVendorsFromDetailsSheet();
  const mappedVendors = mapHarwinVendorsWithEtc(harwinVendors, etcVendors);
  const etcVendorsInHarwin = new Set(harwinVendors.filter(v => etcVendors.has(v)));

  Logger.log(`After ETC mapping: ${mappedVendors.length} vendors`);
  Logger.log(`Mapped vendors: ${mappedVendors.join(', ')}`);

  const unitSize = 4;
  const groups = buildHarwinVendorGroups(mappedVendors, unitSize);
  const targetSheetNames = groups.map(group => group.name);

  Logger.log(`Creating ${groups.length} HARWIN sheets: ${targetSheetNames.join(', ')}`);

  const props = PropertiesService.getDocumentProperties();
  const previous = props.getProperty('HARWIN_DETAIL_SHEETS');
  const previousSheetNames = previous ? previous.split(',').map(s => s.trim()).filter(Boolean) : [];

  // Remove previously managed sheets that are no longer needed
  for (const oldName of previousSheetNames) {
    if (!targetSheetNames.includes(oldName)) {
      const oldSheet = getSheet(oldName);
      if (oldSheet) {
        getActiveSpreadsheet().deleteSheet(oldSheet);
        Logger.log(`Deleted old sheet: ${oldName}`);
      }
    }
  }

  for (const group of groups) {
    Logger.log(`\nProcessing sheet: ${group.name}`);
    Logger.log(`  Vendors: ${group.vendors.join(', ')}`);

    ensureSheetExists(group.name);
    prepareHarwinDetailSheetLayout(group.name, group.vendors, unitSize);
    populateHarwinDetailSheetData(group.name, group.vendors, harwinVendors, etcVendorsInHarwin, unitSize);
  }

  props.setProperty('HARWIN_DETAIL_SHEETS', targetSheetNames.join(','));

  Logger.log('\nHARWIN detail sheets created and populated.');
  writeToLog('HARWIN', 'HARWIN detail sheets created and populated');
}

/**
 * Reads vendor names after UNION (skip the blank row) from BASIC.
 * @return {string[]}
 */
function getHarwinVendorsFromBasicAfterUnion() {
  const basicSheet = getSheet(SHEET_NAMES.BASIC);
  if (!basicSheet) {
    Logger.log('ERROR: BASIC sheet not found');
    return [];
  }

  const basicData = basicSheet.getDataRange().getValues().slice(1);
  const vendorNames = basicData.map(row => normalizeVendorName(row[COLUMN_INDICES.BASIC.VENDOR - 1]));

  const unionIndex = vendorNames.indexOf('UNION');
  if (unionIndex === -1) {
    Logger.log('ERROR: UNION not found in BASIC sheet');
    return [];
  }

  // Skip UNION and the blank row after it
  let startIndex = unionIndex + 1;
  for (let i = unionIndex + 1; i < vendorNames.length; i++) {
    if (!vendorNames[i]) {
      startIndex = i + 1;
      break;
    }
  }

  const harwinVendors = [];
  for (let i = startIndex; i < vendorNames.length; i++) {
    if (vendorNames[i]) {
      harwinVendors.push(vendorNames[i]);
    }
  }

  return harwinVendors;
}

/**
 * Replace vendors found in ETC list with a single "ETC" entry.
 * ETC is always placed at the end, not alphabetically.
 * @param {string[]} vendors
 * @param {Set<string>} etcVendors
 * @return {string[]}
 */
function mapHarwinVendorsWithEtc(vendors, etcVendors) {
  const result = [];
  let hasEtc = false;

  for (const name of vendors) {
    if (etcVendors.has(name)) {
      hasEtc = true;
    } else if (name) {
      result.push(name);
    }
  }

  // Add ETC at the end if any ETC vendors were found
  if (hasEtc) {
    result.push('ETC');
  }

  return result;
}

/**
 * Distributes vendors across sheets in order, using a fixed unit size.
 * Creates HARWIN-1, HARWIN-2, etc.
 * @param {string[]} vendors
 * @param {number} unitSize
 * @return {Array<{name: string, vendors: string[]}>}
 */
function buildHarwinVendorGroups(vendors, unitSize) {
  const groups = [];
  for (let i = 0; i < vendors.length; i += unitSize) {
    const vendorsForSheet = vendors.slice(i, i + unitSize);
    const sheetNumber = Math.floor(i / unitSize) + 1;
    const name = `HARWIN-${sheetNumber}`;
    groups.push({ name, vendors: vendorsForSheet });
  }
  return groups;
}

/**
 * Ensure sheet exists, create if not.
 * @param {string} sheetName
 * @return {Sheet}
 */
function ensureSheetExists(sheetName) {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName}`);
  }
  return sheet;
}

/**
 * Prepare HARWIN detail sheet layout.
 * Row 1 and 3 are filled; row 2 is untouched (user data).
 * @param {string} sheetName
 * @param {string[]} vendorsForSheet - Display names (with ETC)
 * @param {number} unitSize
 */
function prepareHarwinDetailSheetLayout(sheetName, vendorsForSheet, unitSize) {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_COLUMNS = UNIT_COLUMNS * unitSize;

  // Clear data area only (rows 5+). Rows 1-4 will be managed by this function.
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 4) {
    sheet.getRange(5, 1, maxRows - 4, maxCols).clearContent();
  }

  // Row 1: vendor names (do not touch row 2)
  // Fill entire row to clear old data if vendor count decreased
  const row1Values = new Array(maxCols).fill('');
  for (let i = 0; i < vendorsForSheet.length && i < unitSize; i++) {
    row1Values[i * UNIT_COLUMNS] = vendorsForSheet[i] || '';
  }
  sheet.getRange(1, 1, 1, maxCols).setValues([row1Values]);

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

  Logger.log(`  Layout prepared for ${sheetName}`);
}

/**
 * HARWIN 시트의 병합 상태와 레이아웃을 디버깅하는 함수
 */
function debugHarwinSheet() {
  Logger.log('\n========== HARWIN SHEET DEBUG START ==========');

  const ss = getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const sheetNames = props.getProperty('HARWIN_DETAIL_SHEETS');
  const sheets = sheetNames ? sheetNames.split(',').map(s => s.trim()).filter(Boolean) : [];

  if (sheets.length === 0) {
    Logger.log('No HARWIN sheets found in properties');
    SpreadsheetApp.getUi().alert('Error: HARWIN 시트를 찾을 수 없습니다.');
    return;
  }

  Logger.log(`Found ${sheets.length} HARWIN sheets: ${sheets.join(', ')}`);

  for (const sheetName of sheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`\n❌ Sheet not found: ${sheetName}`);
      continue;
    }

    Logger.log(`\n========== ${sheetName} ==========`);
    Logger.log(`최대 행 수: ${sheet.getMaxRows()}`);
    Logger.log(`최대 열 수: ${sheet.getMaxColumns()}`);
    Logger.log(`고정된 행: ${sheet.getFrozenRows()}`);

    // Row 1: 벤더 이름
    const row1Values = sheet.getRange(1, 1, 1, 20).getValues()[0];
    const vendors = [];
    for (let i = 0; i < 20; i += 5) {
      if (row1Values[i]) vendors.push(row1Values[i]);
    }
    Logger.log(`벤더: ${vendors.join(', ')}`);

    // Row 4: 병합 상태
    const row4Range = sheet.getRange(4, 1, 1, 20);
    const mergedRanges4 = row4Range.getMergedRanges();
    Logger.log(`Row 4 병합 수: ${mergedRanges4.length}`);
    if (mergedRanges4.length > 0) {
      mergedRanges4.forEach(range => {
        Logger.log(`  ${range.getA1Notation()}: "${range.getValue()}"`);
      });
    }
  }

  Logger.log('\n========== HARWIN SHEET DEBUG END ==========');
  SpreadsheetApp.getUi().alert('HARWIN 시트 디버그 완료!\n\n자세한 내용은 보기 > 로그를 확인하세요.');
}

/**
 * HARWIN 벤더 목록 확인
 */
function debugHarwinVendors() {
  Logger.log('\n========== HARWIN VENDORS DEBUG ==========');

  // 1. BASIC 시트에서 HARWIN 벤더 읽기
  Logger.log('\n1. BASIC 시트에서 HARWIN 벤더 읽기:');
  const harwinVendors = getHarwinVendorsFromBasicAfterUnion();
  Logger.log(`총 ${harwinVendors.length}개 벤더:`);
  harwinVendors.forEach((v, idx) => Logger.log(`  [${idx + 1}] ${v}`));

  // 2. ETC 벤더 읽기
  Logger.log('\n2. ETC 벤더 읽기:');
  const etcVendors = getEtcVendorsFromDetailsSheet();
  Logger.log(`총 ${etcVendors.size}개 ETC 벤더:`);
  [...etcVendors].forEach(v => Logger.log(`  ${v}`));

  // 3. 매핑된 벤더 (ETC로 치환)
  Logger.log('\n3. 매핑된 벤더 (ETC로 치환):');
  const mappedVendors = mapHarwinVendorsWithEtc(harwinVendors, etcVendors);
  Logger.log(`총 ${mappedVendors.length}개 벤더:`);
  mappedVendors.forEach((v, idx) => Logger.log(`  [${idx + 1}] ${v}`));

  // 4. 그룹 분리
  Logger.log('\n4. 4개씩 그룹 분리:');
  const groups = buildHarwinVendorGroups(mappedVendors, 4);
  groups.forEach(group => {
    Logger.log(`  ${group.name}: ${group.vendors.join(', ')}`);
  });

  Logger.log('\n========== HARWIN VENDORS DEBUG END ==========');

  SpreadsheetApp.getUi().alert(
    'HARWIN 벤더 디버그 완료!\n\n' +
    `BASIC의 HARWIN 벤더: ${harwinVendors.length}개\n` +
    `ETC 벤더: ${etcVendors.size}개\n` +
    `매핑 후 벤더: ${mappedVendors.length}개\n` +
    `생성될 시트: ${groups.length}개\n\n` +
    '자세한 내용은 보기 > 로그를 확인하세요.'
  );
}
