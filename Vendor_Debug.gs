/**
 * @OnlyCurrentDoc
 *
 * Vendor_Debug.gs
 * VENDOR 시트 디버깅 함수들
 */

/**
 * VENDOR 시트의 구조를 확인하는 디버그 함수
 */
function debugVendorStructure() {
  Logger.log('========== VENDOR STRUCTURE DEBUG ==========');

  const vendorSheet = getSheet(SHEET_NAMES.VENDOR);

  if (!vendorSheet) {
    Logger.log('ERROR: Could not find VENDOR sheet.');
    SpreadsheetApp.getUi().alert('Error: Could not find VENDOR sheet.');
    return;
  }

  const values = vendorSheet.getDataRange().getValues();
  Logger.log('Total rows in VENDOR sheet: ' + values.length);
  Logger.log('\nColumn A contents (first 100 rows):');
  Logger.log('Row# | Value | Trimmed | UpperCase');
  Logger.log('-----|-------|---------|----------');

  for (let i = 0; i < Math.min(values.length, 100); i++) {
    const cell = values[i][0];
    const cellStr = String(cell || '');
    const trimmed = cellStr.trim();
    const upper = trimmed.toUpperCase();

    if (trimmed) {
      Logger.log(`${i + 1} | "${cellStr}" | "${trimmed}" | "${upper}"`);
    }
  }

  Logger.log('\n========== Looking for key labels ==========');

  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0] || '').trim().toUpperCase();

    if (SECTION_LABELS.some(s => cell === s.toUpperCase())) {
      Logger.log(`Found SECTION "${values[i][0]}" at row ${i + 1}`);
    }
    if (cell.includes(SUBTOTAL_LABEL.toUpperCase())) {
      Logger.log(`Found SUBTOTAL "${values[i][0]}" at row ${i + 1}`);
    }
    if (cell === MONTH_ROW_LABEL.toUpperCase()) {
      Logger.log(`Found MONTH at row ${i + 1}`);
    }
    if (cell === GM_GRAND_TOTAL_LABEL.toUpperCase()) {
      Logger.log(`Found GM GRAND TOTAL at row ${i + 1}`);
    }
  }

  Logger.log('\n========== Checking Year/Month Headers ==========');
  if (values.length > HEADER_ROWS.VENDOR.MONTH) {
    const yearHeader = values[HEADER_ROWS.VENDOR.YEAR - 1];
    const monthHeader = values[HEADER_ROWS.VENDOR.MONTH - 1];

    Logger.log('Row 3 (Year): ' + yearHeader.slice(0, 30).join(' | '));
    Logger.log('Row 4 (Month): ' + monthHeader.slice(0, 30).join(' | '));
  }

  Logger.log('\n========== STRUCTURE DEBUG END ==========');
  SpreadsheetApp.getUi().alert('VENDOR 구조 확인 완료!\n\n자세한 내용은 보기 > 로그를 확인하세요.');
}

/**
 * VENDOR 동기화 디버그 함수
 */
function debugSyncVendorSummary() {
  Logger.log('========== VENDOR DEBUG MODE START ==========');

  // 1. 시트 존재 확인
  if (!checkMultipleSheetsExist(SHEET_NAMES.INPUT, SHEET_NAMES.VENDOR, SHEET_NAMES.BASIC)) {
    Logger.log('ERROR: Could not find INPUT, VENDOR, or BASIC sheet.');
    SpreadsheetApp.getUi().alert('Error: INPUT, VENDOR, BASIC 시트를 찾을 수 없습니다.');
    return;
  }

  const vendorSheet = getSheet(SHEET_NAMES.VENDOR);

  // 2. BASIC 시트에서 payment method 읽기
  const paymentMethodMap = readBasicPaymentMethods();
  Logger.log(`\nTotal payment methods loaded: ${Object.keys(paymentMethodMap).length}`);

  // 3. INPUT 시트에서 인보이스 데이터 읽기
  const invoicesData = readVendorInvoicesFromInput(paymentMethodMap);

  // 4. VENDOR 시트 분석
  const vendorRange = vendorSheet.getDataRange();
  const vendorValues = vendorRange.getValues();

  // Parse year/month column headers
  const yearMonthCols = parseVendorYearMonthColumns(vendorValues);
  if (!yearMonthCols) return;

  // Analyze sheet structure
  const structure = analyzeVendorSheetStructure(vendorValues);

  Logger.log('\n========== STEP 3: Vendor Matching ==========');
  const vendorSheetVendors = new Set(Object.keys(structure.vendors));
  const inputVendors = new Set(Object.keys(invoicesData));

  Logger.log('Vendors in VENDOR sheet: ' + vendorSheetVendors.size);
  Logger.log('VENDOR Sheet Vendors:');
  for (const vendor in structure.vendors) {
    const isProtected = isProtectedLabel(vendor);
    const marker = isProtected ? ' ⚠️ PROTECTED - SHOULD NOT BE HERE!' : '';
    Logger.log(`  "${vendor}" -> Row ${structure.vendors[vendor].row + 1} (${structure.vendors[vendor].section})${marker}`);
  }

  Logger.log('\nVendors in INPUT: ' + inputVendors.size);
  Logger.log('INPUT Vendors: ' + [...inputVendors].join(', '));

  const newVendors = [...inputVendors].filter(v => !vendorSheetVendors.has(v));
  if (newVendors.length > 0) {
    Logger.log('\n⚠️  WARNING: New vendors in INPUT not found in VENDOR:');
    Logger.log('  ' + newVendors.join(', '));
  }

  // 5. 데이터 매칭 분석
  Logger.log('\n========== STEP 4: Invoice Data Analysis ==========');
  let matchCount = 0;
  let mismatchCount = 0;

  for (const vendorName in invoicesData) {
    if (!structure.vendors[vendorName]) {
      mismatchCount++;
      Logger.log(`\n❌ Vendor "${vendorName}" from INPUT NOT FOUND in VENDOR`);

      // Check for similar names
      const similar = [...vendorSheetVendors].filter(v =>
        v.toLowerCase().includes(vendorName.toLowerCase()) ||
        vendorName.toLowerCase().includes(v.toLowerCase())
      );
      if (similar.length > 0) {
        Logger.log(`   Possible matches: ${similar.join(', ')}`);
      }
      continue;
    }

    matchCount++;
    const vendorRow = structure.vendors[vendorName].row;
    Logger.log(`\n✅ Vendor "${vendorName}" matched -> Row ${vendorRow + 1}`);

    for (const year in invoicesData[vendorName]) {
      const colsForYear = yearMonthCols[year];
      if (!colsForYear) {
        Logger.log(`   ❌ Year ${year} NOT FOUND in VENDOR header`);
        continue;
      }

      Logger.log(`   Year ${year}:`);
      for (const month in invoicesData[vendorName][year]) {
        const startCol = colsForYear[month];
        if (!startCol) {
          Logger.log(`     ❌ Month ${month} NOT FOUND in VENDOR header`);
          continue;
        }

        const invoices = invoicesData[vendorName][year][month];
        const limitedInvoices = limitAndMergeInvoices(invoices, vendorName);
        Logger.log(`     ✅ Month ${month}: ${invoices.length} invoices (showing ${limitedInvoices.length}) -> Starting at Col ${startCol}`);

        for (let i = 0; i < limitedInvoices.length; i++) {
          const inv = limitedInvoices[i];
          const dateStr = formatPaymentDate(inv.payMonth, inv.payDate);
          const methodStr = formatPaymentMethod(inv);
          Logger.log(`        Invoice ${i + 1}: $${inv.amount} ${dateStr}${methodStr}`);
        }
      }
    }
  }

  // 6. 요약 출력
  Logger.log('\n========== SUMMARY ==========');
  Logger.log('Total vendors in INPUT: ' + inputVendors.size);
  Logger.log('Matched vendors: ' + matchCount);
  Logger.log('Unmatched vendors: ' + mismatchCount);
  Logger.log('========== VENDOR DEBUG MODE END ==========');

  SpreadsheetApp.getUi().alert(
    'DEBUG 완료!\n\n' +
    'INPUT 벤더: ' + inputVendors.size + '개\n' +
    'VENDOR 벤더: ' + vendorSheetVendors.size + '개\n' +
    '매칭된 벤더: ' + matchCount + '개\n' +
    '매칭 안된 벤더: ' + mismatchCount + '개\n\n' +
    '자세한 내용은 보기 > 로그를 확인하세요.'
  );
}

/**
 * ETC 벤더의 Outstanding 상태를 디버깅하는 함수
 */
function debugEtcOutstanding() {
  Logger.log('========== ETC OUTSTANDING DEBUG START ==========');

  // 1. 시트 확인
  if (!checkMultipleSheetsExist(SHEET_NAMES.INPUT, SHEET_NAMES.VENDOR, SHEET_NAMES.BASIC)) {
    SpreadsheetApp.getUi().alert('Error: INPUT, VENDOR, BASIC 시트를 찾을 수 없습니다.');
    return;
  }

  const vendorSheet = getSheet(SHEET_NAMES.VENDOR);

  // 2. BASIC 시트에서 payment method 읽기
  const paymentMethodMap = readBasicPaymentMethods();

  // 3. ETC 벤더 목록 읽기
  const etcVendors = getEtcVendorsFromDetailsSheet();
  Logger.log(`\nETC 상세 시트에서 읽은 벤더 수: ${etcVendors.size}`);
  Logger.log(`ETC 벤더 목록: ${[...etcVendors].join(', ')}`);

  // 4. INPUT 시트에서 인보이스 데이터 읽기
  const invoicesData = readVendorInvoicesFromInput(paymentMethodMap);

  // 5. VENDOR 시트 분석
  const vendorRange = vendorSheet.getDataRange();
  const vendorValues = vendorRange.getValues();
  const yearMonthCols = parseVendorYearMonthColumns(vendorValues);
  if (!yearMonthCols) return;

  const structure = analyzeVendorSheetStructure(vendorValues);

  // 6. ETC 벤더가 있는지 확인
  if (!structure.vendors['ETC']) {
    Logger.log('\n❌ ERROR: ETC 벤더를 VENDOR 시트에서 찾을 수 없습니다!');
    SpreadsheetApp.getUi().alert('Error: ETC 벤더를 VENDOR 시트에서 찾을 수 없습니다.');
    return;
  }

  Logger.log(`\n✅ ETC 벤더 발견: Row ${structure.vendors['ETC'].row + 1}`);

  // 7. ETC 벤더별 인보이스 상세 분석
  Logger.log('\n========== ETC 벤더별 인보이스 분석 ==========');

  let totalOutstandingCount = 0;
  let totalInvoiceCount = 0;

  for (const etcVendorName of etcVendors) {
    if (!invoicesData[etcVendorName]) {
      Logger.log(`\n⚠️ "${etcVendorName}": INPUT에 데이터 없음`);
      continue;
    }

    Logger.log(`\n📋 "${etcVendorName}":`);

    for (const year in invoicesData[etcVendorName]) {
      for (const month in invoicesData[etcVendorName][year]) {
        const invoices = invoicesData[etcVendorName][year][month];
        Logger.log(`  ${year}-${month}: ${invoices.length}개 인보이스`);

        for (let i = 0; i < invoices.length; i++) {
          const inv = invoices[i];
          totalInvoiceCount++;

          const outstandingMark = inv.isOutstanding ? '🔵 OUTSTANDING' : '🟢 PAID';
          if (inv.isOutstanding) totalOutstandingCount++;

          Logger.log(`    [${i + 1}] $${inv.amount} | ${inv.payYear}-${inv.payMonth}-${inv.payDate} | ${inv.paymentMethod} | ${outstandingMark}`);
        }
      }
    }
  }

  // 8. ETC 합산 후 분석
  Logger.log('\n========== ETC 합산 후 분석 ==========');

  // ETC 데이터 합산 (메인 로직과 동일)
  const etcAggregated = {};

  for (const etcVendorName of etcVendors) {
    if (invoicesData[etcVendorName]) {
      for (const year in invoicesData[etcVendorName]) {
        if (!etcAggregated[year]) etcAggregated[year] = {};

        for (const month in invoicesData[etcVendorName][year]) {
          if (!etcAggregated[year][month]) etcAggregated[year][month] = [];

          etcAggregated[year][month].push(...invoicesData[etcVendorName][year][month]);
        }
      }
    }
  }

  Logger.log('\nETC 합산 결과:');

  let totalOutstandingAfterMerge = 0;
  let totalInvoicesAfterMerge = 0;

  for (const year in etcAggregated) {
    for (const month in etcAggregated[year]) {
      const invoices = etcAggregated[year][month];
      const limited = limitAndMergeInvoices(invoices, 'ETC');

      Logger.log(`\n  ${year}-${month}: ${invoices.length}개 -> 병합 후 ${limited.length}개`);

      for (let i = 0; i < limited.length; i++) {
        const inv = limited[i];
        totalInvoicesAfterMerge++;

        const outstandingMark = inv.isOutstanding ? '🔵 OUTSTANDING' : '🟢 PAID';
        if (inv.isOutstanding) totalOutstandingAfterMerge++;

        Logger.log(`    [${i + 1}] $${inv.amount} | ${inv.payYear}-${inv.payMonth}-${inv.payDate} | ${outstandingMark}`);
      }
    }
  }

  // 9. 요약
  Logger.log('\n========== 요약 ==========');
  Logger.log(`ETC 상세 벤더 수: ${etcVendors.size}`);
  Logger.log(`병합 전 총 인보이스 수: ${totalInvoiceCount}`);
  Logger.log(`병합 전 Outstanding 수: ${totalOutstandingCount}`);
  Logger.log(`병합 후 총 인보이스 수: ${totalInvoicesAfterMerge}`);
  Logger.log(`병합 후 Outstanding 수: ${totalOutstandingAfterMerge}`);
  Logger.log('========== ETC OUTSTANDING DEBUG END ==========');

  SpreadsheetApp.getUi().alert(
    'ETC Outstanding 디버그 완료!\n\n' +
    `ETC 벤더: ${etcVendors.size}개\n` +
    `병합 전 인보이스: ${totalInvoiceCount}개\n` +
    `병합 전 Outstanding: ${totalOutstandingCount}개\n` +
    `병합 후 인보이스: ${totalInvoicesAfterMerge}개\n` +
    `병합 후 Outstanding: ${totalOutstandingAfterMerge}개\n\n` +
    '자세한 내용은 보기 > 로그를 확인하세요.'
  );
}

/**
 * VENDOR 시트의 특정 셀 배경색 확인
 */
function debugCellBackgrounds() {
  Logger.log('========== CELL BACKGROUND DEBUG ==========');

  const vendorSheet = getSheet(SHEET_NAMES.VENDOR);
  if (!vendorSheet) {
    Logger.log('ERROR: Could not find VENDOR sheet.');
    SpreadsheetApp.getUi().alert('Error: VENDOR 시트를 찾을 수 없습니다.');
    return;
  }

  // 확인할 셀들
  const cellsToCheck = ['A6', 'A12', 'A32'];

  Logger.log('\n셀 배경색 확인:');

  for (const cell of cellsToCheck) {
    const range = vendorSheet.getRange(cell);
    const background = range.getBackground();
    const value = range.getValue();
    const fontColor = range.getFontColor();
    const fontSize = range.getFontSize();
    const isHidden = vendorSheet.isRowHiddenByUser(range.getRow());

    Logger.log(`\n${cell}:`);
    Logger.log(`  값: "${value}"`);
    Logger.log(`  배경색: ${background}`);
    Logger.log(`  폰트색: ${fontColor}`);
    Logger.log(`  폰트크기: ${fontSize}`);
    Logger.log(`  행 숨김: ${isHidden}`);
  }

  // 추가로 모든 행의 A열 셀 배경색 확인
  Logger.log('\n\n========== 모든 A열 셀 배경색 (처음 50행) ==========');
  const values = vendorSheet.getDataRange().getValues();

  for (let i = 0; i < Math.min(50, values.length); i++) {
    const cell = vendorSheet.getRange(i + 1, 1);
    const background = cell.getBackground();
    const value = String(values[i][0] || '').trim();
    const isHidden = vendorSheet.isRowHiddenByUser(i + 1);

    if (value) {
      Logger.log(`Row ${i + 1}: "${value}" | BG: ${background} | Hidden: ${isHidden}`);
    }
  }

  Logger.log('\n========== CELL BACKGROUND DEBUG END ==========');
  SpreadsheetApp.getUi().alert('배경색 확인 완료!\n\n자세한 내용은 보기 > 로그를 확인하세요.');
}
