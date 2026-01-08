/**
 * @OnlyCurrentDoc
 *
 * Monthly_Debug.gs
 * MONTHLY 시트 디버깅 함수들
 */

/**
 * MONTHLY 시트의 구조를 확인하는 디버그 함수
 * Column A의 모든 값을 로그에 출력
 */
function debugMonthlyStructure() {
  Logger.log('========== MONTHLY STRUCTURE DEBUG ==========');

  const monthlySheet = getSheet(SHEET_NAMES.MONTHLY);

  if (!monthlySheet) {
    Logger.log('ERROR: Could not find MONTHLY sheet.');
    SpreadsheetApp.getUi().alert('Error: Could not find MONTHLY sheet.');
    return;
  }

  const values = monthlySheet.getDataRange().getValues();
  Logger.log('Total rows in MONTHLY sheet: ' + values.length);
  Logger.log('\nColumn A contents (first 100 rows):');
  Logger.log('Row# | Value | Trimmed | UpperCase');
  Logger.log('-----|-------|---------|----------');

  for (let i = 0; i < Math.min(values.length, 100); i++) {
    const cell = values[i][0];
    const cellStr = String(cell || '');
    const trimmed = cellStr.trim();
    const upper = trimmed.toUpperCase();

    if (trimmed) { // Only log non-empty cells
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

  Logger.log('\n========== STRUCTURE DEBUG END ==========');
  SpreadsheetApp.getUi().alert('MONTHLY 구조 확인 완료!\n\n자세한 내용은 보기 > 로그를 확인하세요.');
}

/**
 * MONTHLY 동기화 디버그 함수
 * 실제로 데이터를 수정하지 않고 각 단계의 정보를 로그로 출력
 */
function debugSyncMonthlySummary() {
  Logger.log('========== DEBUG MODE START ==========');

  // 1. 시트 존재 확인
  if (!checkSheetsExist(SHEET_NAMES.INPUT, SHEET_NAMES.MONTHLY)) {
    Logger.log('ERROR: Could not find INPUT or MONTHLY sheet.');
    SpreadsheetApp.getUi().alert('Error: Could not find INPUT or MONTHLY sheet.');
    return;
  }

  const inputSheet = getSheet(SHEET_NAMES.INPUT);
  const monthlySheet = getSheet(SHEET_NAMES.MONTHLY);

  // 2. INPUT 데이터 읽기 및 집계
  const inputData = readAndAggregateInputData();
  if (!inputData) return;

  const { summary, allInputVendors, validRowCount, invalidRowCount } = inputData;

  // 3. MONTHLY 시트 분석
  Logger.log('\n========== STEP 2: Analyzing MONTHLY Sheet ==========');
  const monthlyRange = monthlySheet.getDataRange();
  const monthlyValues = monthlyRange.getValues();

  // Parse year/month column headers
  const yearMonthCols = parseMonthlyYearMonthColumns(monthlyValues);
  if (!yearMonthCols) return;

  const yearsInHeader = Object.keys(yearMonthCols).map(y => parseInt(y, 10));

  // Analyze sheet structure
  const structure = analyzeSheetStructure(monthlyValues);

  // 4. 벤더 매칭 분석
  Logger.log('\n========== STEP 3: Vendor Matching ==========');
  const monthlyVendors = new Set(Object.keys(structure.vendors));
  Logger.log('Vendors found in MONTHLY sheet: ' + monthlyVendors.size);
  Logger.log('MONTHLY Vendors:');

  for (const vendor in structure.vendors) {
    const isProtected = isProtectedLabel(vendor);
    const marker = isProtected ? ' ⚠️ PROTECTED - SHOULD NOT BE HERE!' : '';
    Logger.log(`  "${vendor}" -> Row ${structure.vendors[vendor].row + 1} (${structure.vendors[vendor].section})${marker}`);
  }

  const newVendors = [...allInputVendors].filter(v => !monthlyVendors.has(v));
  if (newVendors.length > 0) {
    Logger.log('\n⚠️  WARNING: New vendors in INPUT not found in MONTHLY:');
    Logger.log('  ' + newVendors.join(', '));
  }

  // 5. 데이터 매칭 분석
  Logger.log('\n========== STEP 4: Data Matching Analysis ==========');
  let matchCount = 0;
  let mismatchCount = 0;

  for (const vendorName in summary) {
    if (!structure.vendors[vendorName]) {
      mismatchCount++;
      Logger.log(`\n❌ Vendor "${vendorName}" from INPUT NOT FOUND in MONTHLY`);

      // Check for similar names
      const similar = [...monthlyVendors].filter(v =>
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

    for (const year in summary[vendorName]) {
      const colsForYear = yearMonthCols[year];
      if (!colsForYear) {
        Logger.log(`   ❌ Year ${year} NOT FOUND in MONTHLY header`);
        continue;
      }

      Logger.log(`   Year ${year}:`);
      for (const month in summary[vendorName][year]) {
        const colIndex = colsForYear[month];
        if (!colIndex) {
          Logger.log(`     ❌ Month ${month} NOT FOUND in MONTHLY header`);
          continue;
        }

        const amount = summary[vendorName][year][month];
        Logger.log(`     ✅ Month ${month}: ${amount} -> Cell(Row ${vendorRow + 1}, Col ${colIndex})`);
      }
    }
  }

  // 6. 요약 출력
  Logger.log('\n========== SUMMARY ==========');
  Logger.log('Total vendors in INPUT: ' + allInputVendors.size);
  Logger.log('Matched vendors: ' + matchCount);
  Logger.log('Unmatched vendors: ' + mismatchCount);
  Logger.log('========== DEBUG MODE END ==========');

  SpreadsheetApp.getUi().alert(
    'DEBUG 완료!\n\n' +
    'INPUT 벤더: ' + allInputVendors.size + '개\n' +
    'MONTHLY 벤더: ' + monthlyVendors.size + '개\n' +
    '매칭된 벤더: ' + matchCount + '개\n' +
    '매칭 안된 벤더: ' + mismatchCount + '개\n\n' +
    '자세한 내용은 보기 > 로그를 확인하세요.'
  );
}
