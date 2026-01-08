/**
 * @OnlyCurrentDoc
 *
 * DataReader.gs
 * INPUT 시트에서 데이터를 읽고 집계하는 함수들
 */

/**
 * INPUT 시트에서 데이터를 읽어서 집계
 * @return {object} 집계된 데이터 { summary, allInputVendors, validRowCount, invalidRowCount }
 *   - summary: { vendor: { year: { month: amount } } }
 *   - allInputVendors: Set of vendor names
 *   - validRowCount: number of valid rows
 *   - invalidRowCount: number of invalid rows
 */
function readAndAggregateInputData() {
  Logger.log('\n========== Reading INPUT Sheet ==========');

  const inputSheet = getSheet(SHEET_NAMES.INPUT);

  if (!inputSheet) {
    Logger.log('ERROR: Could not find INPUT sheet.');
    SpreadsheetApp.getUi().alert('Error: Could not find INPUT sheet.');
    return null;
  }

  const inputData = inputSheet.getDataRange().getValues().slice(1); // Exclude header
  Logger.log('Total INPUT rows (excluding header): ' + inputData.length);

  const summary = {}; // { vendor: { year: { month: amount } } }
  const allInputVendors = new Set();
  let validRowCount = 0;
  let invalidRowCount = 0;
  let filteredCount = 0;

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    const vendor = row[COLUMN_INDICES.INPUT.VENDOR - 1];
    const year = parseInt(row[COLUMN_INDICES.INPUT.YEAR - 1], 10);
    const month = parseInt(row[COLUMN_INDICES.INPUT.MONTH - 1], 10);
    const amount = parseFloat(row[COLUMN_INDICES.INPUT.AMOUNT - 1]);

    if (vendor && year && month && !isNaN(amount)) {
      // 날짜 필터 적용 (Common.gs의 DATA_FILTER_FROM_DATE 기준)
      if (!shouldProcessInvoiceDate(year, month)) {
        filteredCount++;
        continue;
      }
      validRowCount++;
      allInputVendors.add(vendor);
      if (!summary[vendor]) summary[vendor] = {};
      if (!summary[vendor][year]) summary[vendor][year] = {};
      if (!summary[vendor][year][month]) summary[vendor][year][month] = 0;
      summary[vendor][year][month] += amount;

      // Log first 5 valid rows as sample
      if (validRowCount <= 5) {
        Logger.log(`Sample row ${i + 2}: Vendor="${vendor}", Year=${year}, Month=${month}, Amount=${amount}`);
      }
    } else {
      invalidRowCount++;
      if (invalidRowCount <= 5) {
        Logger.log(`Invalid row ${i + 2}: Vendor="${vendor}", Year=${year}, Month=${month}, Amount=${amount}`);
      }
    }
  }

  Logger.log('\nValid rows processed: ' + validRowCount);
  Logger.log('Invalid rows skipped: ' + invalidRowCount);
  if (filteredCount > 0 && DATA_FILTER_FROM_DATE) {
    Logger.log(`⚡ Filtered out ${filteredCount} rows before ${DATA_FILTER_FROM_DATE.year}/${DATA_FILTER_FROM_DATE.month}`);
  }
  Logger.log('Unique vendors in INPUT: ' + allInputVendors.size);
  Logger.log('INPUT Vendors: ' + [...allInputVendors].join(', '));

  Logger.log('\nSummary data structure:');
  for (const vendor in summary) {
    Logger.log(`\n  Vendor: "${vendor}"`);
    for (const year in summary[vendor]) {
      Logger.log(`    Year ${year}:`);
      for (const month in summary[vendor][year]) {
        Logger.log(`      Month ${month}: ${summary[vendor][year][month]}`);
      }
    }
  }

  return {
    summary,
    allInputVendors,
    validRowCount,
    invalidRowCount
  };
}

/**
 * MONTHLY 시트의 Year/Month 헤더를 파싱하여 컬럼 매핑 생성
 * @param {any[][]} monthlyValues - MONTHLY 시트의 모든 값
 * @return {object} yearMonthCols - { year: { month: colIndex(1-based) } }
 */
function parseMonthlyYearMonthColumns(monthlyValues) {
  Logger.log('\n========== Parsing MONTHLY Year/Month Headers ==========');

  const yearHeader = monthlyValues[HEADER_ROWS.MONTHLY.YEAR - 1] || [];
  const monthHeader = monthlyValues[HEADER_ROWS.MONTHLY.MONTH - 1] || [];

  Logger.log('Year header (row 4): ' + yearHeader.slice(0, 20).join(' | '));
  Logger.log('Month header (row 5): ' + monthHeader.slice(0, 20).join(' | '));

  const yearMonthCols = {}; // { year: { month: colIndex(1-based) } }

  Logger.log('\nParsing year/month columns:');
  for (let c = 0; c < monthHeader.length; c++) {
    const rawYear = yearHeader[c];
    const rawMonth = monthHeader[c];
    if (!rawYear || !rawMonth) continue;

    const year = parseInt(rawYear, 10);
    if (isNaN(year)) {
      Logger.log(`  Col ${c + 1}: Skipped (year "${rawYear}" is not a number)`);
      continue;
    }

    let month = parseInt(rawMonth, 10);
    if (isNaN(month)) {
      const upper = String(rawMonth).trim().toUpperCase();
      if (upper === 'TOTAL') {
        Logger.log(`  Col ${c + 1}: Skipped (TOTAL column)`);
        continue;
      }
      month = parseMonthNameToNumber(rawMonth);
    }

    if (month < 1 || month > 12) {
      Logger.log(`  Col ${c + 1}: Skipped (month "${rawMonth}" -> ${month} is invalid)`);
      continue;
    }

    if (!yearMonthCols[year]) yearMonthCols[year] = {};
    yearMonthCols[year][month] = c + 1; // 1-based column
    Logger.log(`  Col ${c + 1}: Year ${year}, Month ${month} (${rawMonth})`);
  }

  if (Object.keys(yearMonthCols).length === 0) {
    Logger.log('\nERROR: No valid year/month columns found!');
    SpreadsheetApp.getUi().alert('Error: 4행(연도)과 5행(월)에서 유효한 컬럼을 찾지 못했습니다.');
    return null;
  }

  Logger.log('\nYear/Month column mapping:');
  Logger.log(JSON.stringify(yearMonthCols, null, 2));

  const yearsInHeader = Object.keys(yearMonthCols).map(y => parseInt(y, 10));
  Logger.log('Years found in MONTHLY header: ' + yearsInHeader.join(', '));

  return yearMonthCols;
}
