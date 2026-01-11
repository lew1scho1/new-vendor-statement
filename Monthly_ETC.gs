/**
 * @OnlyCurrentDoc
 *
 * Monthly_ETC.gs
 * ETC 상세 시트 관련 로직
 */

/**
 * ETC 상세 시트 생성 또는 업데이트
 * @param {object} etcDetails - ETC 벤더별 상세 데이터 { vendor: { year: { month: amount } } }
 * @param {object} yearMonthCols - Year/Month 컬럼 매핑 { year: { month: colIndex } }
 */
function updateEtcDetailsSheet(etcDetails, yearMonthCols) {
  Logger.log('\n========== Updating ETC Details Sheet ==========');

  const ss = getActiveSpreadsheet();
  let etcSheet = ss.getSheetByName(SHEET_NAMES.ETC_DETAILS);

  if (!etcSheet) {
    etcSheet = ss.insertSheet(SHEET_NAMES.ETC_DETAILS);
    Logger.log(`Created new sheet: ${SHEET_NAMES.ETC_DETAILS}`);
  }

  // 기존 A열 (회사 이름) 읽기
  const existingVendors = [];
  const lastRow = etcSheet.getLastRow();
  if (lastRow > 1) {
    const vendorColumn = etcSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < vendorColumn.length; i++) {
      const vendorName = String(vendorColumn[i][0] || '').trim();
      if (vendorName) {
        existingVendors.push(vendorName);
      }
    }
    Logger.log(`기존 ETC 벤더 ${existingVendors.length}개 유지: ${existingVendors.join(', ')}`);
  }

  // Prepare headers: Vendor, Year-Month columns
  const allYears = Object.keys(yearMonthCols).sort();
  const headers = ['Vendor'];

  // Add year/month columns
  for (const year of allYears) {
    const months = Object.keys(yearMonthCols[year]).sort((a, b) => parseInt(a) - parseInt(b));
    for (const month of months) {
      headers.push(`${year}-${String(month).padStart(2, '0')}`);
    }
  }

  Logger.log(`Headers: ${headers.join(', ')}`);

  // 데이터 영역만 초기화 (헤더와 A열 제외한 나머지)
  if (lastRow > 1) {
    const maxCols = etcSheet.getMaxColumns();
    if (maxCols > 1) {
      etcSheet.getRange(2, 2, lastRow - 1, maxCols - 1).clearContent();
    }
  }

  // 헤더 업데이트
  etcSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 기존 A열 벤더에 대한 데이터만 업데이트
  const rows = [];
  for (const vendorName of existingVendors) {
    const row = [vendorName];

    for (const year of allYears) {
      const months = Object.keys(yearMonthCols[year]).sort((a, b) => parseInt(a) - parseInt(b));
      for (const month of months) {
        const amount = etcDetails[vendorName]?.[year]?.[month] || 0;
        row.push(amount);
      }
    }

    rows.push(row);
  }

  // Write to sheet
  if (rows.length > 0) {
    etcSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    Logger.log(`ETC Details sheet updated with ${rows.length} vendors`);
  } else {
    Logger.log('No existing vendors found in ETC Details sheet');
  }

  // Format header row
  const headerRange = etcSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8f0fe');

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    etcSheet.autoResizeColumn(i);
  }
}

/**
 * ETC 집계 로직 - ETC 상세 시트의 A열에 있는 벤더들을 ETC로 모음
 * @param {object} summary - INPUT 데이터 집계 { vendor: { year: { month: amount } } }
 * @return {object} { etcDetails, etcSummary, etcVendors }
 */
function aggregateEtcData(summary) {
  Logger.log('\n========== Aggregating ETC Data ==========');

  // ETC 상세 시트에서 벤더 목록 읽기
  const etcVendors = getEtcVendorsFromDetailsSheet();

  const etcDetails = {}; // { vendor: { year: { month: amount } } }
  const etcSummary = {}; // { year: { month: totalAmount } }

  if (etcVendors.size > 0) {
    Logger.log(`ETC 벤더 (${etcVendors.size}개): ${[...etcVendors].join(', ')}`);

    // ETC 상세 시트에 정의된 벤더들의 데이터를 ETC로 집계
    for (const vendorName of etcVendors) {
      if (summary[vendorName]) {
        etcDetails[vendorName] = summary[vendorName];

        for (const year in summary[vendorName]) {
          if (!etcSummary[year]) etcSummary[year] = {};
          for (const month in summary[vendorName][year]) {
            if (!etcSummary[year][month]) etcSummary[year][month] = 0;
            etcSummary[year][month] += summary[vendorName][year][month];
          }
        }
      }
    }

    Logger.log(`ETC Summary:`, JSON.stringify(etcSummary, null, 2));
  } else {
    Logger.log('No ETC vendors defined in ETC 상세 sheet - ETC will be empty');
  }

  return {
    etcDetails,
    etcSummary,
    etcVendors: [...etcVendors]
  };
}
