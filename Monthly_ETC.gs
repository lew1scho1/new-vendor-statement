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
  } else {
    etcSheet.clear();
    Logger.log(`Cleared existing sheet: ${SHEET_NAMES.ETC_DETAILS}`);
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

  // Prepare data rows
  const rows = [headers];

  for (const vendorName in etcDetails) {
    const row = [vendorName];

    for (const year of allYears) {
      const months = Object.keys(yearMonthCols[year]).sort((a, b) => parseInt(a) - parseInt(b));
      for (const month of months) {
        const amount = etcDetails[vendorName][year]?.[month] || 0;
        row.push(amount);
      }
    }

    rows.push(row);
    Logger.log(`Added ETC vendor: ${vendorName}`);
  }

  // Write to sheet
  if (rows.length > 1) {
    etcSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

    // Format header row
    const headerRange = etcSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      etcSheet.autoResizeColumn(i);
    }

    Logger.log(`ETC Details sheet updated with ${rows.length - 1} vendors`);
  } else {
    Logger.log('No ETC data to write');
  }
}

/**
 * ETC 집계 로직 - 매칭되지 않은 벤더들을 ETC로 모음
 * @param {Set} allInputVendors - INPUT의 모든 벤더 Set
 * @param {Set} monthlyVendors - MONTHLY의 모든 벤더 Set
 * @param {object} summary - INPUT 데이터 집계 { vendor: { year: { month: amount } } }
 * @return {object} { etcDetails, etcSummary, unmatchedVendors }
 */
function aggregateEtcData(allInputVendors, monthlyVendors, summary) {
  Logger.log('\n========== Aggregating ETC Data ==========');

  const unmatchedVendors = [...allInputVendors].filter(v => !monthlyVendors.has(v));
  const etcDetails = {}; // { vendor: { year: { month: amount } } }
  const etcSummary = {}; // { year: { month: totalAmount } }

  if (unmatchedVendors.length > 0) {
    Logger.log(`Unmatched vendors will be added to ETC: ${unmatchedVendors.join(', ')}`);

    // Aggregate unmatched vendor data for ETC
    for (const vendorName of unmatchedVendors) {
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
    Logger.log('No unmatched vendors - ETC will be empty');
  }

  return {
    etcDetails,
    etcSummary,
    unmatchedVendors
  };
}
