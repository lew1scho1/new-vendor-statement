/**
 * @OnlyCurrentDoc
 *
 * Monthly_Main.gs
 * MONTHLY 시트 동기화 메인 로직
 */

/**
 * MONTHLY 시트의 구조를 분석하여 섹션, 벤더, 토탈 행 위치를 파악
 * @param {any[][]} values - MONTHLY 시트의 2D 배열 값
 * @return {object} 시트 구조 정보
 */
function analyzeSheetStructure(values) {
  const structure = {
    vendors: {},
    sections: {},
    gmGrandTotalRow: -1,
    grandTotalRow: 0
  };

  // Find all MONTH and SUBTOTAL rows
  const monthRows = [];
  const subtotalRows = [];

  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0] || '').trim().toUpperCase();

    if (cell === MONTH_ROW_LABEL.toUpperCase()) {
      monthRows.push(i);
    } else if (cell === SUBTOTAL_LABEL.toUpperCase()) {
      subtotalRows.push(i);
    } else if (cell.replace(/\s+/g, ' ') === GM_GRAND_TOTAL_LABEL.toUpperCase().replace(/\s+/g, ' ')) {
      // Normalize spaces for GM GRAND TOTAL matching
      structure.gmGrandTotalRow = i;
    }
  }

  // Define sections based on MONTH/SUBTOTAL pairs and remaining vendors
  let sectionIndex = 0;

  // Process first two sections (with MONTH headers)
  for (let s = 0; s < monthRows.length && s < subtotalRows.length; s++) {
    const sectionName = `Section${sectionIndex + 1}`;
    const startRow = monthRows[s] + 1; // Start after MONTH row
    const endRow = subtotalRows[s] - 1; // End before SUBTOTAL row

    structure.sections[sectionName] = {
      headerRow: monthRows[s],
      vendors: {},
      monthRows: [monthRows[s]],
      subtotalRows: [subtotalRows[s]],
      vendorStartRow: startRow,
      vendorEndRow: endRow
    };

    // Add all vendors in this range, excluding protected labels
    for (let i = startRow; i <= endRow; i++) {
      const vendorName = String(values[i][0] || '').trim();
      if (vendorName) {
        const upperVendorName = vendorName.toUpperCase();
        // Skip if it's a protected label
        if (PROTECTED_LABELS.some(label => upperVendorName.includes(label))) {
          continue;
        }
        structure.vendors[vendorName] = { row: i, section: sectionName };
        structure.sections[sectionName].vendors[vendorName] = i;
      }
    }

    sectionIndex++;
  }

  // Process third section (no MONTH header, starts after previous SUBTOTAL)
  if (subtotalRows.length >= 3) {
    const sectionName = `Section${sectionIndex + 1}`;
    const startRow = subtotalRows[1] + 2; // Skip blank row after 2nd SUBTOTAL
    const endRow = subtotalRows[2] - 1; // End before 3rd SUBTOTAL

    structure.sections[sectionName] = {
      headerRow: -1,
      vendors: {},
      monthRows: [],
      subtotalRows: [subtotalRows[2]],
      vendorStartRow: startRow,
      vendorEndRow: endRow
    };

    // Add all vendors in this range, excluding protected labels
    for (let i = startRow; i <= endRow; i++) {
      const vendorName = String(values[i][0] || '').trim();
      if (vendorName) {
        const upperVendorName = vendorName.toUpperCase();
        // Skip if it's a protected label
        if (PROTECTED_LABELS.some(label => upperVendorName.includes(label))) {
          continue;
        }
        structure.vendors[vendorName] = { row: i, section: sectionName };
        structure.sections[sectionName].vendors[vendorName] = i;
      }
    }
  }

  return structure;
}

/**
 * MONTHLY 시트 동기화 메인 함수
 */
function syncMonthlySummary() {
  Logger.log('========== MONTHLY 동기화 시작 ==========');

  // 1. 시트 존재 확인
  if (!checkSheetsExist(SHEET_NAMES.INPUT, SHEET_NAMES.MONTHLY)) {
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
  const monthlyRange = monthlySheet.getDataRange();
  const monthlyValues = monthlyRange.getValues();

  // Parse year/month column headers
  const yearMonthCols = parseMonthlyYearMonthColumns(monthlyValues);
  if (!yearMonthCols) return;

  const yearsInHeader = Object.keys(yearMonthCols).map(y => parseInt(y, 10));

  // Analyze sheet structure (vendors, sections, protected rows)
  const structure = analyzeSheetStructure(monthlyValues);

  // 4. ETC 집계 (ETC 상세 시트의 A열 기준)
  const etcData = aggregateEtcData(summary);
  const { etcDetails, etcSummary, etcVendors } = etcData;

  // 5. 보호된 행과 열 백업
  const protectedRowsData = backupProtectedRows(monthlySheet, monthlyValues);
  const protectedColumnFormulas = backupProtectedColumnFormulas(monthlySheet, monthlyValues);
  const protectedRowIndices = getProtectedRowIndices(protectedRowsData.protectedRows);

  // 6. 데이터 준비 및 클리어
  Logger.log('\n========== Preparing Output Data ==========');
  let outputValues = JSON.parse(JSON.stringify(monthlyValues));

  Logger.log(`Found ${Object.keys(structure.vendors).length} vendors in MONTHLY sheet`);
  Logger.log(`Protected row indices: ${[...protectedRowIndices].join(', ')}`);

  // Clear previous data ONLY for filtered date range (8월 이후만 클리어, 이전 데이터는 보존)
  for (const vendorName in structure.vendors) {
    // Skip protected rows (SUBTOTAL, etc.)
    if (isProtectedLabel(vendorName)) {
      continue;
    }

    const vendorInfo = structure.vendors[vendorName];

    // ADDITIONAL CHECK: Skip if this row index is in the protected rows
    if (isProtectedRow(vendorInfo.row, protectedRowIndices)) {
      Logger.log(`⚠️ Skipping protected row ${vendorInfo.row + 1} in clearing loop (vendor: ${vendorName})`);
      continue;
    }

    for (const year of yearsInHeader) {
      const colsForYear = yearMonthCols[year];
      if (!colsForYear) continue;
      for (const month in colsForYear) {
        // 날짜 필터: 필터 기준 이후 데이터만 클리어 (이전 데이터는 보존)
        if (!shouldProcessInvoiceDate(year, parseInt(month, 10))) {
          continue; // 8월 이전 데이터는 클리어하지 않음
        }

        const colIndex = colsForYear[month];
        if (outputValues[vendorInfo.row]) {
          outputValues[vendorInfo.row][colIndex - 1] = 0;
        }
      }
    }
  }

  // 7. 새 데이터 채우기
  Logger.log('\n========== Populating Vendor Data ==========');
  const etcVendorSet = new Set(etcVendors);
  for (const vendorName in summary) {
    // ETC 상세 시트에 정의된 벤더는 ETC로 집계하므로 개별 행에 넣지 않음
    if (etcVendorSet.has(vendorName)) {
      Logger.log(`Skipping ${vendorName} (ETC 벤더)`);
      continue;
    }

    if (!structure.vendors[vendorName]) continue;
    const vendorRow = structure.vendors[vendorName].row;

    // ADDITIONAL CHECK: Skip if this row index is in the protected rows
    if (isProtectedRow(vendorRow, protectedRowIndices)) {
      Logger.log(`⚠️ Skipping protected row ${vendorRow + 1} in population loop (vendor: ${vendorName})`);
      continue;
    }

    for (const year in summary[vendorName]) {
      const colsForYear = yearMonthCols[year];
      if (!colsForYear) continue;
      for (const month in summary[vendorName][year]) {
        const colIndex = colsForYear[month];
        if (colIndex) {
          outputValues[vendorRow][colIndex - 1] = summary[vendorName][year][month];
        }
      }
    }
  }

  // 8. ETC 데이터 추가
  if (structure.vendors['ETC'] && Object.keys(etcSummary).length > 0) {
    Logger.log('\n========== Adding ETC Aggregated Data ==========');
    const etcRow = structure.vendors['ETC'].row;
    for (const year in etcSummary) {
      const colsForYear = yearMonthCols[year];
      if (!colsForYear) continue;
      for (const month in etcSummary[year]) {
        const colIndex = colsForYear[month];
        if (colIndex) {
          // Add to existing ETC value (if any)
          const existingValue = outputValues[etcRow][colIndex - 1] || 0;
          outputValues[etcRow][colIndex - 1] = existingValue + etcSummary[year][month];
        }
      }
    }
  }

  // 9. 시트에 값 쓰기
  Logger.log('\n========== Writing Values to Sheet ==========');
  monthlyRange.setValues(outputValues);

  // 10. 보호된 열과 행 복원
  restoreProtectedColumnFormulas(monthlySheet, protectedColumnFormulas);
  restoreProtectedRows(monthlySheet, protectedRowsData.protectedRows, protectedRowsData.protectedFormulas, protectedRowsData.protectedValues);

  // 11. Apply alternating row colors to match VENDOR
  applyAlternatingMonthlyRowColors(monthlySheet, structure);

  // 12. ETC 상세 시트 업데이트
  if (Object.keys(etcDetails).length > 0) {
    updateEtcDetailsSheet(etcDetails, yearMonthCols);
  }

  // 13. 완료 메시지 (LOG 시트에 기록)
  Logger.log('\n========== MONTHLY 동기화 완료 ==========');
  const etcVendorCount = etcVendors.length;
  const logMsg = etcVendorCount > 0
    ? `MONTHLY 시트 업데이트 완료 | ETC 벤더: ${etcVendorCount}개 (${etcVendors.join(', ')})`
    : 'MONTHLY 시트 업데이트 완료';

  writeToLog('MONTHLY', logMsg);
}

/**
 * Apply alternating row colors to MONTHLY based on VENDOR section colors.
 * @param {Sheet} sheet - MONTHLY sheet
 * @param {object} structure - Sheet structure info
 */
function applyAlternatingMonthlyRowColors(sheet, structure) {
  Logger.log('\n========== Applying Alternating Monthly Row Colors ==========');

  const vendorSheet = getSheet(SHEET_NAMES.VENDOR);
  if (!vendorSheet) {
    Logger.log('WARNING: VENDOR sheet not found. Skipping MONTHLY row color application.');
    return;
  }

  // Match VENDOR colors
  const hairColor = vendorSheet.getRange(6, 1).getBackground();
  const gmColor = '#c4bee2';
  const whiteColor = '#ffffff';

  const sectionColors = {
    'Section1': hairColor,
    'Section2': gmColor,
    'Section3': gmColor
  };

  for (const sectionName in structure.sections) {
    const section = structure.sections[sectionName];
    const sectionColor = sectionColors[sectionName];
    if (!sectionColor) {
      Logger.log(`WARNING: No color defined for ${sectionName}, skipping`);
      continue;
    }

    const visibleRows = [];
    for (let row = section.vendorStartRow; row <= section.vendorEndRow; row++) {
      if (!sheet.isRowHiddenByUser(row + 1)) {
        visibleRows.push(row);
      }
    }

    let colorIndex = 0;
    for (const row of visibleRows) {
      const bgColor = (colorIndex % 2 === 0) ? whiteColor : sectionColor;
      sheet.getRange(row + 1, 1, 1, sheet.getMaxColumns()).setBackground(bgColor);
      colorIndex++;
    }
  }

  Logger.log('\nAlternating row colors applied to MONTHLY');
}
