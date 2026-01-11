  /**
  * @OnlyCurrentDoc
  *
  * Vendor_Main.gs
  * VENDOR 시트 동기화 메인 로직
  */

  /**
  * VENDOR 시트의 Year/Month 헤더를 파싱하여 컬럼 매핑 생성
  * @param {any[][]} vendorValues - VENDOR 시트의 모든 값
  * @return {object} yearMonthCols - { year: { month: startColIndex } }
  *   startColIndex는 해당 월의 첫 번째 칸(1-1)의 컬럼 인덱스 (1-based)
  */
  function parseVendorYearMonthColumns(vendorValues) {
    Logger.log('\n========== Parsing VENDOR Year/Month Headers ==========');

    const yearHeader = vendorValues[HEADER_ROWS.VENDOR.YEAR - 1] || [];
    const monthHeader = vendorValues[HEADER_ROWS.VENDOR.MONTH - 1] || [];

    Logger.log('Year header (row 3): ' + yearHeader.slice(0, 30).join(' | '));
    Logger.log('Month header (row 4): ' + monthHeader.slice(0, 30).join(' | '));

    const yearMonthCols = {}; // { year: { month: startColIndex } }

    Logger.log('\nParsing year/month columns:');

    // B열(인덱스 1)은 건너뛰고 C열(인덱스 2)부터 시작
    for (let c = 2; c < monthHeader.length; c += VENDOR_CELLS_PER_MONTH) {
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
        month = parseMonthNameToNumber(rawMonth);
      }

      if (month < 1 || month > 12) {
        Logger.log(`  Col ${c + 1}: Skipped (month "${rawMonth}" -> ${month} is invalid)`);
        continue;
      }

      if (!yearMonthCols[year]) yearMonthCols[year] = {};
      yearMonthCols[year][month] = c + 1; // 1-based column (해당 월의 첫 번째 칸)
      Logger.log(`  Col ${c + 1}: Year ${year}, Month ${month} (${rawMonth}) - starts at column ${c + 1}`);
    }

    if (Object.keys(yearMonthCols).length === 0) {
      Logger.log('\nERROR: No valid year/month columns found!');
      SpreadsheetApp.getUi().alert('Error: VENDOR 시트의 3행(연도)과 4행(월)에서 유효한 컬럼을 찾지 못했습니다.');
      return null;
    }

    Logger.log('\nYear/Month column mapping:');
    Logger.log(JSON.stringify(yearMonthCols, null, 2));

    return yearMonthCols;
  }

  /**
  * VENDOR 시트의 구조를 분석 (MONTHLY와 동일한 로직 사용)
  * @param {any[][]} values - VENDOR 시트의 2D 배열 값
  * @return {object} 시트 구조 정보
  */
  function analyzeVendorSheetStructure(values) {
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
        structure.gmGrandTotalRow = i;
      }
    }

    let sectionIndex = 0;

    // Process first two sections (with MONTH headers)
    for (let s = 0; s < monthRows.length && s < subtotalRows.length; s++) {
      const sectionName = `Section${sectionIndex + 1}`;
      const startRow = monthRows[s] + 1;
      const endRow = subtotalRows[s] - 1;

      structure.sections[sectionName] = {
        headerRow: monthRows[s],
        vendors: {},
        monthRows: [monthRows[s]],
        subtotalRows: [subtotalRows[s]],
        vendorStartRow: startRow,
        vendorEndRow: endRow
      };

      for (let i = startRow; i <= endRow; i++) {
        const vendorName = String(values[i][0] || '').trim();
        if (vendorName) {
          const upperVendorName = vendorName.toUpperCase();
          if (PROTECTED_LABELS.some(label => upperVendorName.includes(label))) {
            continue;
          }
          structure.vendors[vendorName] = { row: i, section: sectionName };
          structure.sections[sectionName].vendors[vendorName] = i;
        }
      }

      sectionIndex++;
    }

    // Process third section (no MONTH header)
    if (subtotalRows.length >= 3) {
      const sectionName = `Section${sectionIndex + 1}`;
      const startRow = subtotalRows[1] + 2;
      const endRow = subtotalRows[2] - 1;

      structure.sections[sectionName] = {
        headerRow: -1,
        vendors: {},
        monthRows: [],
        subtotalRows: [subtotalRows[2]],
        vendorStartRow: startRow,
        vendorEndRow: endRow
      };

      for (let i = startRow; i <= endRow; i++) {
        const vendorName = String(values[i][0] || '').trim();
        if (vendorName) {
          const upperVendorName = vendorName.toUpperCase();
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
  * VENDOR 시트 동기화 메인 함수
  */
  function syncVendorSummary() {
    Logger.log('========== VENDOR 동기화 시작 ==========');

    // 1. 시트 존재 확인
    if (!checkMultipleSheetsExist(SHEET_NAMES.INPUT, SHEET_NAMES.VENDOR, SHEET_NAMES.BASIC)) {
      SpreadsheetApp.getUi().alert('Error: INPUT, VENDOR, BASIC 시트를 찾을 수 없습니다.');
      return;
    }

    const vendorSheet = getSheet(SHEET_NAMES.VENDOR);

    // 2. BASIC 시트에서 payment method 읽기
    const paymentMethodMap = readBasicPaymentMethods();

    // 3. ETC 벤더 목록 읽기
    const etcVendors = getEtcVendorsFromDetailsSheet();

    // 4. INPUT 시트에서 인보이스 데이터 읽기
    const invoicesData = readVendorInvoicesFromInput(paymentMethodMap);

    // 5. VENDOR 시트 분석
    const vendorRange = vendorSheet.getDataRange();
    const vendorValues = vendorRange.getValues();

    // Parse year/month column headers
    const yearMonthCols = parseVendorYearMonthColumns(vendorValues);
    if (!yearMonthCols) return;

    // Analyze sheet structure
    const structure = analyzeVendorSheetStructure(vendorValues);

    Logger.log(`\nFound ${Object.keys(structure.vendors).length} vendors in VENDOR sheet`);

    // 6. ETC 벤더가 VENDOR 시트에 있는 경우, ETC 상세 시트의 실제 벤더들을 합산
    if (structure.vendors['ETC'] && etcVendors.size > 0) {
      Logger.log('\n========== Processing ETC Vendor ==========');
      Logger.log(`ETC row found at: ${structure.vendors['ETC'].row + 1}`);
      Logger.log(`ETC contains ${etcVendors.size} vendors from ETC 상세 sheet`);

      // ETC 벤더들의 데이터를 합산하여 invoicesData['ETC']에 저장
      invoicesData['ETC'] = {};

      // Logger batching for ETC processing (Phase 1 optimization)
      const etcLogBuffer = [];

      for (const etcVendorName of etcVendors) {
        if (invoicesData[etcVendorName]) {
          etcLogBuffer.push(`  Adding ${etcVendorName} to ETC aggregation`);

          for (const year in invoicesData[etcVendorName]) {
            if (!invoicesData['ETC'][year]) invoicesData['ETC'][year] = {};

            for (const month in invoicesData[etcVendorName][year]) {
              if (!invoicesData['ETC'][year][month]) invoicesData['ETC'][year][month] = [];

              // ETC 벤더의 모든 인보이스를 ETC에 추가
              invoicesData['ETC'][year][month].push(...invoicesData[etcVendorName][year][month]);
            }
          }
        }
      }

      // Batch log output for ETC
      if (etcLogBuffer.length > 0) {
        Logger.log(etcLogBuffer.join('\n'));
      }

      Logger.log(`ETC aggregation completed`);
    }

    // 7. 보호된 행 백업
    const protectedRowsData = backupProtectedRows(vendorSheet, vendorValues);
    const protectedRowIndices = getProtectedRowIndices(protectedRowsData.protectedRows);

    // 8. 데이터 준비 (VENDOR 행만 업데이트, 나머지는 보존)
    Logger.log('\n========== Preparing Output Data ==========');
    let outputValues = JSON.parse(JSON.stringify(vendorValues));

    Logger.log(`Found ${Object.keys(structure.vendors).length} vendors in VENDOR sheet`);
    Logger.log(`Protected row indices: ${[...protectedRowIndices].join(', ')}`);

    // 9. 각 벤더의 데이터 채우기
    Logger.log('\n========== Populating Vendor Invoice Data ==========');

    // 강조 셀 추적을 위한 배열 (빨간색 텍스트로 표시할 셀)
    const highlightedCells = [];

    // Logger batching for performance (Phase 1 optimization)
    const logBuffer = [];

    for (const vendorName in structure.vendors) {
      const vendorRow = structure.vendors[vendorName].row;

      // 보호된 행이면 건너뛰기
      if (isProtectedRow(vendorRow, protectedRowIndices)) {
        logBuffer.push(`⚠️ Skipping protected row ${vendorRow + 1} (vendor: ${vendorName})`);
        continue;
      }

      // 해당 벤더의 인보이스 데이터가 있는지 확인
      if (!invoicesData[vendorName]) {
        logBuffer.push(`No invoice data for vendor: ${vendorName}`);
        continue;
      }

      logBuffer.push(`\nProcessing vendor: ${vendorName} (row ${vendorRow + 1})`);

      // 각 year/month에 대해 처리
      for (const year in yearMonthCols) {
        const monthsInYear = yearMonthCols[year];

        for (const month in monthsInYear) {
          const startCol = monthsInYear[month]; // 1-based column index

          // 날짜 필터: 필터 기준 이전 데이터는 건드리지 않음 (기존 값 보존)
          if (!shouldProcessInvoiceDate(parseInt(year, 10), parseInt(month, 10))) {
            continue; // 8월 이전 칸은 업데이트하지 않고 기존 값 유지
          }

          // 해당 year/month의 인보이스가 있는지 확인
          if (!invoicesData[vendorName][year] || !invoicesData[vendorName][year][month]) {
            // 인보이스 없음 - 8칸 모두 비우기 (8월 이후만)
            for (let i = 0; i < VENDOR_CELLS_PER_MONTH; i++) {
              outputValues[vendorRow][startCol - 1 + i] = '';
            }
            continue;
          }

          // 인보이스 데이터 가져오기 및 제한 (vendorName 전달)
          const invoices = limitAndMergeInvoices(invoicesData[vendorName][year][month], vendorName);

          logBuffer.push(`  ${year}-${month}: ${invoices.length} invoices`);

          // 8칸 채우기 (최대 4개 인보이스)
          for (let i = 0; i < VENDOR_MAX_INVOICES; i++) {
            const amountColIndex = startCol - 1 + (i * 2);     // 0-based
            const payDateColIndex = startCol - 1 + (i * 2) + 1; // 0-based

            if (i < invoices.length) {
              const invoice = invoices[i];

              // Amount 칸
              outputValues[vendorRow][amountColIndex] = invoice.amount;

              // Payment Date + Method 칸
              const dateStr = formatPaymentDate(invoice.payMonth, invoice.payDate);
              const methodStr = formatPaymentMethod(invoice);
              outputValues[vendorRow][payDateColIndex] = `${dateStr}${methodStr}`;

              // 강조 조건(Paid='O')인 경우 빨간색 표시를 위해 기록 (1-based row, col)
              if (invoice.isHighlighted) {
                highlightedCells.push({
                  row: vendorRow + 1,
                  col: amountColIndex + 1
                });
                logBuffer.push(`    Invoice ${i + 1}: $${invoice.amount} ${dateStr}${methodStr} [HIGHLIGHT]`);
              } else {
                logBuffer.push(`    Invoice ${i + 1}: $${invoice.amount} ${dateStr}${methodStr}`);
              }
            } else {
              // 인보이스 없음 - 빈 칸
              outputValues[vendorRow][amountColIndex] = '';
              outputValues[vendorRow][payDateColIndex] = '';
            }
          }
        }
      }
    }

    // Batch log output (Phase 1 optimization)
    if (logBuffer.length > 0) {
      Logger.log(logBuffer.join('\n'));
    }

    // 10. 시트에 값 쓰기
    Logger.log('\n========== Writing Values to Sheet ==========');
    vendorRange.setValues(outputValues);

    // 11. 보호된 행 복원
    restoreProtectedRows(vendorSheet, protectedRowsData.protectedRows, protectedRowsData.protectedFormulas, protectedRowsData.protectedValues);

    // 12. 모든 데이터 셀에 색상 적용
    Logger.log('\n========== Applying Colors to All Data Cells ==========');
    for (const vendorName in structure.vendors) {
      const vendorRow = structure.vendors[vendorName].row;

      // 보호된 행이면 색상 적용도 건너뛰기
      if (isProtectedRow(vendorRow, protectedRowIndices)) {
        continue;
      }

      for (const year in yearMonthCols) {
        const monthsInYear = yearMonthCols[year];
        for (const month in monthsInYear) {
          const startCol = monthsInYear[month];

          // 각 인보이스 슬롯 처리 (최대 4개)
          for (let i = 0; i < VENDOR_MAX_INVOICES; i++) {
            const amountColIndex = startCol - 1 + (i * 2);     // 0-based
            const payDateColIndex = startCol - 1 + (i * 2) + 1; // 0-based

            const amountCell = vendorSheet.getRange(vendorRow + 1, amountColIndex + 1);
            const payDateCell = vendorSheet.getRange(vendorRow + 1, payDateColIndex + 1);

            // 기본 서식 설정
            amountCell.setHorizontalAlignment('right').setFontSize(8);
            payDateCell.setHorizontalAlignment('center').setFontSize(8);

            // 값이 있는 셀만 색상 적용
            if (outputValues[vendorRow][amountColIndex] !== '') {
              // 강조 여부 확인
              const isHighlightedCell = highlightedCells.some(c =>
                c.row === vendorRow + 1 && c.col === amountColIndex + 1
              );

              if (isHighlightedCell) {
                // Highlighted: Amount (빨간색) - Date/Method (검은색)
                amountCell.setFontColor('#FF0000');
                payDateCell.setFontColor('#000000');
                Logger.log(`  Row ${vendorRow + 1}, Cols ${amountColIndex + 1}-${payDateColIndex + 1}: Highlighted (Red-Black)`);
              } else {
                // Normal: Amount (검은색) - Date/Method (검은색)
                amountCell.setFontColor('#000000');
                payDateCell.setFontColor('#000000');
                Logger.log(`  Row ${vendorRow + 1}, Cols ${amountColIndex + 1}-${payDateColIndex + 1}: Normal (Black-Black)`);
              }
            }
          }
        }
      }
    }

    // 13. 벤더 행 배경색 적용
    applyAlternatingVendorRowColors(vendorSheet, structure);

    // 14. 완료 메시지 (LOG 시트에 기록)
    Logger.log('\n========== VENDOR 동기화 완료 ==========');
    const etcVendorCount = etcVendors.size;
    const logMsg = etcVendorCount > 0
      ? `VENDOR 시트 업데이트 완료 | Highlighted: ${highlightedCells.length}개 | ETC 벤더: ${etcVendorCount}개`
      : `VENDOR 시트 업데이트 완료 | Highlighted: ${highlightedCells.length}개`;

    writeToLog('VENDOR', logMsg);
  }

  /**
  * 벤더 행에 섹션별 교차 배경색 적용
  * @param {Sheet} sheet - VENDOR 시트
  * @param {object} structure - 시트 구조 정보
  */
  function applyAlternatingVendorRowColors(sheet, structure) {
    Logger.log('\n========== Applying Alternating Vendor Row Colors ==========');

    // 섹션별 참조 색상 읽기
    const hairColor = sheet.getRange(6, 1).getBackground();  // A6 (HAIR 섹션 색상)
    const gmColor = '#c4bee2';   // A12 (GM 섹션 색상)
    const whiteColor = '#ffffff';

    Logger.log(`HAIR section color (A6): ${hairColor}`);
    Logger.log(`GM section color (A12): ${gmColor}`);

    // 섹션별 색상 매핑
    const sectionColors = {
      'Section1': hairColor,  // HAIR
      'Section2': gmColor,    // GM
      'Section3': gmColor     // HARWIN
    };

    // 각 섹션 처리
    for (const sectionName in structure.sections) {
      const section = structure.sections[sectionName];
      const sectionColor = sectionColors[sectionName];

      if (!sectionColor) {
        Logger.log(`⚠️ No color defined for ${sectionName}, skipping`);
        continue;
      }

      Logger.log(`\nProcessing ${sectionName}:`);
      Logger.log(`  Start row: ${section.vendorStartRow + 1}`);
      Logger.log(`  End row: ${section.vendorEndRow + 1}`);
      Logger.log(`  Section color: ${sectionColor}`);

      // 보이는 행만 수집
      const visibleRows = [];
      for (let row = section.vendorStartRow; row <= section.vendorEndRow; row++) {
        if (!sheet.isRowHiddenByUser(row + 1)) {  // 1-based row number
          visibleRows.push(row);
        }
      }

      Logger.log(`  Visible rows: ${visibleRows.length}`);

      // 교차 색상 적용
      let colorIndex = 0;
      for (const row of visibleRows) {
        const bgColor = (colorIndex % 2 === 0) ? whiteColor : sectionColor;
        const rowRange = sheet.getRange(row + 1, 1, 1, sheet.getMaxColumns());
        rowRange.setBackground(bgColor);

        Logger.log(`  Row ${row + 1}: ${bgColor === whiteColor ? 'White' : 'Colored'}`);
        colorIndex++;
      }
    }

    Logger.log('\n✅ Alternating row colors applied successfully');
  }
