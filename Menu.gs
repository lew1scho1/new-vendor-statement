/*
 * @OnlyCurrentDoc
 *
 * Menu.gs
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('자동화')
      .addItem('1️⃣ 전체 업데이트 (MONTHLY + VENDOR + DETAIL)', 'updateAllSheets')
      .addItem('2️⃣ MONTHLY + VENDOR 업데이트', 'updateMontlyAndVendor')
      .addItem('3️⃣ DETAIL 업데이트', 'updateAllDetailSheets')
      .addSeparator()
      .addItem('데일리 대시보드 이메일 발송', 'sendDashboardEmail')
      .addSeparator()
      .addItem('🔍 ETC 벤더 목록 확인', 'debugEtcVendors')
      .addToUi();
}

/**
 * MONTHLY + VENDOR 업데이트
 * Common.gs의 DATA_FILTER_FROM_DATE 설정에 따라 필터링됩니다.
 * 현재: 2025년 1월부터 처리 (1월 이전 데이터 제외)
 */
function updateMontlyAndVendor() {
  syncMonthlySummary();
  syncVendorSummary();

  writeToLog('SYSTEM', 'MONTHLY + VENDOR 업데이트 완료');
}

/**
 * DETAIL 시트 전체 업데이트
 * Common.gs의 DATA_FILTER_FROM_DATE 설정에 따라 필터링됩니다.
 * 현재: 2025년 1월부터 처리 (1월 이전 데이터 제외)
 */
function updateAllDetailSheets() {
  // Read vendor data once for all detail sheets (Phase 1 optimization)
  Logger.log('========== Reading All Vendor Data (Optimized) ==========');
  const allVendorData = readAllVendorDetailData();

  createAndPopulateHairSheet(allVendorData);
  createAndPopulateGmDetailSheets(allVendorData);
  createAndPopulateHarwinDetailSheets(allVendorData);

  writeToLog('SYSTEM', 'DETAIL 시트 업데이트 완료 (HAIR, GM, HARWIN 재생성)');
}

/**
 * 전체 업데이트 (MONTHLY + VENDOR + DETAIL)
 */
function updateAllSheets() {
  syncMonthlySummary();
  syncVendorSummary();
  updateAllDetailSheets();

  writeToLog('SYSTEM', '전체 업데이트 완료 (MONTHLY, VENDOR, DETAIL 모두 업데이트)');
}

/**
 * ETC 벤더 목록 디버그
 * INPUT에 있는 모든 벤더와 MONTHLY에 있는 벤더를 비교하여
 * ETC로 분류될 벤더 목록을 보여줍니다.
 */
function debugEtcVendors() {
  Logger.log('\n========== ETC 벤더 목록 디버그 시작 ==========');

  // 1. INPUT 시트에서 모든 벤더 읽기
  const inputData = readAndAggregateInputData();
  if (!inputData) {
    SpreadsheetApp.getUi().alert('ERROR: INPUT 시트를 읽을 수 없습니다.');
    return;
  }

  const { summary, allInputVendors } = inputData;
  Logger.log(`\n1. INPUT 시트의 전체 벤더 (${allInputVendors.size}개):`);
  const inputVendorsList = [...allInputVendors].sort();
  inputVendorsList.forEach((v, idx) => Logger.log(`  [${idx + 1}] ${v}`));

  // 2. MONTHLY 시트에서 벤더 읽기
  const monthlySheet = getSheet(SHEET_NAMES.MONTHLY);
  if (!monthlySheet) {
    SpreadsheetApp.getUi().alert('ERROR: MONTHLY 시트를 찾을 수 없습니다.');
    return;
  }

  const monthlyValues = monthlySheet.getDataRange().getValues();
  const structure = analyzeSheetStructure(monthlyValues);
  const monthlyVendors = new Set(Object.keys(structure.vendors));

  Logger.log(`\n2. MONTHLY 시트의 전체 벤더 (${monthlyVendors.size}개):`);
  const monthlyVendorsList = [...monthlyVendors].sort();
  monthlyVendorsList.forEach((v, idx) => Logger.log(`  [${idx + 1}] ${v}`));

  // 3. ETC 집계 (ETC 상세 시트의 A열 기준)
  const etcData = aggregateEtcData(summary);
  const { etcDetails, etcSummary, etcVendors } = etcData;

  Logger.log(`\n3. ETC로 분류된 벤더 (${etcVendors.length}개):`);
  Logger.log(`   (ETC 상세 시트의 A열에서 읽음)`);
  if (etcVendors.length === 0) {
    Logger.log('  (없음 - ETC 상세 시트에 벤더가 정의되지 않음)');
  } else {
    etcVendors.forEach((v, idx) => {
      // 각 벤더의 총액 계산
      let totalAmount = 0;
      if (etcDetails[v]) {
        for (const year in etcDetails[v]) {
          for (const month in etcDetails[v][year]) {
            totalAmount += etcDetails[v][year][month];
          }
        }
      }
      Logger.log(`  [${idx + 1}] ${v} (총액: $${totalAmount.toFixed(2)})`);
    });
  }

  // 4. ETC 월별 총액
  Logger.log(`\n4. ETC 월별 총액:`);
  if (Object.keys(etcSummary).length === 0) {
    Logger.log('  (없음)');
  } else {
    for (const year in etcSummary) {
      const months = Object.keys(etcSummary[year]).sort((a, b) => parseInt(a) - parseInt(b));
      for (const month of months) {
        const amount = etcSummary[year][month];
        Logger.log(`  ${year}-${String(month).padStart(2, '0')}: $${amount.toFixed(2)}`);
      }
    }
  }

  // 5. UI 알림
  const message = etcVendors.length > 0
    ? `ETC 벤더 디버그 완료!\n\n` +
      `INPUT 벤더: ${allInputVendors.size}개\n` +
      `MONTHLY 벤더: ${monthlyVendors.size}개\n` +
      `ETC 벤더: ${etcVendors.length}개\n` +
      `(ETC 상세 시트의 A열에서 정의)\n\n` +
      `ETC 벤더 목록:\n${etcVendors.join('\n')}\n\n` +
      `자세한 내용은 보기 > 로그를 확인하세요.`
    : `ETC 벤더 디버그 완료!\n\n` +
      `INPUT 벤더: ${allInputVendors.size}개\n` +
      `MONTHLY 벤더: ${monthlyVendors.size}개\n` +
      `ETC 벤더: 0개\n\n` +
      `ETC 상세 시트에 벤더가 정의되지 않았습니다.\n\n` +
      `자세한 내용은 보기 > 로그를 확인하세요.`;

  Logger.log('\n========== ETC 벤더 목록 디버그 완료 ==========');
  SpreadsheetApp.getUi().alert(message);
}
