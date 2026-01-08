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
      .addToUi();
}

/**
 * MONTHLY + VENDOR 업데이트 (8월 이후 데이터만)
 * 8월 이전 데이터는 기존 시트에 보존됩니다.
 */
function updateMontlyAndVendor() {
  syncMonthlySummary();
  syncVendorSummary();

  writeToLog('SYSTEM', 'MONTHLY + VENDOR 업데이트 완료 (8월 이전 데이터 보존)');
}

/**
 * DETAIL 시트 전체 업데이트
 * 모든 데이터(8월 이전 포함)를 다시 생성합니다.
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
