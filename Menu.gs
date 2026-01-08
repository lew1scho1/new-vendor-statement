/*
 * @OnlyCurrentDoc
 *
 * Menu.gs
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('자동화')
      .addItem('전체 업데이트', 'updateAllSheets')
      .addItem('MONTHLY 요약 업데이트', 'syncMonthlySummary')
      .addSeparator()
      .addItem('VENDOR 요약 업데이트', 'syncVendorSummary')
      .addSeparator()
      .addItem('DETAIL 전체 업데이트', 'updateAllDetailSheets')
      .addItem('DETAIL HAIR 시트 생성', 'createAndPopulateHairSheet')
      .addItem('DETAIL GM 시트 생성', 'createAndPopulateGmDetailSheets')
      .addItem('DETAIL HARWIN 시트 생성', 'createAndPopulateHarwinDetailSheets')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('DEBUG')
          .addItem('MONTHLY DEBUG 모드', 'debugSyncMonthlySummary')
          .addItem('MONTHLY 구조 확인', 'debugMonthlyStructure')
          .addSeparator()
          .addItem('VENDOR DEBUG 모드', 'debugSyncVendorSummary')
          .addItem('VENDOR 구조 확인', 'debugVendorStructure'))
      .addToUi();
}

function updateAllDetailSheets() {
  createAndPopulateHairSheet();
  createAndPopulateGmDetailSheets();
  createAndPopulateHarwinDetailSheets();
}

function updateAllSheets() {
  syncMonthlySummary();
  syncVendorSummary();
  updateAllDetailSheets();
}
