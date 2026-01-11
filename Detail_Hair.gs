/**
 * @OnlyCurrentDoc
 *
 * Detail_Hair.gs
 */

/**
 * Creates the HAIR detail sheet layout.
 * @param {Array<string>} vendorNames
 * @param {Array<string>} vendorInfos
 */
function createHairDetailSheet(vendorNames, vendorInfos) {
  Logger.log('========== Creating HAIR Detail Sheet ==========');

  const ss = getActiveSpreadsheet();
  const sheetName = 'HAIR';
  let hairSheet = ss.getSheetByName(sheetName);

  if (!hairSheet) {
    hairSheet = ss.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName}`);
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_UNITS = 4;
  const TOTAL_COLUMNS = UNIT_COLUMNS * TOTAL_UNITS;

  // Clear data area only (rows 5+). Rows 1-3 are user-managed.
  const maxRows = hairSheet.getMaxRows();
  const maxCols = hairSheet.getMaxColumns();
  if (maxRows > 4) {
    hairSheet.getRange(5, 1, maxRows - 4, maxCols).clearContent();
  }

  // Row 4: year row (merged) - will be updated with actual year during data population
  const yearRange = hairSheet.getRange(4, 1, 1, TOTAL_COLUMNS);
  yearRange.breakApart();
  yearRange.merge();
  yearRange.setValue('2025');
  yearRange.setHorizontalAlignment('center');
  yearRange.setVerticalAlignment('middle');
  yearRange.setFontSize(18);
  yearRange.setFontWeight('bold');
  yearRange.setBackground('#7db3a6');

  applyDetailSheetBorders(hairSheet, TOTAL_COLUMNS, UNIT_COLUMNS);

  hairSheet.hideRows(1);
  hairSheet.setFrozenRows(4);

  Logger.log(`HAIR detail sheet created with ${TOTAL_UNITS} units`);
  writeToLog('HAIR', `HAIR detail sheet created (${TOTAL_UNITS} units)`);
}

/**
 * Adds a new HAIR detail unit.
 * @param {number} position
 * @param {string} vendorInfo
 */
function addHairDetailUnit(position, vendorInfo) {
  Logger.log(`========== Adding HAIR Detail Unit at position ${position} ==========`);

  const ss = getActiveSpreadsheet();
  const hairSheet = ss.getSheetByName('HAIR');

  if (!hairSheet) {
    Logger.log('ERROR: HAIR sheet not found');
    return;
  }

  const UNIT_COLUMNS = 5;
  const startCol = (position - 1) * UNIT_COLUMNS + 1;

  hairSheet.insertColumnsAfter(startCol - 1, UNIT_COLUMNS);

  Logger.log(`Added new unit at position ${position}`);
  writeToLog('HAIR', `Added new unit (position: ${position})`);
}

/**
 * Moves a HAIR detail unit.
 * @param {number} fromPosition
 * @param {number} toPosition
 */
function moveHairDetailUnit(fromPosition, toPosition) {
  Logger.log(`========== Moving HAIR Detail Unit from ${fromPosition} to ${toPosition} ==========`);

  const ss = getActiveSpreadsheet();
  const hairSheet = ss.getSheetByName('HAIR');

  if (!hairSheet) {
    Logger.log('ERROR: HAIR sheet not found');
    return;
  }

  const UNIT_COLUMNS = 5;
  const fromStartCol = (fromPosition - 1) * UNIT_COLUMNS + 1;
  const toStartCol = (toPosition - 1) * UNIT_COLUMNS + 1;

  const maxRows = hairSheet.getMaxRows();
  const sourceRange = hairSheet.getRange(1, fromStartCol, maxRows, UNIT_COLUMNS);
  const sourceData = sourceRange.getValues();
  const sourceFormats = sourceRange.getBackgrounds();
  const sourceFontWeights = sourceRange.getFontWeights();

  const targetRange = hairSheet.getRange(1, toStartCol, maxRows, UNIT_COLUMNS);
  targetRange.setValues(sourceData);
  targetRange.setBackgrounds(sourceFormats);
  targetRange.setFontWeights(sourceFontWeights);

  Logger.log(`Moved unit from position ${fromPosition} to ${toPosition}`);
  writeToLog('HAIR', `Moved unit (${fromPosition} -> ${toPosition})`);
}
