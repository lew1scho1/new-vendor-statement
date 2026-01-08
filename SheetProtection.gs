/**
 * @OnlyCurrentDoc
 *
 * SheetProtection.gs
 * 보호된 행과 열의 백업/복원 관련 함수들
 */

/**
 * 보호된 행(SUBTOTAL, GM GRAND TOTAL, MONTH, GRAND TOTAL)을 백업
 * @param {Sheet} sheet - MONTHLY 시트
 * @param {any[][]} values - MONTHLY 시트의 모든 값
 * @return {object} 백업된 데이터 { protectedRows, protectedFormulas, protectedValues }
 */
function backupProtectedRows(sheet, values) {
  Logger.log('\n========== Backing Up Protected Rows ==========');

  const protectedRows = [];
  const protectedFormulas = [];
  const protectedValues = [];

  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0] || '').trim().toUpperCase();
    const normalizedCell = cell.replace(/\s+/g, ' ');

    if (PROTECTED_LABELS.some(label => {
      const normalizedLabel = label.toUpperCase().replace(/\s+/g, ' ');
      return cell === normalizedLabel || normalizedCell === normalizedLabel;
    })) {
      protectedRows.push(i);
      // Get both formulas and values for this entire row
      const rowFormulas = sheet.getRange(i + 1, 1, 1, values[i].length).getFormulas()[0];
      const rowValues = sheet.getRange(i + 1, 1, 1, values[i].length).getValues()[0];
      protectedFormulas.push(rowFormulas);
      protectedValues.push(rowValues);
      Logger.log(`📋 Backed up protected row ${i + 1}: "${values[i][0]}"`);
    }
  }

  Logger.log(`Total protected rows backed up: ${protectedRows.length}`);

  return {
    protectedRows,
    protectedFormulas,
    protectedValues
  };
}

/**
 * 보호된 열(O=15, AB=28)의 수식을 백업
 * @param {Sheet} sheet - MONTHLY 시트
 * @param {any[][]} values - MONTHLY 시트의 모든 값
 * @return {Array<object>} 백업된 수식 배열 (각 행마다 { col: formula } 객체)
 */
function backupProtectedColumnFormulas(sheet, values) {
  Logger.log('\n========== Backing Up Protected Column Formulas ==========');

  const protectedColumnFormulas = []; // Store formulas for protected columns

  for (let i = 0; i < values.length; i++) {
    const rowFormulas = {};
    for (const col of PROTECTED_COLUMNS) {
      const formula = sheet.getRange(i + 1, col).getFormula();
      if (formula) {
        rowFormulas[col] = formula;
      }
    }
    protectedColumnFormulas.push(rowFormulas);
  }

  Logger.log(`📋 Backed up formulas for protected columns (O, AB) across ${values.length} rows`);

  return protectedColumnFormulas;
}

/**
 * 보호된 행의 수식과 값을 복원
 * @param {Sheet} sheet - MONTHLY 시트
 * @param {number[]} protectedRows - 보호된 행 인덱스 배열 (0-based)
 * @param {string[][]} protectedFormulas - 보호된 행의 수식 배열
 * @param {any[][]} protectedValues - 보호된 행의 값 배열
 */
function restoreProtectedRows(sheet, protectedRows, protectedFormulas, protectedValues) {
  Logger.log('\n========== Restoring Protected Rows ==========');

  for (let i = 0; i < protectedRows.length; i++) {
    const rowIndex = protectedRows[i];
    const formulas = protectedFormulas[i];
    const values = protectedValues[i];

    if (formulas && formulas.length > 0) {
      // First restore values (including A column labels)
      sheet.getRange(rowIndex + 1, 1, 1, values.length).setValues([values]);

      // Then restore formulas (this will overwrite cells that have formulas)
      for (let col = 0; col < formulas.length; col++) {
        if (formulas[col]) { // Only set if there's a formula
          sheet.getRange(rowIndex + 1, col + 1).setFormula(formulas[col]);
        }
      }

      Logger.log(`✅ Restored protected row ${rowIndex + 1}`);
    }
  }

  Logger.log(`Total protected rows restored: ${protectedRows.length}`);
}

/**
 * 보호된 열(O=15, AB=28)의 수식을 복원
 * @param {Sheet} sheet - MONTHLY 시트
 * @param {Array<object>} protectedColumnFormulas - 백업된 수식 배열
 */
function restoreProtectedColumnFormulas(sheet, protectedColumnFormulas) {
  Logger.log('\n========== Restoring Protected Column Formulas ==========');

  for (let i = 0; i < protectedColumnFormulas.length; i++) {
    const rowFormulas = protectedColumnFormulas[i];
    for (const col in rowFormulas) {
      const formula = rowFormulas[col];
      if (formula) {
        sheet.getRange(i + 1, parseInt(col)).setFormula(formula);
      }
    }
  }

  Logger.log(`✅ Restored formulas for protected columns (O, AB) across all rows`);
}

/**
 * 보호된 행 인덱스를 Set으로 반환
 * @param {number[]} protectedRows - 보호된 행 인덱스 배열 (0-based)
 * @return {Set<number>} 보호된 행 인덱스 Set
 */
function getProtectedRowIndices(protectedRows) {
  return new Set(protectedRows);
}

/**
 * 특정 행이 보호된 행인지 확인
 * @param {number} rowIndex - 확인할 행 인덱스 (0-based)
 * @param {Set<number>} protectedRowIndices - 보호된 행 인덱스 Set
 * @return {boolean} 보호된 행이면 true
 */
function isProtectedRow(rowIndex, protectedRowIndices) {
  return protectedRowIndices.has(rowIndex);
}

/**
 * 벤더 이름이 보호된 라벨인지 확인
 * @param {string} vendorName - 확인할 벤더 이름
 * @return {boolean} 보호된 라벨이면 true
 */
function isProtectedLabel(vendorName) {
  const upperVendorName = vendorName.toUpperCase().trim();
  return PROTECTED_LABELS.some(label => upperVendorName.includes(label));
}
