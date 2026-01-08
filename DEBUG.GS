/**
 * Logs the number of rows between each "MONTH" and the next "SUBTOTAL" in column A.
 * Looks for exact match (trimmed) in column A.
 */
function debugCountRowsBetweenMonthSubtotal() {
  const SHEET_NAME = 'MONTHLY';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet not found: ${SHEET_NAME}`);
    return;
  }

  const values = sheet.getDataRange().getValues();
  const monthRows = [];
  const subtotalRows = [];

  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0] || '').trim();
    if (cell === 'MONTH') monthRows.push(i);
    if (cell === 'SUBTOTAL') subtotalRows.push(i);
  }

  Logger.log(`MONTH rows (1-based): ${monthRows.map(r => r + 1).join(', ')}`);
  Logger.log(`SUBTOTAL rows (1-based): ${subtotalRows.map(r => r + 1).join(', ')}`);

  // Pair each MONTH with the next SUBTOTAL after it
  for (let m = 0; m < monthRows.length; m++) {
    const monthRow = monthRows[m];
    const nextSubtotal = subtotalRows.find(r => r > monthRow);
    if (nextSubtotal === undefined) {
      Logger.log(`MONTH at row ${monthRow + 1}: no following SUBTOTAL found`);
      continue;
    }
    const countBetween = Math.max(0, nextSubtotal - monthRow - 1);
    Logger.log(
      `MONTH at row ${monthRow + 1} -> SUBTOTAL at row ${nextSubtotal + 1}: ${countBetween} row(s) between`
    );
  }
}
