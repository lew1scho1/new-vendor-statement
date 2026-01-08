/**
 * @OnlyCurrentDoc
 *
 * Detail_GM_Data.gs
 */

/**
 * Populate GM detail data into a GM sheet.
 * @param {string} sheetName
 * @param {string[]} vendorNames
 * @param {number} unitSize
 * @param {object} [preReadVendorData] - Optional pre-read vendor data (Phase 1 optimization)
 */
function populateGmDetailSheetData(sheetName, vendorNames, unitSize, preReadVendorData) {
  Logger.log(`\n========== Populating GM Detail Sheet: ${sheetName} ==========`);

  const sheet = getSheet(sheetName);
  if (!sheet) {
    Logger.log(`ERROR: Sheet not found: ${sheetName}`);
    return;
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_COLUMNS = UNIT_COLUMNS * unitSize;
  const yearBgColor = '#7db3a6';
  const spacerColor = '#fdee09';

  const vendorData = preReadVendorData || readHairDetailData(vendorNames);

  // Build { year: { month: { vendorName: [invoices] } } }
  const yearMonthMap = {};
  const yearsSet = new Set();

  for (const vendorName of vendorNames) {
    if (!vendorName || !vendorData[vendorName]) continue;
    const invoices = vendorData[vendorName];
    for (const invoice of invoices) {
      if (!yearMonthMap[invoice.year]) yearMonthMap[invoice.year] = {};
      if (!yearMonthMap[invoice.year][invoice.month]) yearMonthMap[invoice.year][invoice.month] = {};
      if (!yearMonthMap[invoice.year][invoice.month][vendorName]) {
        yearMonthMap[invoice.year][invoice.month][vendorName] = [];
      }
      yearMonthMap[invoice.year][invoice.month][vendorName].push(invoice);
      yearsSet.add(invoice.year);
    }
  }

  const years = Array.from(yearsSet).sort((a, b) => a - b);
  let currentRow = 5;

  for (let y = 0; y < years.length; y++) {
    const year = years[y];

    // Insert year row only on year change (row 4 already has the first year)
    if (y > 0) {
      const yearRange = sheet.getRange(currentRow, 1, 1, TOTAL_COLUMNS);
      yearRange.breakApart();
      yearRange.merge();
      yearRange.setValue(year);
      yearRange.setHorizontalAlignment('center');
      yearRange.setVerticalAlignment('middle');
      yearRange.setFontSize(18);
      yearRange.setFontWeight('bold');
      yearRange.setBackground(yearBgColor);
      yearRange.setBorder(true, true, true, true, false, false, SpreadsheetApp.BorderStyle.SOLID, null);
      currentRow++;
    }

    const months = Object.keys(yearMonthMap[year] || {}).map(Number).sort((a, b) => a - b);
    for (let m = 0; m < months.length; m++) {
      const month = months[m];
      const monthData = yearMonthMap[year][month] || {};

      let maxCount = 0;
      for (const vendorName of vendorNames) {
        if (!vendorName) continue;
        const list = monthData[vendorName] || [];
        if (list.length > maxCount) maxCount = list.length;
      }
      if (maxCount === 0) continue;

      for (let unitIndex = 0; unitIndex < vendorNames.length; unitIndex++) {
        const vendorName = vendorNames[unitIndex];
        const invoices = monthData[vendorName] || [];
        const startCol = unitIndex * UNIT_COLUMNS + 1;

        const rows = Array.from({ length: maxCount }, () => ['', '', '', '', '']);
        for (let i = 0; i < invoices.length; i++) {
          const invoice = invoices[i];
          const dateStr = `${invoice.month}/${invoice.day}`;
          const payDateStr = `${invoice.payMonth}/${invoice.payDate}`;
          const source = invoice.checkNum && invoice.checkNum !== '-' && invoice.checkNum !== ''
            ? `#${invoice.checkNum}`
            : invoice.paymentMethod;
          rows[i] = [dateStr, invoice.invoice, invoice.amount, payDateStr, source];
        }

        const range = sheet.getRange(currentRow, startCol, maxCount, UNIT_COLUMNS);
        range.setValues(rows);

        // Apply formatting: center alignment for all except AMOUNT (column 3)
        sheet.getRange(currentRow, startCol, maxCount, 1).setHorizontalAlignment('center'); // DATE
        sheet.getRange(currentRow, startCol + 1, maxCount, 1).setHorizontalAlignment('center'); // INVOICE
        sheet.getRange(currentRow, startCol + 2, maxCount, 1).setNumberFormat('$#,##0.00').setHorizontalAlignment('right'); // AMOUNT
        sheet.getRange(currentRow, startCol + 3, maxCount, 1).setHorizontalAlignment('center'); // PAY DATE
        sheet.getRange(currentRow, startCol + 4, maxCount, 1).setHorizontalAlignment('center'); // SOURCE
      }

      const blockRange = sheet.getRange(currentRow, 1, maxCount, TOTAL_COLUMNS);
      blockRange.setBorder(true, true, true, true, true, true, SpreadsheetApp.BorderStyle.SOLID, null);

      currentRow += maxCount;

      // Insert spacer row between months (not after the last month of the year)
      if (m < months.length - 1) {
        const spacerRange = sheet.getRange(currentRow, 1, 1, TOTAL_COLUMNS);
        spacerRange.clearContent();
        spacerRange.setBackground(spacerColor);
        currentRow++;
      }
    }
  }

  Logger.log(`GM detail sheet populated: ${sheetName}`);
  writeToLog('GM', `GM detail sheet populated: ${sheetName}`);

  applyDetailSheetBorders(sheet, TOTAL_COLUMNS, UNIT_COLUMNS);
}
