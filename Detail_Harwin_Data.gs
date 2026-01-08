/**
 * @OnlyCurrentDoc
 *
 * Detail_Harwin_Data.gs
 */

/**
 * Populate HARWIN detail data into a HARWIN sheet.
 * @param {string} sheetName - Sheet name (e.g., "HARWIN-1")
 * @param {string[]} displayVendorNames - Vendor names to display (with ETC)
 * @param {string[]} originalVendors - Original vendor names from BASIC
 * @param {Set<string>} etcVendorsInHarwin - Set of vendors that are ETC
 * @param {number} unitSize - Number of vendor columns per sheet
 */
function populateHarwinDetailSheetData(sheetName, displayVendorNames, originalVendors, etcVendorsInHarwin, unitSize) {
  Logger.log(`\n========== Populating HARWIN Detail Sheet: ${sheetName} ==========`);

  const sheet = getSheet(sheetName);
  if (!sheet) {
    Logger.log(`ERROR: Sheet not found: ${sheetName}`);
    return;
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_COLUMNS = UNIT_COLUMNS * unitSize;
  const yearBgColor = '#7db3a6';
  const spacerColor = '#fdee09';

  // Read data for all original vendors (including those mapped to ETC)
  const vendorData = readHairDetailData(originalVendors);

  // Build { year: { month: { displayName: [invoices] } } }
  const yearMonthMap = {};
  const yearsSet = new Set();

  // Aggregate data by display name (ETC or original)
  for (const vendorName of originalVendors) {
    if (!vendorName || !vendorData[vendorName]) continue;
    const invoices = vendorData[vendorName];
    const displayName = etcVendorsInHarwin.has(vendorName) ? 'ETC' : vendorName;

    // Only include if this display name is in the current sheet
    if (!displayVendorNames.includes(displayName)) continue;

    for (const invoice of invoices) {
      if (!yearMonthMap[invoice.year]) yearMonthMap[invoice.year] = {};
      if (!yearMonthMap[invoice.year][invoice.month]) yearMonthMap[invoice.year][invoice.month] = {};
      if (!yearMonthMap[invoice.year][invoice.month][displayName]) {
        yearMonthMap[invoice.year][invoice.month][displayName] = [];
      }
      // Add original vendor name to invoice for ETC display
      const invoiceWithVendor = { ...invoice, originalVendor: vendorName };
      yearMonthMap[invoice.year][invoice.month][displayName].push(invoiceWithVendor);
      yearsSet.add(invoice.year);
    }
  }

  const years = Array.from(yearsSet).sort((a, b) => a - b);
  let currentRow = 5;

  // Update Row 4 with first year if we have data
  if (years.length > 0) {
    const yearCell = sheet.getRange(4, 1);
    yearCell.setValue(years[0]);
  }

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
      for (const vendorName of displayVendorNames) {
        if (!vendorName) continue;
        const list = monthData[vendorName] || [];
        if (list.length > maxCount) maxCount = list.length;
      }
      if (maxCount === 0) continue;

      for (let unitIndex = 0; unitIndex < displayVendorNames.length; unitIndex++) {
        const vendorName = displayVendorNames[unitIndex];
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

          // For ETC vendors, show original vendor name in invoice column
          const invoiceDisplay = vendorName === 'ETC' && invoice.originalVendor
            ? `${invoice.originalVendor}`
            : invoice.invoice;

          rows[i] = [dateStr, invoiceDisplay, invoice.amount, payDateStr, source];
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

  Logger.log(`  HARWIN detail sheet populated: ${sheetName}`);
  writeToLog('HARWIN', `HARWIN detail sheet populated: ${sheetName}`);

  applyDetailSheetBorders(sheet, TOTAL_COLUMNS, UNIT_COLUMNS);
}
