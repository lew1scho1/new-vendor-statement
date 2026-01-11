/**
 * @OnlyCurrentDoc
 *
 * Detail_Hair_Data.gs
 */

/**
 * Read ALL vendor detail data from INPUT (optimized - single read).
 * This function reads INPUT and BASIC sheets once and returns data for all vendors.
 * @param {Array<string>} vendorNames - Optional: specific vendors to include (if empty, reads all)
 * @return {object} Map of vendorName -> invoices array
 */
function readAllVendorDetailData(vendorNames) {
  Logger.log('\n========== Reading ALL Vendor Detail Data from INPUT (Optimized) ==========');

  const inputSheet = getSheet(SHEET_NAMES.INPUT);
  const basicSheet = getSheet(SHEET_NAMES.BASIC);

  if (!inputSheet || !basicSheet) {
    Logger.log('ERROR: Could not find INPUT or BASIC sheet.');
    return {};
  }

  // Read BASIC data once
  const basicData = basicSheet.getDataRange().getValues().slice(1);
  const paymentMethodMap = {};

  for (let i = 0; i < basicData.length; i++) {
    const vendor = normalizeVendorName(basicData[i][COLUMN_INDICES.BASIC.VENDOR - 1]);
    const paymentMethod = String(basicData[i][COLUMN_INDICES.BASIC.PAYMENT_METHOD - 1] || '').trim();
    if (vendor && paymentMethod) {
      paymentMethodMap[vendor] = paymentMethod;
    }
  }

  // Read INPUT data once
  const inputData = inputSheet.getDataRange().getValues().slice(1);
  const vendorData = {};

  // Initialize vendor data structure
  if (vendorNames && vendorNames.length > 0) {
    for (const vendorName of vendorNames) {
      if (vendorName) {
        vendorData[vendorName] = [];
      }
    }
  }

  // Process all input data
  let filteredCount = 0;
  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    const vendor = normalizeVendorName(row[COLUMN_INDICES.INPUT.VENDOR - 1]);
    const year = parseInt(row[COLUMN_INDICES.INPUT.YEAR - 1], 10);
    const month = parseInt(row[COLUMN_INDICES.INPUT.MONTH - 1], 10);
    const day = parseInt(row[COLUMN_INDICES.INPUT.DAY - 1], 10);
    const invoice = String(row[COLUMN_INDICES.INPUT.INVOICE - 1] || '').trim();
    const amount = parseFloat(row[COLUMN_INDICES.INPUT.AMOUNT - 1]);
    const payYear = parseInt(row[COLUMN_INDICES.INPUT.PAY_YEAR - 1], 10);
    const payMonth = parseInt(row[COLUMN_INDICES.INPUT.PAY_MONTH - 1], 10);
    const payDate = parseInt(row[COLUMN_INDICES.INPUT.PAY_DATE - 1], 10);
    const checkNum = String(row[COLUMN_INDICES.INPUT.CHECK_NUM - 1] || '').trim();

    if (!vendor) continue;

    // 날짜 필터 적용 (Common.gs의 DATA_FILTER_FROM_DATE 기준)
    if (!shouldProcessInvoiceDate(year, month)) {
      filteredCount++;
      continue;
    }

    // Initialize vendor if not in list (read all mode)
    if (vendorData[vendor] === undefined) {
      vendorData[vendor] = [];
    }

    if (year && month && day && !isNaN(amount) && payYear && payMonth && payDate) {
      vendorData[vendor].push({
        year: year,
        month: month,
        day: day,
        invoice: invoice,
        amount: amount,
        payYear: payYear,
        payMonth: payMonth,
        payDate: payDate,
        checkNum: checkNum,
        paymentMethod: paymentMethodMap[vendor] || ''
      });
    }
  }

  // Sort invoices by date
  const logBuffer = [];
  for (const vendor in vendorData) {
    vendorData[vendor].sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      if (a.month !== b.month) return a.month - b.month;
      return a.day - b.day;
    });
    logBuffer.push(`${vendor}: ${vendorData[vendor].length} invoices`);
  }

  // Batch log output
  if (logBuffer.length > 0) {
    Logger.log(logBuffer.join('\n'));
  }

  Logger.log(`Total vendors with data: ${Object.keys(vendorData).length}`);
  if (filteredCount > 0 && DATA_FILTER_FROM_DATE) {
    Logger.log(`⚡ Filtered out ${filteredCount} invoices before ${DATA_FILTER_FROM_DATE.year}/${DATA_FILTER_FROM_DATE.month}`);
  }

  return vendorData;
}

/**
 * Read HAIR detail data from INPUT.
 * DEPRECATED: Use readAllVendorDetailData() for better performance.
 * This function is kept for backward compatibility.
 * @param {Array<string>} vendorNames
 * @return {object} Map of vendorName -> invoices array
 */
function readHairDetailData(vendorNames) {
  return readAllVendorDetailData(vendorNames);
}

/**
 * Populate HAIR detail data into the HAIR sheet.
 * @param {Array<string>} vendorNames
 * @param {object} [preReadVendorData] - Optional pre-read vendor data (Phase 1 optimization)
 */
function populateHairDetailData(vendorNames, preReadVendorData) {
  Logger.log('\n========== Populating HAIR Detail Sheet ==========');

  const hairSheet = getSheet('HAIR');
  if (!hairSheet) {
    Logger.log('ERROR: HAIR sheet not found');
    return;
  }

  const UNIT_COLUMNS = 5;
  const TOTAL_UNITS = 4;
  const sheetVendorNames = getHairVendorNamesFromSheet(hairSheet, UNIT_COLUMNS, TOTAL_UNITS);
  const vendorData = preReadVendorData || readHairDetailData(sheetVendorNames);
  const TOTAL_COLUMNS = UNIT_COLUMNS * sheetVendorNames.length;
  const yearBgColor = '#7db3a6';
  const spacerColor = '#fdee09';

  // Build { year: { month: { vendorName: [invoices] } } }
  // 기존 배경색 및 병합 초기화 (5행부터 끝까지)
  const maxRows = hairSheet.getMaxRows();
  if (maxRows > 4) {
    const dataRange = hairSheet.getRange(5, 1, maxRows - 4, TOTAL_COLUMNS);
    dataRange.breakApart();
    dataRange.setBackground(null);
  }

  const yearMonthMap = {};
  const yearsSet = new Set();

  for (const vendorName of sheetVendorNames) {
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

  // Update Row 4 with first year if we have data
  if (years.length > 0) {
    const yearCell = hairSheet.getRange(4, 1);
    yearCell.setValue(years[0]);
  }

  for (let y = 0; y < years.length; y++) {
    const year = years[y];

    // Insert year row only on year change (row 4 already has the first year)
    if (y > 0) {
      const yearRange = hairSheet.getRange(currentRow, 1, 1, TOTAL_COLUMNS);
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
      for (const vendorName of sheetVendorNames) {
        if (!vendorName) continue;
        const list = monthData[vendorName] || [];
        if (list.length > maxCount) maxCount = list.length;
      }
      if (maxCount === 0) continue;

      for (let unitIndex = 0; unitIndex < sheetVendorNames.length; unitIndex++) {
        const vendorName = sheetVendorNames[unitIndex];
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

        const range = hairSheet.getRange(currentRow, startCol, maxCount, UNIT_COLUMNS);
        range.setValues(rows);

        // Apply formatting: center alignment for all except AMOUNT (column 3)
        hairSheet.getRange(currentRow, startCol, maxCount, 1).setHorizontalAlignment('center'); // DATE
        hairSheet.getRange(currentRow, startCol + 1, maxCount, 1).setHorizontalAlignment('center'); // INVOICE
        hairSheet.getRange(currentRow, startCol + 2, maxCount, 1).setNumberFormat('$#,##0.00').setHorizontalAlignment('right'); // AMOUNT
        hairSheet.getRange(currentRow, startCol + 3, maxCount, 1).setHorizontalAlignment('center'); // PAY DATE
        hairSheet.getRange(currentRow, startCol + 4, maxCount, 1).setHorizontalAlignment('center'); // SOURCE
      }

      const blockRange = hairSheet.getRange(currentRow, 1, maxCount, TOTAL_COLUMNS);
      blockRange.setBorder(true, true, true, true, true, true, SpreadsheetApp.BorderStyle.SOLID, null);

      currentRow += maxCount;

      // Insert spacer row between months
      // Not after the last month of current year
      const isLastMonth = (m === months.length - 1);

      if (!isLastMonth) {
        const spacerRange = hairSheet.getRange(currentRow, 1, 1, TOTAL_COLUMNS);
        spacerRange.clearContent();
        spacerRange.setBackground(spacerColor);
        currentRow++;
      }
    }
  }

  Logger.log('\nHAIR detail data populated successfully');
  writeToLog('HAIR', 'HAIR detail data populated');

  applyDetailSheetBorders(hairSheet, TOTAL_COLUMNS, UNIT_COLUMNS);
}

/**
 * Read HAIR vendors from BASIC.
 * @return {Array<object>} [{ name, info }, ...]
 */
function getHairVendorsFromBasic() {
  Logger.log('\n========== Reading HAIR Vendors from BASIC ==========');

  const basicSheet = getSheet(SHEET_NAMES.BASIC);
  if (!basicSheet) {
    Logger.log('ERROR: Could not find BASIC sheet.');
    return [];
  }

  const basicData = basicSheet.getDataRange().getValues().slice(1);
  const hairVendors = [];

  for (let i = 0; i < basicData.length; i++) {
    const vendorName = normalizeVendorName(basicData[i][COLUMN_INDICES.BASIC.VENDOR - 1]);
    const category = String(basicData[i][COLUMN_INDICES.BASIC.CATEGORY - 1] || '').trim().toUpperCase();
    const paymentMethod = String(basicData[i][COLUMN_INDICES.BASIC.PAYMENT_METHOD - 1] || '').trim();
    const custNum = String(basicData[i][COLUMN_INDICES.BASIC.CUST_NUM - 1] || '').trim();

    if (category === 'HAIR' && vendorName) {
      const vendorInfo = custNum ? `${vendorName}  ${custNum}` : vendorName;
      hairVendors.push({
        name: vendorName,
        info: vendorInfo
      });
      Logger.log(`  Found HAIR vendor: ${vendorName}`);
    }
  }

  Logger.log(`Total HAIR vendors found: ${hairVendors.length}`);
  return hairVendors;
}

/**
 * Create and populate the HAIR sheet.
 * @param {object} [preReadVendorData] - Optional pre-read vendor data (Phase 1 optimization)
 */
function createAndPopulateHairSheet(preReadVendorData) {
  Logger.log('\n========== Auto Create and Populate HAIR Sheet ==========');

  const hairVendors = getHairVendorsFromBasic();
  if (hairVendors.length === 0) {
    SpreadsheetApp.getUi().alert('BASIC sheet has no HAIR vendors.');
    Logger.log('ERROR: No HAIR vendors found in BASIC sheet.');
    return;
  }

  const maxUnits = 4;
  const vendorNames = [];
  const vendorInfos = [];

  for (let i = 0; i < maxUnits; i++) {
    if (i < hairVendors.length) {
      vendorNames.push(hairVendors[i].name);
      vendorInfos.push(hairVendors[i].info);
    } else {
      vendorNames.push('');
      vendorInfos.push('');
    }
  }

  Logger.log(`Using ${hairVendors.length} HAIR vendors (max 4 units)`);
  Logger.log(`Vendor names: ${vendorNames.join(', ')}`);

  createHairDetailSheet(vendorNames, vendorInfos);
  populateHairDetailData(vendorNames, preReadVendorData);

  writeToLog('HAIR', `HAIR sheet created and populated (${hairVendors.length} vendors)`);
  Logger.log('\nHAIR sheet creation and population completed successfully');
}

/**
 * Reads vendor names from row 1, using unit column width.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} unitColumns
 * @param {number} totalUnits
 * @return {string[]}
 */
function getHairVendorNamesFromSheet(sheet, unitColumns, totalUnits) {
  const names = [];
  for (let unit = 0; unit < totalUnits; unit++) {
    const startCol = unit * unitColumns + 1;
    const name = String(sheet.getRange(1, startCol).getValue() || '').trim();
    names.push(name);
  }
  return names;
}
