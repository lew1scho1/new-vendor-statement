/**
 * @OnlyCurrentDoc
 *
 * Vendor_DataReader.gs
 * VENDOR 탭을 위한 INPUT 및 BASIC 시트 데이터 읽기
 */

/**
 * BASIC 탭에서 벤더별 payment method 정보 읽기
 * @return {object} { vendorName: paymentMethod } 형식의 맵
 */
function readBasicPaymentMethods() {
  Logger.log('\n========== Reading BASIC Payment Methods ==========');

  const basicSheet = getSheet(SHEET_NAMES.BASIC);
  if (!basicSheet) {
    Logger.log('ERROR: Could not find BASIC sheet.');
    return {};
  }

  const basicData = basicSheet.getDataRange().getValues().slice(1); // Exclude header
  const paymentMethodMap = {};

  for (let i = 0; i < basicData.length; i++) {
    const row = basicData[i];
    const vendor = normalizeVendorName(row[COLUMN_INDICES.BASIC.VENDOR - 1]);
    const paymentMethod = String(row[COLUMN_INDICES.BASIC.PAYMENT_METHOD - 1] || '').trim().toUpperCase();

    if (vendor && paymentMethod) {
      paymentMethodMap[vendor] = paymentMethod;
      Logger.log(`  ${vendor}: ${paymentMethod}`);
    }
  }

  Logger.log(`Total vendors with payment methods: ${Object.keys(paymentMethodMap).length}`);
  return paymentMethodMap;
}

/**
 * INPUT 탭에서 VENDOR용 인보이스 데이터 읽기
 * @param {object} paymentMethodMap - 벤더별 payment method 맵
 * @return {object} 인보이스 데이터 { vendor: { year: { month: [invoices] } } }
 */
function readVendorInvoicesFromInput(paymentMethodMap) {
  Logger.log('\n========== Reading Vendor Invoices from INPUT ==========');

  const inputSheet = getSheet(SHEET_NAMES.INPUT);
  if (!inputSheet) {
    Logger.log('ERROR: Could not find INPUT sheet.');
    return {};
  }

  const inputData = inputSheet.getDataRange().getValues().slice(1); // Exclude header
  const invoices = {}; // { vendor: { year: { month: [invoice objects] } } }

  let validRowCount = 0;
  let invalidRowCount = 0;
  let filteredCount = 0;

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    const vendor = normalizeVendorName(row[COLUMN_INDICES.INPUT.VENDOR - 1]);
    const year = parseInt(row[COLUMN_INDICES.INPUT.YEAR - 1], 10);
    const month = parseInt(row[COLUMN_INDICES.INPUT.MONTH - 1], 10);
    const amount = parseFloat(row[COLUMN_INDICES.INPUT.AMOUNT - 1]);
    const payYear = parseInt(row[COLUMN_INDICES.INPUT.PAY_YEAR - 1], 10);
    const payMonth = parseInt(row[COLUMN_INDICES.INPUT.PAY_MONTH - 1], 10);
    const payDate = parseInt(row[COLUMN_INDICES.INPUT.PAY_DATE - 1], 10);
    const outstanding = String(row[COLUMN_INDICES.INPUT.OUTSTANDING - 1] || '').trim().toUpperCase();
    const checkNum = row[COLUMN_INDICES.INPUT.CHECK_NUM - 1];

    // 필수 필드 검증
    if (vendor && year && month && !isNaN(amount) && payYear && payMonth && payDate) {
      // 날짜 필터 적용 (Common.gs의 DATA_FILTER_FROM_DATE 기준)
      if (!shouldProcessInvoiceDate(year, month)) {
        filteredCount++;
        continue;
      }
      validRowCount++;

      // 인보이스 객체 생성
      const invoice = {
        amount: amount,
        payYear: payYear,
        payMonth: payMonth,
        payDate: payDate,
        paymentMethod: paymentMethodMap[vendor] || 'UNKNOWN',
        checkNum: checkNum,
        isOutstanding: outstanding === 'O'  // K열이 'O'인지 확인
      };

      // CRITICAL: VENDOR 탭의 year/month 헤더와 매칭하기 위해
      // payYear/payMonth를 키로 사용 (INPUT의 C열 month가 아닌 H열 payMonth!)
      if (!invoices[vendor]) invoices[vendor] = {};
      if (!invoices[vendor][payYear]) invoices[vendor][payYear] = {};
      if (!invoices[vendor][payYear][payMonth]) invoices[vendor][payYear][payMonth] = [];

      invoices[vendor][payYear][payMonth].push(invoice);

      // 첫 5개만 로그 출력
      if (validRowCount <= 5) {
        Logger.log(`Sample row ${i + 2}: ${vendor} PayDate:${payYear}-${payMonth}-${payDate} $${amount} Outstanding:${outstanding}`);
      }
    } else {
      invalidRowCount++;
      if (invalidRowCount <= 5) {
        Logger.log(`Invalid row ${i + 2}: ${vendor} ${year}-${month} Amount:${amount} PayDate:${payMonth}/${payDate}`);
      }
    }
  }

  Logger.log(`\nValid rows: ${validRowCount}, Invalid rows: ${invalidRowCount}`);
  if (filteredCount > 0 && DATA_FILTER_FROM_DATE) {
    Logger.log(`⚡ Filtered out ${filteredCount} rows before ${DATA_FILTER_FROM_DATE.year}/${DATA_FILTER_FROM_DATE.month}`);
  }

  // 각 vendor/year/month별로 payment date 기준으로 정렬
  for (const vendor in invoices) {
    for (const year in invoices[vendor]) {
      for (const month in invoices[vendor][year]) {
        invoices[vendor][year][month].sort((a, b) => {
          // Payment date 기준 정렬 (빠른 순서)
          const dateA = new Date(a.payYear, a.payMonth - 1, a.payDate);
          const dateB = new Date(b.payYear, b.payMonth - 1, b.payDate);
          return dateA - dateB;
        });

        Logger.log(`${vendor} ${year}-${month}: ${invoices[vendor][year][month].length} invoices`);
      }
    }
  }

  return invoices;
}

/**
 * Payment method와 check number를 기반으로 표시 문자열 생성
 * @param {object} invoice - 인보이스 객체
 * @return {string} 표시 문자열 (예: "(cc)", "(ach)", "(#123)")
 */
function formatPaymentMethod(invoice) {
  const method = invoice.paymentMethod;

  if (method === 'CC') {
    return '(cc)';
  } else if (method === 'ACH') {
    return '(ach)';
  } else if (method === 'CHECK') {
    // Check number가 있으면 사용, 없으면 빈 문자열
    const checkNum = String(invoice.checkNum || '').trim();
    if (checkNum && checkNum !== '-') {
      return `(#${checkNum})`;
    }
    return '(check)';
  }

  return '(unknown)';
}

/**
 * 동일한 payment date를 가진 인보이스들을 합침
 * @param {Array} invoices - 인보이스 배열 (이미 정렬됨)
 * @return {Array} Payment date 기준으로 합쳐진 인보이스 배열
 */
function mergeInvoicesBySamePaymentDate(invoices) {
  if (invoices.length === 0) return [];

  const merged = [];
  let current = {
    amount: invoices[0].amount,
    payYear: invoices[0].payYear,
    payMonth: invoices[0].payMonth,
    payDate: invoices[0].payDate,
    paymentMethod: invoices[0].paymentMethod,
    checkNum: invoices[0].checkNum,
    isOutstanding: invoices[0].isOutstanding
  };

  for (let i = 1; i < invoices.length; i++) {
    const inv = invoices[i];

    // 동일한 payment date인지 확인
    if (current.payYear === inv.payYear &&
        current.payMonth === inv.payMonth &&
        current.payDate === inv.payDate) {
      // 동일한 날짜면 amount만 합산
      current.amount += inv.amount;
      // Outstanding 플래그: 하나라도 O면 true
      current.isOutstanding = current.isOutstanding || inv.isOutstanding;
    } else {
      // 다른 날짜면 현재 것을 결과에 추가하고 새로 시작
      merged.push(current);
      current = {
        amount: inv.amount,
        payYear: inv.payYear,
        payMonth: inv.payMonth,
        payDate: inv.payDate,
        paymentMethod: inv.paymentMethod,
        checkNum: inv.checkNum,
        isOutstanding: inv.isOutstanding
      };
    }
  }

  // 마지막 것 추가
  merged.push(current);
  return merged;
}

/**
 * 인보이스를 최대 4개로 제한하고 나머지는 4번째에 합산
 * 특정 벤더(SNG, OUTRE)는 모두 1개로 합산
 * ETC 벤더인 경우 ETC 상세 시트에서 실제 벤더 이름을 찾아 합산
 * @param {Array} invoices - 인보이스 배열 (이미 정렬됨)
 * @param {string} vendorName - 벤더 이름
 * @param {object} etcVendorsData - ETC 상세 시트의 벤더 데이터 (optional)
 * @return {Array} 최대 4개의 인보이스 배열
 */
function limitAndMergeInvoices(invoices, vendorName, etcVendorsData) {
  if (invoices.length === 0) return [];

  // 먼저 동일한 payment date를 가진 것들을 합침
  const mergedByDate = mergeInvoicesBySamePaymentDate(invoices);

  // SNG와 OUTRE는 무조건 1개로 합산
  const alwaysMergeVendors = ['SNG', 'OUTRE'];
  if (alwaysMergeVendors.includes(vendorName.toUpperCase().trim())) {
    const totalAmount = mergedByDate.reduce((sum, inv) => sum + inv.amount, 0);
    const hasAnyOutstanding = mergedByDate.some(inv => inv.isOutstanding);
    return [{
      amount: totalAmount,
      payYear: mergedByDate[0].payYear,
      payMonth: mergedByDate[0].payMonth,
      payDate: mergedByDate[0].payDate,
      paymentMethod: mergedByDate[0].paymentMethod,
      checkNum: mergedByDate[0].checkNum,
      isOutstanding: hasAnyOutstanding
    }];
  }

  // 다른 벤더는 최대 4개까지
  if (mergedByDate.length <= VENDOR_MAX_INVOICES) {
    return mergedByDate;
  }

  // 처음 3개는 그대로 사용
  const result = mergedByDate.slice(0, 3);

  // 4번째부터 끝까지 합산
  const merged = {
    amount: 0,
    payYear: mergedByDate[3].payYear,
    payMonth: mergedByDate[3].payMonth,
    payDate: mergedByDate[3].payDate,
    paymentMethod: mergedByDate[3].paymentMethod,
    checkNum: mergedByDate[3].checkNum,
    isOutstanding: false
  };

  for (let i = 3; i < mergedByDate.length; i++) {
    merged.amount += mergedByDate[i].amount;

    // Outstanding 플래그: 하나라도 O면 true
    merged.isOutstanding = merged.isOutstanding || mergedByDate[i].isOutstanding;

    // 가장 빠른 payment date 사용
    const currentDate = new Date(merged.payYear, merged.payMonth - 1, merged.payDate);
    const invoiceDate = new Date(mergedByDate[i].payYear, mergedByDate[i].payMonth - 1, mergedByDate[i].payDate);

    if (invoiceDate < currentDate) {
      merged.payYear = mergedByDate[i].payYear;
      merged.payMonth = mergedByDate[i].payMonth;
      merged.payDate = mergedByDate[i].payDate;
    }
  }

  result.push(merged);
  return result;
}

/**
 * ETC 상세 시트를 읽어서 ETC에 속한 벤더 목록 가져오기
 * @return {Set<string>} ETC 벤더 이름 Set
 */
function getEtcVendorsFromDetailsSheet() {
  Logger.log('\n========== Reading ETC Details Sheet ==========');

  const ss = getActiveSpreadsheet();
  const etcSheet = ss.getSheetByName(SHEET_NAMES.ETC_DETAILS);

  if (!etcSheet) {
    Logger.log('ETC 상세 시트를 찾을 수 없습니다.');
    return new Set();
  }

  const values = etcSheet.getDataRange().getValues();
  const etcVendors = new Set();

  // 헤더 제외하고 읽기
  for (let i = 1; i < values.length; i++) {
    const vendorName = String(values[i][0] || '').trim();
    if (vendorName) {
      etcVendors.add(vendorName);
      Logger.log(`  ETC 벤더: ${vendorName}`);
    }
  }

  Logger.log(`Total ETC vendors: ${etcVendors.size}`);
  return etcVendors;
}
