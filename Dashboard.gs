/**
 * @OnlyCurrentDoc
 *
 * Dashboard.gs
 * 이메일 대시보드 생성 및 전송
 */

/**
 * 이메일 대시보드를 생성하고 전송하는 메인 함수
 */
function sendDashboardEmail() {
  try {
    const today = new Date();
    
    // 1. BUDGET 데이터 읽기 및 분석
    const budgetData = getBudgetData(today);

    // 2. INPUT 데이터 읽기 및 필터링
    const inputData = getFilteredInputData(today);

    // 3. 이메일 HTML 생성
    const emailBody = createEmailBody(today, budgetData, inputData);

    // 4. 이메일 전송
    sendEmail(emailBody);

    writeToLog('Dashboard', '이메일 대시보드를 성공적으로 전송했습니다.');

  } catch (e) {
    // 오류 로깅
    Logger.log(e);
    writeToLog('Dashboard', '오류 발생: ' + e.message);
  }
}

/**
 * BUDGET 시트에서 데이터 가져오기
 * @param {Date} today 현재 날짜
 * @return {object} 예산 관련 데이터
 */
function getBudgetData(today) {
  const budgetSheet = getSheet(SHEET_NAMES.BUDGET);
  if (!budgetSheet) {
    throw new Error('Could not find BUDGET sheet.');
  }

  const data = budgetSheet.getDataRange().getValues();

  // 1. 날짜 헬퍼 생성 (Row 25: Year, C26: Month, C27: Date)
  const yearRow = data[24];  // Row 25 (0-indexed)
  const monthRow = data[25]; // Row 26 (0-indexed)
  const dateRow = data[26]; // Row 27 (0-indexed)
  const dateColumns = [];
  let currentYear = today.getFullYear(); // Fallback
  for (let c = 2; c < monthRow.length; c++) { // Starts from C
    if (yearRow[c]) { // Update current year if a new one is specified
        currentYear = parseInt(yearRow[c], 10);
    }
    if (!isNaN(currentYear) && monthRow[c] && dateRow[c]) {
      dateColumns[c] = new Date(currentYear, monthRow[c] - 1, dateRow[c]);
    }
  }

  // 2. 오늘 날짜에 해당하는 컬럼 찾기
  let todayCol = -1;
  const todayStr = today.toDateString();
  for (let c = 2; c < dateColumns.length; c++) {
    if (dateColumns[c] && dateColumns[c].toDateString() === todayStr) {
      todayCol = c;
      break;
    }
  }

  if (todayCol === -1) {
    writeToLog('getBudgetData', 'Could not find today\'s column in BUDGET sheet.');
    return { dailyBalances: [], upcomingSpendings: [] };
  }
  
  // 3. 주요 행 인덱스 찾기
  const rows = { bank: -1, lease: -1, empTax: -1 };
  for(let i = 0; i < data.length; i++) {
    const header = String(data[i][0]).trim().toUpperCase();
    if (header === 'BANK') rows.bank = i;
    else if (header === 'LEASE') rows.lease = i;
    else if (header === 'EMP TAX') rows.empTax = i;
  }
   if (rows.bank === -1 || rows.lease === -1 || rows.empTax === -1) {
    throw new Error('Could not find required rows (BANK, LEASE, EMP TAX) in BUDGET sheet.');
  }

  // 4. 향후 30일 Bank 잔액 데이터 추출
  const dailyBalances = [];
  const bankRowData = data[rows.bank];
  for (let i = 0; i < 30 && (todayCol + i) < bankRowData.length; i++) {
    dailyBalances.push({
      date: dateColumns[todayCol + i],
      balance: bankRowData[todayCol + i]
    });
  }

  // 5. 향후 14일 Spending >= $2,000 목록 추출
  const upcomingSpendings = [];
   for (let r = rows.lease; r <= rows.empTax; r++) {
    const category = data[r][0];
    const rowData = data[r];
    for (let i = 0; i < 14 && (todayCol + i) < rowData.length; i++) {
      const col = todayCol + i;
      const amount = parseFloat(rowData[col]);
      if (amount && amount >= 2000) {
        upcomingSpendings.push({
          category: category,
          date: dateColumns[col],
          amount: amount
        });
      }
    }
  }

  // 날짜 오름차순 정렬
  upcomingSpendings.sort((a, b) => a.date - b.date);

  return { dailyBalances, upcomingSpendings };
}

/**
 * INPUT 시트에서 데이터 가져오기 (최근 3개월)
 * @param {Date} today 현재 날짜
 * @return {object} 필터링된 데이터
 */
function getFilteredInputData(today) {
  const inputSheet = getSheet(SHEET_NAMES.INPUT);
  if (!inputSheet) {
    throw new Error('Could not find INPUT sheet.');
  }

  const data = inputSheet.getDataRange().getValues().slice(1); // 헤더 제외

  const threeMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 3, 1);

  const deliveries = [];
  const payments = [];

  const col = COLUMN_INDICES.INPUT;

  for (const row of data) {
    const year = parseInt(row[col.YEAR - 1], 10);
    const month = parseInt(row[col.MONTH - 1], 10);

    if (isNaN(year) || isNaN(month)) continue;

    const invoiceDate = new Date(year, month - 1, 1);
    if (invoiceDate < threeMonthsAgo) continue;

    const deliveredValue = String(row[col.DELIVERED - 1]).trim().toUpperCase();
    const paidValue = String(row[col.PAID - 1]).trim().toUpperCase();

    // 배송 완료: DELIVERED == 'O'
    const isDelivered = deliveredValue === 'O';
    // 지불 완료: PAID == 'O'
    const isPaid = paidValue === 'O';

    // 배송 예정 (DELIVERED != 'O')
    if (!isDelivered) {
      deliveries.push({
        vendor: row[col.VENDOR - 1],
        invoice: row[col.INVOICE - 1],
        amount: row[col.AMOUNT - 1],
        delivered: deliveredValue || ''
      });
    }

    // 지불 예정 (PAID != 'O')
    if (!isPaid) {
       const payYear = row[col.PAY_YEAR-1];
       const payMonth = row[col.PAY_MONTH-1];
       const payDay = row[col.PAY_DATE-1];

        payments.push({
            vendor: row[col.VENDOR - 1],
            invoice: row[col.INVOICE - 1],
            amount: row[col.AMOUNT - 1],
            payDate: (payYear && payMonth && payDay) ? new Date(payYear, payMonth - 1, payDay) : null,
            checkNum: row[col.CHECK_NUM - 1],
            paid: paidValue || ''
        });
    }
  }

  // 지불 예정 목록을 날짜순으로 정렬
  payments.sort((a, b) => {
    if (a.payDate && b.payDate) return a.payDate - b.payDate;
    if (a.payDate) return -1;
    if (b.payDate) return 1;
    return 0;
  });

  return {
    deliveries,
    payments
  };
}

/**
 * 이메일 본문 HTML 생성
 * @param {Date} today 현재 날짜
 * @param {object} budgetData 예산 데이터
 * @param {object} inputData INPUT 데이터
 * @return {string} HTML 형식의 이메일 본문
 */
function createEmailBody(today, budgetData, inputData) {
  const { dailyBalances, upcomingSpendings } = budgetData;
  const { deliveries, payments } = inputData;

  // --- Helper to create a table ---
  const createTable = (headers, rows) => {
    let table = '<table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">';
    table += '<thead><tr>';
    headers.forEach(h => table += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;">${h}</th>`);
    table += '</tr></thead>';
    table += '<tbody>';
    rows.forEach(row => {
      table += '<tr>';
      row.forEach(cell => table += `<td style="border: 1px solid #ddd; padding: 8px;">${cell}</td>`);
      table += '</tr>';
    });
    table += '</tbody></table>';
    return table;
  };
  
  // --- 1. Budget Graph (QuickChart API) ---
  let chartHtml = '';

  if (dailyBalances.length > 0) {
    const balances = dailyBalances.map(d => parseFloat(d.balance) || 0);
    const dates = dailyBalances.map(d => d.date.getMonth() + 1 + '/' + d.date.getDate());

    // QuickChart 설정 생성
    const chartConfig = {
      type: 'line',
      data: {
        labels: dates,
        datasets: [{
          label: 'Bank Balance',
          data: balances,
          borderColor: 'rgb(66, 133, 244)',
          backgroundColor: 'rgba(66, 133, 244, 0.1)',
          fill: true,
          tension: 0.4
        }]
      },
      options: {
        title: {
          display: true,
          text: '30-Day Bank Balance Forecast',
          fontSize: 16
        },
        scales: {
          yAxes: [{
            ticks: {
              callback: '(value) => "$" + value.toLocaleString()'
            }
          }]
        },
        legend: {
          display: false
        }
      }
    };

    const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(chartConfig))}&width=700&height=400`;

    chartHtml = `<img src="${chartUrl}" alt="30-Day Bank Balance Chart" style="max-width: 100%; height: auto;">`;
  } else {
    chartHtml = '<p>No balance data available.</p>';
  }

  let html = `
    <h1 style="font-family: Arial, sans-serif; color: #333;">Dashboard for ${today.toLocaleDateString()}</h1>
    <h2 style="font-family: Arial, sans-serif; color: #555;">30-Day Bank Balance Forecast</h2>
    ${chartHtml}
  `;

  // --- 2. Spending >= $2,000 (향후 14일) ---
  html += `<h2 style="font-family: Arial, sans-serif; color: #555;">Upcoming Spendings</h2>`;
  if (upcomingSpendings.length > 0) {
    const spendHeaders = ['Category', 'Date', 'Amount'];
    const spendRows = upcomingSpendings.map(s => [
      s.category, 
      s.date.toLocaleDateString(), 
      `$${s.amount.toFixed(2)}`
    ]);
    html += createTable(spendHeaders, spendRows);
  } else {
    html += '<p>No significant upcoming spendings.</p>';
  }

  // --- 3. DELIVERY 예정 ---
  html += `<h2 style="font-family: Arial, sans-serif; color: #555;">Upcoming Deliveries</h2>`;
  if (deliveries.length > 0) {
    const deliveryHeaders = ['Vendor', 'Invoice #', 'Amount', 'Delivered'];
    const deliveryRows = deliveries.map(d => [d.vendor, d.invoice, d.amount, d.delivered]);
    html += createTable(deliveryHeaders, deliveryRows);
  } else {
    html += '<p>No upcoming deliveries.</p>';
  }

  // --- 4. PAYMENT 예정 ---
  html += `<h2 style="font-family: Arial, sans-serif; color: #555;">Upcoming Payments</h2>`;
  if (payments.length > 0) {
    const paymentHeaders = ['Vendor', 'Invoice #', 'Amount', 'Pay Date', 'Check #'];
    const paymentRows = payments.map(p => [
      p.vendor,
      p.invoice,
      p.amount,
      p.payDate ? p.payDate.toLocaleDateString() : 'N/A',
      p.checkNum
    ]);
    html += createTable(paymentHeaders, paymentRows);
  } else {
    html += '<p>No upcoming payments.</p>';
  }
  
  // --- 5. UPS Tracking (TRACKING 시트에서 읽기) ---
  html += `<h2 style="font-family: Arial, sans-serif; color: #555;">UPS Tracking Status</h2>`;
  const trackingData = getTrackingData();
  if (trackingData.length > 0) {
    const trackingHeaders = ['Tracking #', 'Vendor', 'Status', 'Delivery Date', 'Keep?'];
    const trackingRows = trackingData.map(t => [
      t.trackingNumber,
      t.vendor,
      t.status,
      t.deliveryDate ? t.deliveryDate : 'N/A',
      t.keep
    ]);
    html += createTable(trackingHeaders, trackingRows);
  } else {
    html += '<p>No tracking information available. Run refreshUPSTrackingList() to build the list.</p>';
  }

  return html;
}

/**
 * TRACKING 시트에서 트래킹 데이터 가져오기
 * @return {Array<object>} 트래킹 정보 배열
 */
function getTrackingData() {
  const ss = getActiveSpreadsheet();
  const trackingSheet = ss.getSheetByName('TRACKING');

  if (!trackingSheet) {
    return [];
  }

  const data = trackingSheet.getDataRange().getValues();

  if (data.length <= 1) {
    return []; // 헤더만 있거나 데이터 없음
  }

  const trackingData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    trackingData.push({
      trackingNumber: row[0],
      vendor: row[1],
      status: row[3],
      deliveryDate: row[4],
      keep: row[5]
    });
  }

  return trackingData;
}

/**
 * 이메일 전송
 * @param {string} htmlBody HTML 형식의 이메일 본문
 */
function sendEmail(htmlBody) {
  const recipient = 'lewis.choi.ssc@gmail.com';
  const subject = `일일 대시보드: ${new Date().toLocaleDateString()}`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });
}
