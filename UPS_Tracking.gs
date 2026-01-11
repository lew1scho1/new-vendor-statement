/**
 * @OnlyCurrentDoc
 *
 * UPS_Tracking.gs
 * AfterShip API를 사용한 배송 추적
 */

// AfterShip API 설정 (Script Properties에서 관리)
const AFTERSHIP_CONFIG = {
  API_KEY: PropertiesService.getScriptProperties().getProperty('AFTERSHIP_API_KEY'),
  API_URL: 'https://api.aftership.com/v4',
  BATCH_SIZE: 50 // Batch API 최대 크기
};

/**
 * 배송 추적 정보를 업데이트하는 메인 함수
 * 하루 1회 자동 실행 (트리거 설정)
 */
function updateTrackingInfo() {
  try {
    // 1. INPUT 시트에서 배송 완료되지 않은 트래킹 번호 수집
    const trackingNumbers = getActiveTrackingNumbers();

    if (trackingNumbers.length === 0) {
      writeToLog('Tracking', '추적할 트래킹 번호가 없습니다.');
      return;
    }

    // 2. TRACKING 시트 준비
    const trackingSheet = prepareTrackingSheet();

    // 3. Batch API로 트래킹 정보 조회
    const trackingResults = fetchTrackingBatch(trackingNumbers);

    // 4. 결과를 TRACKING 시트에 저장
    saveTrackingResults(trackingSheet, trackingResults);

    // 5. 배송 완료 후 3일 지난 항목 자동 제거
    cleanupOldDeliveries(trackingSheet);

    writeToLog('Tracking', `${trackingResults.length}개 트래킹 정보 업데이트 완료`);

  } catch (e) {
    Logger.log(e);
    writeToLog('Tracking', '오류 발생: ' + e.message);
  }
}

/**
 * 셀 값에서 UPS 트래킹 번호 추출
 * @param {string} cellValue - 셀의 값 (링크 또는 트래킹 번호)
 * @return {string|null} 추출된 트래킹 번호 또는 null
 */
function extractTrackingNumber(cellValue) {
  if (!cellValue || cellValue.length === 0) {
    return null;
  }

  // UPS 링크 패턴: https://www.ups.com/track?tracknum=1Z999AA10123456784
  const linkPattern = /tracknum=([A-Z0-9]+)/i;
  const linkMatch = cellValue.match(linkPattern);
  if (linkMatch) {
    return linkMatch[1];
  }

  // UPS 트래킹 번호 패턴 검증
  // UPS 트래킹 번호는 보통 1Z로 시작하거나 특정 패턴을 따름
  // 1Z + 6자리 + 10자리 = 총 18자리
  const trackingPattern = /\b(1Z[A-Z0-9]{16})\b/i;
  const trackingMatch = cellValue.match(trackingPattern);
  if (trackingMatch) {
    return trackingMatch[1].toUpperCase();
  }

  // 숫자와 대문자만 포함된 10-30자리 문자열 (다른 배송사 트래킹 번호도 포함)
  const generalPattern = /^[A-Z0-9]{10,30}$/i;
  if (generalPattern.test(cellValue)) {
    return cellValue.toUpperCase();
  }

  return null;
}

/**
 * INPUT 시트에서 배송 완료되지 않은 트래킹 번호 수집
 * @return {Array<string>} 트래킹 번호 배열
 */
function getActiveTrackingNumbers() {
  const inputSheet = getSheet(SHEET_NAMES.INPUT);
  if (!inputSheet) {
    throw new Error('Could not find INPUT sheet.');
  }

  const data = inputSheet.getDataRange().getValues().slice(1); // 헤더 제외
  const col = COLUMN_INDICES.INPUT;
  const trackingNumbers = new Set();
  const today = new Date();
  const threeMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 3, 1);

  for (const row of data) {
    const year = parseInt(row[col.YEAR - 1], 10);
    const month = parseInt(row[col.MONTH - 1], 10);

    if (isNaN(year) || isNaN(month)) continue;

    const invoiceDate = new Date(year, month - 1, 1);
    if (invoiceDate < threeMonthsAgo) continue;

    // 배송이 필요한 항목만 (DELIVERED == 'X')
    const deliveredValue = String(row[col.DELIVERED - 1]).trim().toUpperCase();
    if (deliveredValue !== 'X') continue;

    // UPS 트래킹 번호 수집 (N열 ~ W열)
    for (let i = col.UPS_TRACKING_START - 1; i < col.UPS_TRACKING_END; i++) {
      const cellValue = String(row[i]).trim();
      if (cellValue && cellValue.length > 0) {
        // UPS 링크에서 트래킹 번호 추출 또는 직접 입력된 번호 사용
        const trackingNumber = extractTrackingNumber(cellValue);
        if (trackingNumber) {
          trackingNumbers.add(trackingNumber);
        }
      }
    }
  }

  return Array.from(trackingNumbers);
}

/**
 * TRACKING 시트 준비 (없으면 생성)
 * @return {Sheet} TRACKING 시트
 */
function prepareTrackingSheet() {
  const ss = getActiveSpreadsheet();
  let trackingSheet = ss.getSheetByName('TRACKING');

  if (!trackingSheet) {
    trackingSheet = ss.insertSheet('TRACKING');
    // 헤더 추가
    const headers = ['Tracking Number', 'Courier', 'Status', 'Location', 'Estimated Delivery', 'Last Update', 'Delivered Date'];
    trackingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    trackingSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    trackingSheet.setFrozenRows(1);
  }

  return trackingSheet;
}

/**
 * AfterShip Batch API로 트래킹 정보 조회
 * @param {Array<string>} trackingNumbers - 트래킹 번호 배열
 * @return {Array<object>} 트래킹 결과 배열
 */
function fetchTrackingBatch(trackingNumbers) {
  const apiKey = AFTERSHIP_CONFIG.API_KEY;

  // API 키가 설정되지 않은 경우 더미 데이터 반환
  if (!apiKey) {
    writeToLog('Tracking', 'AfterShip API Key가 설정되지 않았습니다. setupAfterShipCredentials()를 실행하세요.');
    return trackingNumbers.map(num => ({
      trackingNumber: num,
      courier: 'ups',
      status: 'Pending',
      location: 'Setup API Key',
      estimatedDelivery: null,
      lastUpdate: new Date(),
      deliveredDate: null
    }));
  }

  const results = [];

  // Batch 크기로 나누어 처리
  for (let i = 0; i < trackingNumbers.length; i += AFTERSHIP_CONFIG.BATCH_SIZE) {
    const batch = trackingNumbers.slice(i, i + AFTERSHIP_CONFIG.BATCH_SIZE);

    try {
      const batchResults = fetchAfterShipBatch(batch, apiKey);
      results.push(...batchResults);

      // Rate limit 방지
      if (i + AFTERSHIP_CONFIG.BATCH_SIZE < trackingNumbers.length) {
        Utilities.sleep(1000);
      }

    } catch (e) {
      Logger.log(`Batch ${i} error: ${e.message}`);
      // 에러 발생 시 해당 배치는 에러 상태로 추가
      batch.forEach(num => {
        results.push({
          trackingNumber: num,
          courier: 'ups',
          status: 'ERROR',
          location: e.message,
          estimatedDelivery: null,
          lastUpdate: new Date(),
          deliveredDate: null
        });
      });
    }
  }

  return results;
}

/**
 * AfterShip API로 단일 배치 조회
 * @param {Array<string>} trackingNumbers - 트래킹 번호 배열
 * @param {string} apiKey - AfterShip API Key
 * @return {Array<object>} 트래킹 결과
 */
function fetchAfterShipBatch(trackingNumbers, apiKey) {
  const results = [];

  // AfterShip은 개별 조회만 지원하므로 각 트래킹 번호를 순차 조회
  for (const trackingNum of trackingNumbers) {
    try {
      const result = fetchSingleTracking(trackingNum, apiKey);
      if (result) {
        results.push(result);
      }

      // Rate limit 방지 (초당 10회 제한)
      Utilities.sleep(100);

    } catch (e) {
      Logger.log(`Tracking ${trackingNum} error: ${e.message}`);
      results.push({
        trackingNumber: trackingNum,
        courier: 'ups',
        status: 'ERROR',
        location: e.message.substring(0, 100),
        estimatedDelivery: null,
        lastUpdate: new Date(),
        deliveredDate: null
      });
    }
  }

  return results;
}

/**
 * AfterShip API로 단일 트래킹 조회
 * @param {string} trackingNumber - 트래킹 번호
 * @param {string} apiKey - AfterShip API Key
 * @return {object} 트래킹 결과
 */
function fetchSingleTracking(trackingNumber, apiKey) {
  // URL 인코딩된 트래킹 번호
  const encodedTracking = encodeURIComponent(trackingNumber);
  const url = `${AFTERSHIP_CONFIG.API_URL}/trackings/ups/${encodedTracking}`;

  const options = {
    method: 'get',
    headers: {
      'Content-Type': 'application/json',
      'aftership-api-key': apiKey
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode === 404) {
    // 트래킹 정보가 없으면 등록 시도
    return createAndFetchTracking(trackingNumber, apiKey);
  }

  if (responseCode !== 200) {
    throw new Error(`API Error: ${responseCode} - ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());

  if (data.data && data.data.tracking) {
    return parseTrackingData(data.data.tracking);
  }

  return null;
}

/**
 * AfterShip에 트래킹 등록 후 조회
 * @param {string} trackingNumber - 트래킹 번호
 * @param {string} apiKey - AfterShip API Key
 * @return {object} 트래킹 결과
 */
function createAndFetchTracking(trackingNumber, apiKey) {
  // 1. 트래킹 등록
  const createUrl = `${AFTERSHIP_CONFIG.API_URL}/trackings`;
  const createPayload = {
    tracking: {
      tracking_number: trackingNumber,
      slug: 'ups'
    }
  };

  const createOptions = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'aftership-api-key': apiKey
    },
    payload: JSON.stringify(createPayload),
    muteHttpExceptions: true
  };

  const createResponse = UrlFetchApp.fetch(createUrl, createOptions);
  const createCode = createResponse.getResponseCode();

  if (createCode !== 201 && createCode !== 200) {
    throw new Error(`Create failed: ${createCode} - ${createResponse.getContentText()}`);
  }

  const createData = JSON.parse(createResponse.getContentText());

  if (createData.data && createData.data.tracking) {
    return parseTrackingData(createData.data.tracking);
  }

  return null;
}

/**
 * AfterShip 트래킹 데이터 파싱
 * @param {object} tracking - AfterShip tracking 객체
 * @return {object} 파싱된 트래킹 정보
 */
function parseTrackingData(tracking) {
  const lastCheckpoint = tracking.checkpoints && tracking.checkpoints.length > 0
    ? tracking.checkpoints[tracking.checkpoints.length - 1]
    : null;

  const location = lastCheckpoint
    ? `${lastCheckpoint.city || ''}, ${lastCheckpoint.state || ''}`.trim().replace(/^,\s*/, '')
    : '';

  // 배송 완료 여부 확인
  const isDelivered = tracking.tag === 'Delivered';
  const deliveredDate = isDelivered && lastCheckpoint
    ? new Date(lastCheckpoint.checkpoint_time)
    : null;

  return {
    trackingNumber: tracking.tracking_number,
    courier: tracking.slug,
    status: tracking.tag || 'Unknown',
    location: location || 'N/A',
    estimatedDelivery: tracking.expected_delivery ? new Date(tracking.expected_delivery) : null,
    lastUpdate: new Date(),
    deliveredDate: deliveredDate
  };
}


/**
 * 트래킹 결과를 TRACKING 시트에 저장
 * @param {Sheet} sheet - TRACKING 시트
 * @param {Array<object>} results - 트래킹 결과 배열
 */
function saveTrackingResults(sheet, results) {
  // 기존 데이터 읽기
  const existingData = sheet.getDataRange().getValues();
  const existingMap = new Map();

  // 헤더 제외하고 기존 데이터를 Map으로 저장
  for (let i = 1; i < existingData.length; i++) {
    const trackingNum = existingData[i][0];
    existingMap.set(trackingNum, existingData[i]);
  }

  // 새 데이터로 업데이트
  results.forEach(r => {
    existingMap.set(r.trackingNumber, [
      r.trackingNumber,
      r.courier,
      r.status,
      r.location,
      r.estimatedDelivery || '',
      r.lastUpdate,
      r.deliveredDate || ''
    ]);
  });

  // 시트 초기화 후 새 데이터 작성
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const rows = Array.from(existingMap.values());
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);

    // 날짜 포맷 설정
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/**
 * 배송 완료 후 3일 지난 항목 자동 제거
 * @param {Sheet} sheet - TRACKING 시트
 */
function cleanupOldDeliveries(sheet) {
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const threeDaysAgo = new Date(today.getTime() - (3 * 24 * 60 * 60 * 1000));

  const rowsToKeep = [data[0]]; // 헤더 유지

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[2]; // Status 컬럼
    const deliveredDate = row[6]; // Delivered Date 컬럼

    // 배송 완료되지 않았거나, 배송 완료 후 3일 이내인 경우 유지
    if (status !== 'Delivered' || !deliveredDate || new Date(deliveredDate) >= threeDaysAgo) {
      rowsToKeep.push(row);
    }
  }

  // 데이터 다시 작성
  sheet.clearContents();
  if (rowsToKeep.length > 0) {
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

    // 헤더 포맷
    sheet.getRange(1, 1, 1, rowsToKeep[0].length)
      .setFontWeight('bold')
      .setBackground('#e8f0fe');

    // 날짜 포맷
    if (rowsToKeep.length > 1) {
      sheet.getRange(2, 5, rowsToKeep.length - 1, 1).setNumberFormat('yyyy-mm-dd');
      sheet.getRange(2, 6, rowsToKeep.length - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      sheet.getRange(2, 7, rowsToKeep.length - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
  }

  const removedCount = data.length - rowsToKeep.length;
  if (removedCount > 0) {
    writeToLog('Tracking', `배송 완료 후 3일 지난 항목 ${removedCount}개 제거됨`);
  }
}

/**
 * AfterShip API 설정을 Script Properties에 저장
 * 최초 1회 실행 필요
 */
function setupAfterShipCredentials() {
  const ui = SpreadsheetApp.getUi();

  const apiKey = ui.prompt(
    'AfterShip API Key 입력',
    'AfterShip API Key를 입력하세요:\n(https://admin.aftership.com/settings/api-keys)',
    ui.ButtonSet.OK_CANCEL
  );

  if (apiKey.getSelectedButton() === ui.Button.OK) {
    const key = apiKey.getResponseText().trim();

    if (key) {
      PropertiesService.getScriptProperties().setProperty('AFTERSHIP_API_KEY', key);
      ui.alert('성공', 'AfterShip API Key가 저장되었습니다.\n\nupdateTrackingInfo() 함수를 실행하여 트래킹 정보를 가져올 수 있습니다.', ui.ButtonSet.OK);
    } else {
      ui.alert('오류', 'API Key를 입력해주세요.', ui.ButtonSet.OK);
    }
  }
}

/**
 * 트리거 설정 (하루 1회 자동 실행)
 * 최초 1회 실행 필요
 */
function setupDailyTrigger() {
  // 기존 트리거 제거
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateTrackingInfo') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 새 트리거 생성 (매일 오전 8시~9시 사이)
  ScriptApp.newTrigger('updateTrackingInfo')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('트리거 설정 완료', '매일 오전 8시~9시에 자동으로 트래킹 정보가 업데이트됩니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}
