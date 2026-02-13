# Vendor Statement Management System

Google Apps Script 기반 벤더 명세서 관리 및 배송 추적 시스템

## 주요 기능

### 1. 데이터 관리
- **MONTHLY**: 월별 집계 시트
- **VENDOR**: 벤더별 인보이스 추적
- **DETAIL**: HAIR, GM, HARWIN 상세 시트

### 2. AfterShip 배송 추적
- 다중 운송사 자동 감지 (UPS, FedEx, DHL, USPS, Korea Post, CJ대한통운)
- INPUT 시트에서 송장 번호 자동 추출 (N~W열)
- AfterShip API 자동 등록 (누락된 송장)
- 배송 완료 3일 후 자동 삭제
- 일일 이메일 알림 (lewis.choi.ssc@gmail.com)

### 3. 자동화
- 일일 자동 업데이트 스케줄링
- 배송 추적 이메일 발송
- ETC 벤더 자동 집계

## 설정 방법

### AfterShip API 설정

⚠️ **중요**: Google Apps Script는 `.env` 파일을 사용하지 않습니다!

**API 키 설정 방법:**
1. Google Sheet 열기
2. 메뉴: `자동화 > ⚙️ AfterShip API 설정` 클릭
3. API 키 입력 (https://admin.aftership.com/settings/api-keys 에서 발급)
4. 키는 Google Script Properties에 암호화되어 저장됩니다

**또는 스크립트 에디터에서 직접 설정:**
```javascript
PropertiesService.getScriptProperties()
  .setProperty('AFTERSHIP_API_KEY', 'your_api_key_here');
```

### 메뉴 사용법

#### 데이터 업데이트
- `1️⃣ 전체 업데이트`: MONTHLY + VENDOR + DETAIL 모두 업데이트
- `2️⃣ MONTHLY + VENDOR 업데이트`: 월별/벤더별 집계만
- `3️⃣ DETAIL 업데이트`: HAIR, GM, HARWIN 상세 시트만

#### 배송 추적
- `📦 배송 추적 정보 업데이트`: INPUT 시트에서 송장 정보 읽고 AfterShip 업데이트
- `📧 배송 추적 이메일 발송`: 현재 배송 상태 요약 이메일 발송
- `⚙️ AfterShip API 설정`: API 키 입력/변경
- `⏰ 자동 업데이트 설정`: 매일 자동 추적 업데이트 (이메일 미포함)
- `⏰ 자동 업데이트+이메일 설정`: 매일 자동 업데이트 + 이메일 발송

#### 기타
- `🔍 ETC 벤더 목록 확인`: ETC로 분류된 벤더 디버깅

## 날짜 필터링

현재 설정: **2025년 1월 이후 데이터만 처리**

변경 방법: [Common.gs](Common.gs#L28-L32)의 `DATA_FILTER_FROM_DATE` 수정

```javascript
const DATA_FILTER_FROM_DATE = {
  year: 2025,
  month: 1
};
```

## 파일 구조

### Core Files
- [Common.gs](Common.gs) - 공통 설정 및 유틸리티
- [Menu.gs](Menu.gs) - 사용자 메뉴 인터페이스
- [Logger.gs](Logger.gs) - 로깅 시스템

### Data Processing
- [Input_DataReader.gs](Input_DataReader.gs) - INPUT 시트 읽기
- [Monthly_Main.gs](Monthly_Main.gs) - MONTHLY 시트 생성
- [Vendor_Main.gs](Vendor_Main.gs) - VENDOR 시트 생성
- [Vendor_DataReader.gs](Vendor_DataReader.gs) - VENDOR 데이터 읽기

### Detail Sheets
- [Detail_Hair.gs](Detail_Hair.gs) - HAIR 상세 시트
- [Detail_Hair_Data.gs](Detail_Hair_Data.gs) - HAIR 데이터 처리
- [Detail_GM.gs](Detail_GM.gs) - GM 상세 시트
- [Detail_Harwin.gs](Detail_Harwin.gs) - HARWIN 상세 시트
- [Detail_Harwin_Data.gs](Detail_Harwin_Data.gs) - HARWIN 데이터 처리
- [Detail_Styling.gs](Detail_Styling.gs) - 스타일링 유틸리티

### Tracking
- [UPS_Tracking.gs](UPS_Tracking.gs) - AfterShip 멀티 캐리어 추적

## 보안

### API 키 저장
- ✅ Google Script Properties 사용 (암호화, Git 미포함)
- ❌ `.env` 파일 사용 안 함 (Google Apps Script는 Node.js가 아님)
- ❌ 코드에 하드코딩 안 함

### Git 제외 파일
`.gitignore` 파일이 다음을 제외:
- `.env` (참조용)
- `.clasprc.json` (clasp 인증 정보)
- `node_modules/`
- OS/IDE 관련 파일

## 문제 해결

### AfterShip API 401 Error
**증상**: `{"meta":{"type":"Unauthorized","message":"The API key is invalid."}}`

**해결책**:
1. API 키가 설정되어 있는지 확인
2. 메뉴: `자동화 > ⚙️ AfterShip API 설정` 실행
3. 유효한 API 키 입력 (https://admin.aftership.com/settings/api-keys)

### 운송사 자동 감지 실패
**증상**: 잘못된 운송사가 선택됨

**해결책**: [UPS_Tracking.gs](UPS_Tracking.gs#L54-L104)의 `detectCourier()` 함수에서 패턴 확인

지원 운송사:
- UPS: `1Z` + 16자리 (총 18자리)
- FedEx: 12, 15, 20자리 숫자
- DHL: 10-11자리 숫자
- USPS: `94/93/92/91/82/81/80` 시작 또는 `AA123456789US` 형식
- Korea Post: `KR` 종료
- CJ대한통운: 10-12자리 숫자

## 개발 정보

### 최근 변경사항
- 2025-01 날짜 필터 적용
- AfterShip 멀티 캐리어 지원 추가
- 배송 추적 이메일 알림 추가
- ETC 벤더 수동 관리 로직 변경
- Detail 시트 노란색 spacer 행 수정

### Git 커밋 히스토리
```
b8f2528 벤더탭 빨간색 수정. 일단 날짜필터를 2025/1로
1f87f15 셀 병합, etc, 노란셀 문제 수정
ccd5d9a Add date filtering and optimize performance
28b5ae3 FIRST COMMIT
```

## 연락처

배송 추적 알림: lewis.choi.ssc@gmail.com
