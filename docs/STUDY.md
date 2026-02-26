# 영림발주서 자동화 시스템 - 학습 자료

> **작성일**: 2026-01-29  
> **프로젝트**: Google Apps Script 기반 영림발주서 자동화  
> **목적**: 프로젝트 구조, 작동 원리, 유지보수 가이드 제공

---

## 📋 목차
1. [프로젝트 개요](#프로젝트-개요)
2. [시스템 아키텍처](#시스템-아키텍처)
3. [핵심 기능 분석](#핵심-기능-분석)
4. [오늘의 디버깅 세션 (2026-01-29)](#오늘의-디버깅-세션-2026-01-29)
5. [주요 수정 내역](#주요-수정-내역)
6. [주의사항 및 유지보수 가이드](#주의사항-및-유지보수-가이드)

---

## 프로젝트 개요

### 목적
Google Sheets 기반의 **영림발주서** 시트에서 제품 정보를 입력하면, 자동으로 단가를 계산하고 품목 코드를 생성하는 자동화 시스템입니다.

### 기술 스택
- **플랫폼**: Google Apps Script (JavaScript 기반)
- **배포 도구**: `clasp` (Command Line Apps Script Projects)
- **주요 시트**:
  - `영림발주서`: 메인 작업 시트
  - `테스트`: 추가 금액 및 인필 방향 참조 데이터
  - `영림문틀단가표`: 제품 단가 참조 데이터
  - `필터`: 드롭다운 목록 소스 데이터

### 프로젝트 구조
```
이비즈/
├── door_v4.gs          # 현재 활성 스크립트 (유일한 배포 대상)
├── door_v0~v3.gs       # 레거시 버전 (배포 제외)
├── appsscript.json     # Apps Script 매니페스트
├── .clasp.json         # Clasp 설정
├── .claspignore        # 배포 제외 파일 목록
└── docs/
    ├── PROJECT.md
    ├── SESSION.md
    ├── DECISIONS.md
    └── STUDY.md        # 본 문서
```

---

## 시스템 아키텍처

### 데이터 흐름도
```
[사용자 입력]
    ↓
[onEdit 트리거] → 색상(AW), 높이(AV) 감지
    ↓
[자동 완성 로직]
    ├─→ AX열: 색상 코드 매핑
    ├─→ BA열: 가스켓 정보 (12-20행)
    └─→ AQ열: 인필 방향 (22-35행)
    
[메뉴 실행: 전체(단가+코드)]
    ↓
[계산_영림발주서_가격_내부] → A열에 단가 계산
    ↓
[생성_품목코드_문틀_내부] → BC~BE열에 코드/명칭 생성
```

### 핵심 CONFIG 객체
```javascript
const CONFIG = {
  SHEET_NAME: "영림발주서",
  TEST_SHEET_NAME: "테스트",
  START_ROW: 12,
  END_ROW: 35,
  COLS: {
    AP: 42,  // 제품명
    AQ: 43,  // 인필 방향 (Y/N)
    AR: 44,  // 수량
    AS: 45,  // 규격
    AT: 46,  // 너비
    AV: 48,  // 높이
    AW: 49,  // 색상
    AX: 50,  // 색상코드
    BA: 53   // 가스켓
  }
};
```

---

## 핵심 기능 분석

### 1. onEdit 트리거 (자동 완성)
**실행 조건**: 사용자가 `AW`(색상) 또는 `AV`(높이) 열을 수정할 때

**처리 로직**:
1. **색상 코드 자동 완성 (AX열)**
   - `COLOR_MAP` 속성에서 색상명 → 색상코드 매핑
   - 완전 일치 → 공백 제거 일치 → 부분 일치 순으로 검색

2. **가스켓 정보 자동 완성 (BA열, 12-20행)**
   - `GASKET_COLOR_MAP` 속성에서 색상/색상코드로 검색
   - M/P/S 그룹별로 순차 탐색
   - 단종 여부에 따라 상태 또는 가스켓 색상 표시

3. **인필 방향 자동 완성 (AQ열, 22-35행)**
   - **조건**: 높이(AV) >= 2166 AND 색상(AW) 존재
   - **높이 < 2166 또는 색상 없음**: 'N' 표시
   - **높이 >= 2166 AND 색상 있음**:
     - '테스트' 시트의 `AH열`(Y 대상)에서 매칭 → 'Y'
     - '테스트' 시트의 `AG열`(N 대상)에서 매칭 → 'N'
     - 매칭 실패 → 빈칸

### 2. 단가 계산 (계산_영림발주서_가격_내부)
**실행 방법**: 메뉴 `💰 단가계산` 또는 `🚀 전체(단가+코드)`

**처리 흐름**:
1. 영림문틀단가표에서 제품 타입, 규격, 인필 방향으로 공급가 조회
2. 공급가 × 수량 = 기본 단가
3. **추가 금액 계산**:
   - 12-20행: 가스켓이 "없음/단종/단종예정"이 아니면 +5,500원
   - 22-35행: '테스트' 시트의 추가 금액 데이터 조회하여 가산
   - 22-35행: AQ='Y'이고 높이>=2166이면 도어 추가 금액 가산

### 3. 품목 코드 생성 (생성_품목코드_문틀_내부)
**실행 방법**: 메뉴 `📦 코드생성` 또는 `🚀 전체(단가+코드)`

**처리 흐름**:
1. **검증**: AT(너비) 또는 AV(높이) 중 하나라도 500 이상이어야 함
2. **품목명 생성**: 제품명, 색상, 규격, 치수 조합
3. **품목코드 생성**: 규칙 기반 코드 생성 로직
4. **단위 설정**: 12-20행은 "틀", 22-35행은 "짝"

---

## 오늘의 디버깅 세션 (2026-01-29)

### 발생한 문제들

#### 1. AP12:20 입력 차단 오류
**증상**: "입력값은 지정된 범위 내여야 합니다" 오류로 셀 입력 불가

**원인**: 
- 수동으로 설정된 잘못된 Data Validation 규칙
- 스크립트에서 명시적으로 관리하지 않던 범위

**해결**:
- `clearValidation_AP()` 함수로 기존 규칙 제거
- `setDropdowns_AP()` 함수로 올바른 드롭다운 재설정
- '필터' 시트 또는 '영림문틀단가표' 시트를 소스로 사용

#### 2. TypeError in findPriceFromMap_Scan
**증상**: `Cannot read properties of undefined` 오류

**원인**:
- '테스트' 시트의 데이터가 비어있거나 예상치 못한 형식
- `infoArr` 배열 인덱스 접근 시 null/undefined 체크 부재

**해결**:
- `loadTestSheetData_Optimized()`: try-catch 추가, 빈 배열 기본값 반환
- `findPriceFromMap_Scan()`: 모든 파라미터 null 체크 추가

#### 3. 행 수 불일치 오류 (전체 실행 시)
**증상**: "데이터의 행 수는 23개인데 범위의 행 수는 24개입니다"

**원인**:
- `생성_품목코드_문틀_내부()`에서 일부 조건 분기에서 배열 push 누락
- `continue` 문 사용 시 결과 배열에 항목을 추가하지 않음

**해결**:
- 모든 분기에서 반드시 빈 값이라도 배열에 push하도록 수정
- try-catch의 catch 블록에도 빈 값 push 추가

#### 4. AQ 열 자동 입력 미작동
**증상**: 높이(AV)나 색상(AW) 수정 시 인필(AQ)이 자동으로 채워지지 않음

**원인 1**: 트리거 누락
- 기존 로직은 `AW`(색상) 수정 시만 실행
- `AV`(높이) 수정 시에는 트리거되지 않음

**원인 2**: 수식 결과 인식
- `AV` 열에 `=2100-65` 같은 수식이 입력된 경우
- `getValue()`는 수식 결과를 반환하므로 문제없음 (확인됨)

**원인 3**: 비즈니스 로직 오류
- 높이 < 2166일 때 빈칸으로 설정 → **잘못됨**
- 올바른 로직: 높이 < 2166일 때 **'N'** 표시

**해결**:
- `onEdit` 트리거 조건에 `CONFIG.COLS.AV` 추가
- 높이 < 2166 또는 색상 없음 → `setValue("N")`으로 수정
- 디버그 로그 추가하여 실행 흐름 추적 가능하게 개선

---

## 주요 수정 내역

### 파일: door_v4.gs

#### 1. Data Validation 관리 함수 추가
```javascript
// AP12:20 드롭다운 복구
function setDropdowns_AP() { ... }

// AP12:20 유효성 검사 제거
function clearValidation_AP() { ... }

// 드롭다운 소스 탐색
function findCorrectAPSource() { ... }

// 유효성 검사 상태 확인
function inspectValidation() { ... }
```

#### 2. 에러 핸들링 강화
```javascript
// loadTestSheetData_Optimized: try-catch 추가
function loadTestSheetData_Optimized(ss) {
  try {
    // ... 기존 로직
    return { additionalPriceInfo: headers, ... };
  } catch (err) {
    Logger.log("Error: " + err.message);
    return { additionalPriceInfo: [], ... }; // 안전한 기본값
  }
}

// findPriceFromMap_Scan: null 체크 추가
function findPriceFromMap_Scan(target, map, infoArr) {
  if (!target || !map || !infoArr) return null;
  // ... 기존 로직
}
```

#### 3. 배열 길이 일치 보장
```javascript
// 생성_품목코드_문틀_내부: 모든 분기에서 배열 push
if (조건1) {
  names.push([""]); codes.push([""]); // 빈 값이라도 반드시 push
  실패++; continue;
}
if (조건2) {
  names.push([""]); codes.push([""]); // 빈 값이라도 반드시 push
  실패++; continue;
}
try {
  // 성공 로직
} catch (e) {
  names.push([""]); codes.push([""]); // catch에서도 push
  실패++;
}
```

#### 4. onEdit 트리거 확장 및 로직 수정
```javascript
// AV(높이) 트리거 추가
if (colS === CONFIG.COLS.AW || colS === CONFIG.COLS.AV) {
  // ... 처리 로직
}

// AQ 열 로직 수정
if (i >= 22) {
  if (avV >= 2166 && k) {
    // 테스트 시트 조회 로직
  } else {
    // 높이 < 2166 또는 색상 없음 → 'N' 표시
    s.getRange(i, CONFIG.COLS.AQ).setValue("N");
  }
}
```

#### 5. 디버그 로깅 추가
```javascript
function onEdit(e) {
  Logger.log("[onEdit] 발생 - 시트: " + sheetName + ", 행: " + rowS + ", 열: " + colS);
  Logger.log("[onEdit] 행 " + i + " 처리중 - 색상: '" + k + "', 높이: " + avV);
  Logger.log("[onEdit] 행 " + i + " 매칭 결과 - isY: " + isY + ", isN: " + isN);
  // ... 기타 로그
}
```

### 파일: .claspignore
```
# 레거시 스크립트 배포 제외
door_v0.gs
door_v1.gs
door_v2.gs
door_v3.gs
door_clean.gs
```

---

## 주의사항 및 유지보수 가이드

### ⚠️ 중요 주의사항

#### 1. 스크립트 버전 관리
- **절대 금지**: `door_v0.gs` ~ `door_v3.gs` 파일을 `.claspignore`에서 제거하지 말 것
- **이유**: 여러 `onEdit` 함수가 동시에 배포되면 예측 불가능한 동작 발생
- **현재 활성 버전**: `door_v4.gs`만 배포됨

#### 2. CONFIG 객체 수정 시
```javascript
const CONFIG = {
  SHEET_NAME: "영림발주서",  // 시트명 변경 시 여기 수정
  START_ROW: 12,             // 데이터 시작 행
  END_ROW: 35,               // 데이터 종료 행
  COLS: { ... }              // 열 번호 (1-based index)
};
```
- 시트 구조 변경 시 `CONFIG` 객체를 먼저 업데이트할 것
- 열 번호는 **1부터 시작** (A=1, B=2, ..., AP=42)

#### 3. 배열 길이 일치 원칙
```javascript
// ❌ 잘못된 예
for (var i = 0; i < num; i++) {
  if (조건) continue;  // 배열에 push 안함 → 길이 불일치!
  resultArray.push([값]);
}

// ✅ 올바른 예
for (var i = 0; i < num; i++) {
  if (조건) {
    resultArray.push([""]);  // 빈 값이라도 반드시 push
    continue;
  }
  resultArray.push([값]);
}
```

#### 4. onEdit 트리거 제약사항
- **수식 재계산 시 미작동**: `=2100-65`의 결과가 변경되어도 `onEdit`는 실행 안됨
- **해결책**: 사용자가 셀을 직접 수정해야 트리거 발생
- **대안**: 필요 시 메뉴에서 수동 실행 함수 제공

#### 5. 인필(AQ) 로직 기준값
```javascript
// 현재 기준: 높이 2166mm
if (avV >= 2166 && k) {
  // 테스트 시트 조회하여 Y/N 결정
} else {
  // 높이 미달 또는 색상 없음 → 'N'
}
```
- 기준값 변경 필요 시 `2166` 숫자만 수정
- 비즈니스 로직 변경 시 주석 업데이트 필수

### 🔧 디버깅 가이드

#### 로그 확인 방법
1. Apps Script 편집기 열기
2. 왼쪽 메뉴 **실행(Executions)** 클릭
3. 최근 실행 항목 선택
4. 로그 내용 확인

#### 주요 로그 메시지
```
[onEdit] 발생 - 시트: 영림발주서, 행: 22, 열: 48, 값: 2200
[onEdit] 처리 대상 열(AV/AW) 수정됨
[onEdit] 행 22 처리중 - 색상(AW): '화이트', 높이(AV): 2200
[onEdit] 행 22 매칭 결과 - isY: true, isN: false
```

#### 문제 해결 체크리스트
- [ ] 시트 이름이 "영림발주서"인가?
- [ ] 수정한 행이 12~35 범위인가?
- [ ] 수정한 열이 AV(48) 또는 AW(49)인가?
- [ ] 색상(AW) 값이 비어있지 않은가?
- [ ] '테스트' 시트가 존재하는가?
- [ ] COLOR_MAP, GASKET_COLOR_MAP 속성이 설정되어 있는가?

### 📝 배포 절차
```bash
# 1. 코드 수정 후 저장
# 2. 터미널에서 배포
cd "c:\Users\DSAI\Desktop\이비즈"
clasp push

# 3. 배포 확인
# Apps Script 편집기에서 door_v4.gs 내용 확인
```

### 🔄 롤백 절차
문제 발생 시 이전 버전으로 복구:
1. Apps Script 편집기에서 **파일 > 버전 기록** 클릭
2. 이전 정상 버전 선택
3. **복원** 클릭

---

## 참고 자료

### 관련 문서
- [PROJECT.md](./PROJECT.md): 프로젝트 개요
- [SESSION.md](./SESSION.md): 세션 로그
- [DECISIONS.md](./DECISIONS.md): 아키텍처 결정 기록

### Google Apps Script 공식 문서
- [SpreadsheetApp](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app)
- [onEdit Trigger](https://developers.google.com/apps-script/guides/triggers/events#edit)
- [Data Validation](https://developers.google.com/apps-script/reference/spreadsheet/data-validation)

### Clasp 도구
- [Clasp GitHub](https://github.com/google/clasp)
- [Clasp 사용법](https://developers.google.com/apps-script/guides/clasp)

---

**문서 작성**: AI Assistant (Antigravity)  
**최종 업데이트**: 2026-01-29  
**버전**: 1.0
