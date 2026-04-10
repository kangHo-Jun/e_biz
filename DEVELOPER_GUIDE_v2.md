# 📘 영림발주서 스크립트 개발 문서 v2 (Developer Guide)

이 문서는 대폭 최적화된 **`door_v2.gs`** 스크립트의 기술적 구조와 로직을 설명합니다.

---

## 1. ⚡ v2 핵심 최적화 기술

### 1.1 배치 처리 (Batch Processing)
기존 스크립트의 가장 큰 성능 병목은 루프 내에서의 잦은 `getValue()`, `setValue()` 호출이었습니다.
v2에서는 이를 해결하기 위해 다음과 같은 패턴을 모든 함수에 적용했습니다.

```javascript
// [Bad] 기존 방식: n번 통신
for (var i = 0; i < 24; i++) {
  sheet.getRange(i, 1).setValue(val); // 1행마다 API 호출
}

// [Good] v2 방식: 2번 통신
var values = sheet.getRange("A12:A35").getValues(); // 1. 전체 읽기
// ... 메모리 상에서 계산 ...
sheet.getRange("A12:A35").setValues(results); // 2. 전체 쓰기
```

### 1.2 해시 맵 검색 (O(1) Lookup)
기존의 선형 검색(Linear Search) 방식을 제거하고, **Map(Dictionary)** 자료구조를 도입했습니다.

*   **Before**: 1개의 매칭을 위해 테스트 시트 수백 행을 매번 순회 (O(N*M))
*   **After**: 데이터 로드 시점에 Map을 생성하여 즉시 조회 (O(1))
    *   `loadTestSheetData_Optimized()` 함수가 이 역할을 담당합니다.

### 1.3 범위 통일 및 중앙 제어 (`CONFIG`)
하드코딩된 행 번호를 제거하고 `CONFIG` 객체로 통합 관리합니다.

```javascript
const CONFIG = {
  SHEET_NAME: "영림발주서",
  TEST_SHEET_NAME: "테스트",
  START_ROW: 12,
  END_ROW: 35,  // 이전 버전의 34행 제한 문제 해결
  // ...
};
```

---

## 2. ⚙️ 주요 함수 구조

### 2.1 `계산_영림발주서_가격_내부`
*   **역할**: A열의 가격을 계산하고 메모를 업데이트합니다.
*   **범위**: 12행 ~ 35행
*   **프로세스**:
    1.  `CONFIG.START_ROW` ~ `CONFIG.END_ROW` 범위의 데이터를 **일괄 로드** (AP, AQ, AR... 등)
    2.  `테스트` 시트 데이터를 **HashMap** 형태로 메모리에 로드
    3.  각 행에 대해 메모리 상에서 가격 계산 수행
    4.  결과 배열(가격, 메모)을 생성하여 **일괄 기록** (`setValues`, `setNotes`)

### 2.2 `생성_품목코드_문틀_내부`
*   **역할**: BC~BF열에 품목 명, 코드, 단위를 생성합니다.
*   **특이사항**: 불필요하게 셀을 하나씩 건드리지 않고, 전체 범위(12-35)를 한 번에 업데이트합니다.

### 2.3 `onEdit(e)`
*   **개선점**: 다중 행(Multi-row) 편집 지원
*   **로직**:
    `e.range.rowStart`부터 `e.range.rowEnd`까지 루프를 돌며 처리합니다. 따라서 엑셀에서 복사-붙여넣기를 해도 모든 행이 정상적으로 트리거됩니다.

---

## 3. 💾 데이터 캐싱 (`PropertiesService`)

속도 향상을 위해 변경되지 않는 기준 정보는 스크립트 속성에 캐싱됩니다.

*   **`COLOR_MAP`**: 색상 코드 매핑 정보
*   **`GASKET_COLOR_MAP`**: 가스켓 및 단종 정보
*   **갱신**: `updateColorCodeMap`, `updateGasketColorMap` 함수 실행 시 갱신됨

---

## 4. 🛠️ 유지보수 가이드

### 4.1 행 범위 변경 시
`CONFIG` 객체의 `START_ROW`와 `END_ROW` 값만 수정하면 스크립트 전체에 반영됩니다.

### 4.2 새로운 추가금 로직 반영
`findPriceFromMap_Scan` 등의 헬퍼 함수를 통해 추가금 로직이 모듈화되어 있습니다.
특정 조건의 추가금 가격 계산 방식을 변경하려면 `계산_영림발주서_가격_내부` 함수의 Part 1, Part 2 섹션을 확인하세요.

---

**개발자 노트**: `door_v2.gs`는 성능과 유지보수성을 최우선으로 리팩토링되었습니다. 레거시 코드(`door.gs`, `door_v1.gs`)의 비효율적인 루프 구조를 다시 도입하지 않도록 주의해 주세요.
