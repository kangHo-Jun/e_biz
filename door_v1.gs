/**
 * 영림발주서 통합 스크립트 - 최종 버전
 * 
 * 기능:
 * 1. 단가계산: A12~A30에 가격 출력
 * 2. 코드생성: BC12~BF35에 품목명/코드 출력
 * 3. 전체실행: 1+2 순차 실행
 * 
 * 메뉴:
 * - 💰 단가계산
 * - 📦 코드생성
 * - 🚀 전체
 * - 🎨 실행 버튼 만들기
 * - 🗑️ 실행 버튼 삭제
 * - 📋 로그 보기 안내
 */

// ============================================
// 메뉴 및 초기화
// ============================================

/**
 * 시트 열 때 자동 실행 - 메뉴 생성
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 영림발주서')
    .addItem('💰 단가계산', '계산_영림발주서_가격')
    .addItem('📦 코드생성', '생성_품목코드_문틀')
    .addItem('🚀 전체', '전체_실행')
    .addItem('🧹 입력/출력 초기화', '초기화_영림발주서')
    .addToUi();
}

/**
 * [테스트] 23행만 실제 계산 로직으로 실행하고 결과를 팝업으로 표시
 * 단가계산 내부 함수와 완전히 동일한 로직 사용
 */
function testRow23Calculation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 발주서시트 = ss.getSheetByName("영림발주서");
  var 단가표시트 = ss.getSheetByName("영림문틀단가표");
  var 테스트시트 = ss.getSheetByName("테스트");
  
  var log = "=== 23행 실제 로직 테스트 ===\n";
  var i = 23; // 고정
  
  // ===== 데이터 로드 (계산_영림발주서_가격_내부와 동일) =====
  var 단가표데이터 = 단가표시트.getRange("C6:F500").getValues();
  
  var 추가금액정보 = null;
  var 추가금액데이터 = [];
  var 문짝가격데이터 = [];
  
  if (테스트시트) {
     var 전체범위 = 테스트시트.getRange("V3:Z100"); // 100행으로 축소
     var 전체값 = 전체범위.getValues();
     추가금액정보 = 전체값[0];
     추가금액데이터 = 전체값.slice(1);
     log += "Part1 데이터 로드: " + 추가금액데이터.length + "행, 가격정보=" + 추가금액정보 + "\n";
     
     var 문짝가격범위 = 테스트시트.getRange("AD1:AF50").getValues();
     for(var m=0; m<문짝가격범위.length; m++) {
        var kws = 문짝가격범위[m][0];
        var prc = 문짝가격범위[m][2];
        if(kws && prc) {
           문짝가격데이터.push({
              keywords: kws.toString().split(',').map(function(s){ return s.trim(); }), 
              price: Number(prc) || 0
           });
        }
     }
     log += "Part2 데이터 로드: " + 문짝가격데이터.length + "건\n";
  }
  
  // ===== 행 데이터 읽기 =====
  var ap값 = 발주서시트.getRange("AP" + i).getValue();
  var aq값 = 발주서시트.getRange("AQ" + i).getValue();
  var ar값 = 발주서시트.getRange("AR" + i).getValue();
  var as값 = 발주서시트.getRange("AS" + i).getValue();
  var ba값 = 발주서시트.getRange("BA" + i).getValue();
  var aw값 = 발주서시트.getRange("AW" + i).getValue();
  var ax값 = 발주서시트.getRange("AX" + i).getValue();
  var a값 = 발주서시트.getRange("A" + i).getValue();
  
  log += "\n[입력값]\n";
  log += "AP=" + ap값 + ", AQ=" + aq값 + ", AR=" + ar값 + "\n";
  log += "AS=" + as값 + ", BA=" + ba값 + "\n";
  log += "AW=" + aw값 + ", AX=" + ax값 + ", A=" + a값 + "\n";
  
  // ===== 계산 로직 (원본과 동일) =====
  var 최종가격 = 0;
  var 계산성공 = false;
  var isManualBase = false;
  
  // 1. 기본 단가 계산 시도
  if (ap값 && ap값.toString().trim() !== "") {
    var 제품타입 = 추출_제품타입(ap값);
    var 공급가 = 찾기_공급가(단가표데이터, 제품타입, as값, aq값);
    log += "\n[Step 1] 기본계산: 제품타입=" + 제품타입 + ", 공급가=" + 공급가 + "\n";
    
    if (공급가 !== null) {
      최종가격 = 공급가 * ar값;
      계산성공 = true;
      log += "   -> 최종가격 = " + 공급가 + " * " + ar값 + " = " + 최종가격 + "\n";
    }
  }
  
  // 2. 수동 베이스 (22~34행)
  if (!계산성공) {
     if (i >= 22 && i <= 34 && a값 !== "" && a값 !== null) {
        최종가격 = Number(a값);
        if (!isNaN(최종가격)) {
           계산성공 = true;
           isManualBase = true;
           log += "[Step 2] 기본계산 실패 -> A값 사용: " + 최종가격 + "\n";
        }
     }
  }
  
  log += "계산성공=" + 계산성공 + ", isManualBase=" + isManualBase + "\n";
  
  if (계산성공) {
    // Part 1: AW 매칭
    log += "\n[Part 1 시작]\n";
    if (i >= 22 && i <= 34 && aw값 && aw값.toString().trim() !== "" && 추가금액정보) {
        var keyword = aw값.toString().trim();
        var addedPrice = 0;
        var matchedCol = -1;
        log += "   조건 통과. keyword=" + keyword + "\n";
        
        outerLoop:
        for (var r = 0; r < 추가금액데이터.length; r++) {
           var rowData = 추가금액데이터[r];
           for (var c = 0; c < 5; c++) {
              var cellText = rowData[c] ? rowData[c].toString().toUpperCase().trim() : "";
              var kwUpper = keyword.toUpperCase();
              
              if (cellText && (cellText.includes(kwUpper) || kwUpper.includes(cellText))) {
                 matchedCol = c;
                 log += "   MATCH! Row=" + (r+4) + ", Col=" + c + ", Cell='" + cellText + "'\n";
                 break outerLoop;
              }
           }
        }
        
        if (matchedCol !== -1) {
           var priceVal = 추가금액정보[matchedCol];
           log += "   priceVal(col " + matchedCol + ")=" + priceVal + ", type=" + typeof priceVal + "\n";
           if (typeof priceVal === 'number') {
              addedPrice = priceVal;
              최종가격 += addedPrice;
              log += "   -> Part1 추가: " + addedPrice + ", 최종가격=" + 최종가격 + "\n";
           } else {
              log += "   ⚠️ priceVal이 숫자가 아님!\n";
           }
        } else {
           log += "   매칭 실패 (전체 " + 추가금액데이터.length + "행 검색)\n";
        }
    } else {
       log += "   조건 미충족. aw값=" + aw값 + ", 추가금액정보=" + (추가금액정보 ? "있음" : "없음") + "\n";
    }
    
    // Part 2: Door 키워드 매칭
    log += "\n[Part 2 시작]\n";
    if (i >= 22 && i <= 34 && 문짝가격데이터.length > 0) {
        var aqStr = aq값 ? aq값.toString().toUpperCase() : "";
        log += "   AQ(upper)=" + aqStr + "\n";
        
        if (aqStr.includes("Y")) {
           var targetStr = (aw값 ? aw값.toString() : "") + " " + (ax값 ? ax값.toString() : "");
           var targetUpper = targetStr.toUpperCase();
           log += "   targetStr=" + targetStr + "\n";
           
           var doorAddedPrice = 0;
           for(var d=0; d<문짝가격데이터.length; d++) {
              var entry = 문짝가격데이터[d];
              for(var k=0; k<entry.keywords.length; k++) {
                 var kw = entry.keywords[k].toString().toUpperCase().trim();
                 if(kw && targetUpper.includes(kw)) {
                    doorAddedPrice = entry.price;
                    log += "   MATCH! Kw='" + kw + "', Price=" + doorAddedPrice + "\n";
                    최종가격 += doorAddedPrice;
                    log += "   -> Part2 추가: " + doorAddedPrice + ", 최종가격=" + 최종가격 + "\n";
                    break;
                 }
              }
              if(doorAddedPrice > 0) break;
           }
           if (doorAddedPrice === 0) log += "   Part2 매칭 실패\n";
        } else {
           log += "   AQ에 Y 없음 -> 스킵\n";
        }
    }
  } else {
    log += "계산성공=false -> 추가 로직 스킵\n";
  }
  
  log += "\n=============================\n";
  log += "최종 결과: " + 최종가격;
  
  SpreadsheetApp.getUi().alert(log);
}

/**
 * [디버그] 선택한 행의 가격 계산 로직 상세 추적
 */
function debugPriceCalculationRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  
  if (sheet.getName() !== "영림발주서") {
    SpreadsheetApp.getUi().alert("영림발주서 시트에서 실행해주세요.");
    return;
  }
  
  var log = "=== 행 " + row + " 가격 계산 추적 ===\n";
  
  // 1. 입력값 읽기
  var ap = sheet.getRange("AP" + row).getValue();
  var aw = sheet.getRange("AW" + row).getValue();
  var ax = sheet.getRange("AX" + row).getValue();
  var aq = sheet.getRange("AQ" + row).getValue();
  var a = sheet.getRange("A" + row).getValue();
  
  log += "입력값: AP='" + ap + "', AW='" + aw + "', AX='" + ax + "', AQ='" + aq + "', A='" + a + "'\n";
  
  // 2. 테스트 시트 데이터 로드
  var testSheet = ss.getSheetByName("테스트");
  if (!testSheet) { return SpreadsheetApp.getUi().alert("테스트 시트 없음"); }
  
  // Part 1 데이터 (V~Z)
  var range1 = testSheet.getRange("V3:Z100"); // 100행까지만
  var values1 = range1.getValues();
  var priceInfos = values1[0]; // V3~Z3 (가격)
  var data1 = values1.slice(1); // V4~
  
  log += "Part1 가격정보(V3~Z3): " + priceInfos.join(",") + "\n";
  
  // Part 2 데이터 (AD~AF)
  var range2 = testSheet.getRange("AD1:AF100");
  var values2 = range2.getValues();
  var data2 = [];
  for(var i=0; i<values2.length; i++) {
     if(values2[i][0] && values2[i][2]) {
        data2.push({k: values2[i][0], p: values2[i][2]});
     }
  }
  log += "Part2 데이터 로드: " + data2.length + "건\n";
  
  // 3. Part 1 매칭 시뮬레이션
  var keyword = aw ? aw.toString().trim() : "";
  var part1Added = 0;
  
  log += "\n[Part 1 매칭 시도]\n";
  log += "Target Keyword(AW): '" + keyword + "'\n";
  
  if (keyword) {
      outer: for(var r=0; r<data1.length; r++) { // 전체 스캔
          for(var c=0; c<5; c++) {
              var cell = data1[r][c] ? data1[r][c].toString().toUpperCase().trim() : "";
              var kwUpper = keyword.toString().toUpperCase();
              
              if (cell && (cell.includes(kwUpper) || kwUpper.includes(cell))) {
                 log += "MATCH at Row " + (r+4) + " Col " + c + " (Cell='" + cell + "')\n";
                 var p = priceInfos[c];
                 if (typeof p === 'number') {
                    part1Added = p;
                    log += " -> Price: " + p + "\n";
                    break outer;
                 }
              }
          }
      }
      if (part1Added === 0) log += "매칭되는 항목 없음 (전체 " + data1.length + "행 검사)\n";
  } else {
     log += "AW 값 없음 -> 스킵\n";
  }
  
  // 4. Part 2 매칭 시뮬레이션
  var part2Added = 0;
  log += "\n[Part 2 매칭 시도]\n";
  var aqUpper = aq ? aq.toString().toUpperCase() : "";
  
  if (aqUpper.includes("Y")) {
      var targetStr = (aw?aw.toString():"") + " " + (ax?ax.toString():"");
      var targetUpper = targetStr.toUpperCase();
      log += "AQ='Y' 확인. TargetStr: '" + targetStr + "'\n";
      
      for(var i=0; i<data2.length; i++) {
         var kws = data2[i].k.toString().split(',').map(function(s){ return s.toString().toUpperCase().trim(); });
         for(var k=0; k<kws.length; k++) {
            // log += "  FullStr Check: '" + targetUpper + "' vs '" + kws[k] + "'\n";
            if (targetUpper.includes(kws[k])) {
                part2Added = Number(data2[i].p);
                log += "MATCH at Entry " + i + " (Kw='" + kws[k] + "') -> Price: " + part2Added + "\n";
                break; // found
            }
         }
         if (part2Added > 0) break;
      }
  } else {
      log += "AQ에 'Y' 없음 (" + aqUpper + ") -> 스킵\n";
  }
  
  // 5. 결론
  var base = Number(a) || 0;
  var total = base + part1Added + part2Added;
  log += "\n----------------------\n";
  log += "예상 결과: " + base + " + " + part1Added + " + " + part2Added + " = " + total;
  
  Logger.log(log);
  SpreadsheetApp.getUi().alert(log);
}

/**
 * [디버그] 문짝 가격 로직 데이터 검증
 */
function debugDoorPriceLogic() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 테스트시트 = ss.getSheetByName("테스트");
  
  if (!테스트시트) {
    SpreadsheetApp.getUi().alert("'테스트' 시트가 없습니다.");
    return;
  }
  
  // 1. 데이터 로드 확인
  var range = 테스트시트.getRange("AD1:AF50");
  var values = range.getValues();
  
  var log = "=== 문짝 가격 데이터 로드 (AD1:AF50) ===\n";
  var loadedCount = 0;
  
  // 시뮬레이션 타겟
  var targetAW = "PW1102";
  var targetAQ = "Y";
  var matchedPrice = 0;
  var matchedKeyword = "";
  
  for (var i = 0; i < values.length; i++) {
     var kws = values[i][0]; // AD
     var prc = values[i][2]; // AF
     
     if (kws && prc) {
        loadedCount++;
        // 3행(인덱스 2) 집중 확인
        if (i === 2) {
           log += "[Row 3 check] AD='" + kws + "', AF='" + prc + "'\n";
           log += " -> Parsed: " + JSON.stringify(kws.toString().split(',').map(function(s){ return s.trim(); })) + "\n";
        }
        
        var kwList = kws.toString().split(',').map(function(s){ return s.trim(); });
        var targetStr = targetAW; // AW만 테스트
        
        for (var k = 0; k < kwList.length; k++) {
           var kw = kwList[k];
           if (kw && targetStr.includes(kw)) {
             matchedPrice = prc;
             matchedKeyword = kw;
             log += "MATCH FOUND at Row " + (i+1) + ": Keyword='" + kw + "', Price=" + prc + "\n";
           }
        }
     }
  }
  
  log += "Total loaded rows: " + loadedCount + "\n";
  log += "--------------------------------\n";
  log += "Simulation (AW='PW1102', AQ='Y'):\n";
  log += "Matched Price: " + matchedPrice + "\n";
  
  Logger.log(log);
  SpreadsheetApp.getUi().alert(log);
}

/**
 * [디버그] 선택한 행의 품목 타입 분석 확인
 */
function debugSelectRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  
  if (row < 12) {
    SpreadsheetApp.getUi().alert("12행 이상에서 실행해주세요.");
    return;
  }
  
  var ap = sheet.getRange("AP" + row).getValue();
  var type = 구분_품목타입(ap, row);
  var model = "";
  if (type === 'DOOR') {
     model = 모델코드_생성(ap);
  } else {
     model = 플래그코드_생성(ap);
  }
  
  var msg = "행: " + row + "\n" +
            "품명(AP): " + ap + "\n" +
            "판정결과: " + type + "\n" +
            "중간코드: " + model;
            
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 로그 확인 안내
 */
function 로그보기() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('로그 확인 방법',
    '1. 상단 메뉴에서 "보기" 클릭\n' +
    '2. "로그" 또는 "Logs" 선택\n\n' +
    '또는\n\n' +
    '1. Ctrl+Enter (또는 Cmd+Enter) 누르기\n\n' +
    '로그에서 각 행의 처리 과정을 확인할 수 있습니다.',
    ui.ButtonSet.OK);
}

// ============================================
// 1. 단가계산 (A열 출력)
// ============================================

/**
 * 단가계산 실행 (UI용)
 */
function 계산_영림발주서_가격() {
  try {
    var 결과 = 계산_영림발주서_가격_내부();

    Logger.log("\n✅ 단가계산 완료!");
    Logger.log("  성공: " + 결과.성공 + "개");
    Logger.log("  실패: " + 결과.실패 + "개");
    Logger.log("  출력: A12~A30\n");

  } catch (e) {
    Logger.log("❌ 단가계산 오류: " + e.message);
  }
}

/**
 * 단가계산 내부 함수
 */
function 계산_영림발주서_가격_내부() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 발주서시트 = ss.getSheetByName("영림발주서");
  var 단가표시트 = ss.getSheetByName("영림문틀단가표");

  if (!발주서시트) {
    throw new Error('"영림발주서" 시트를 찾을 수 없습니다.');
  }
  if (!단가표시트) {
    throw new Error('"영림문틀단가표" 시트를 찾을 수 없습니다.');
  }

  var 시작행 = 12;
  var 끝행 = 34; // 32->34로 확장
  var 단가표데이터 = 단가표시트.getRange("C6:F500").getValues();

  // [신규] 테스트 시트 V:Z 데이터 로드
  // V3~Z3: 추가금액 정의 (헤더행)
  // V4~Z300: 검색 데이터
  var 테스트시트 = ss.getSheetByName("테스트");
  var 추가금액정보 = null;
  var 추가금액데이터 = []; // 2차원 배열 [행][열]
  
  if (테스트시트) {
     var 전체범위 = 테스트시트.getRange("V3:Z300");
     var 전체값 = 전체범위.getValues();
     추가금액정보 = 전체값[0]; // V3~Z3 (0번째 행)
     추가금액데이터 = 전체값.slice(1); // V4~Z300
     Logger.log("✅ 추가 금액 정보 로드: " + 추가금액정보);
     
     // [신규] 문짝 키워드 가격 정보 로드 (AD:키워드, AF:가격)
     // AD3:AF300 범위 사용 (AD:30열, AF:32열) -> 인덱스: AD(0), AE(1), AF(2) relative to AD
     // 그러나 기존 load 방식 유지를 위해 별도 range 호출
     var 문짝가격범위 = 테스트시트.getRange("AD1:AF50").getValues();
     var 문짝가격데이터 = [];
     for(var m=0; m<문짝가격범위.length; m++) {
        // AD열(0): 키워드들(쉼표구분), AF열(2): 가격
        var kws = 문짝가격범위[m][0];
        var prc = 문짝가격범위[m][2];
        if(kws && prc) {
           문짝가격데이터.push({
              keywords: kws.toString().split(',').map(function(s){ return s.trim(); }), 
              price: Number(prc) || 0
           });
        }
     }
     Logger.log("✅ 문짝 추가 가격 정보 로드: " + 문짝가격데이터.length + "건");
  }

  Logger.log("========================================");
  Logger.log("💰 단가계산 시작");
  Logger.log("단가표 데이터 행 수: " + 단가표데이터.length);
  Logger.log("========================================\n");

  var 처리결과 = [];
  var 성공카운트 = 0;
  var 실패카운트 = 0;

  for (var i = 시작행; i <= 끝행; i++) {
    // [성능 개선] AT, AV 모두 비어있으면 스킵
    var at값_체크 = 발주서시트.getRange("AT" + i).getValue();
    var av값_체크 = 발주서시트.getRange("AV" + i).getValue();
    if (!at값_체크 && !av값_체크) {
       처리결과.push([""]);
       실패카운트++;
       continue;
    }
    
    var apInRange = 발주서시트.getRange("AP" + i);
    var aInRange  = 발주서시트.getRange("A" + i);
    
    var ap값 = apInRange.getValue();
    var aq값 = 발주서시트.getRange("AQ" + i).getValue();
    var ar값 = 발주서시트.getRange("AR" + i).getValue();
    var as값 = 발주서시트.getRange("AS" + i).getValue();
    var ba값 = 발주서시트.getRange("BA" + i).getValue(); 
    var aw값 = 발주서시트.getRange("AW" + i).getValue();
    var ax값 = 발주서시트.getRange("AX" + i).getValue(); // [수정] 누락된 ax값 추가
    var av값 = 발주서시트.getRange("AV" + i).getValue(); // [신규] av값 추가
    var a값  = aInRange.getValue();
    var aNote = aInRange.getNote(); // 메모 확인

    // [중복 방지] 이미 계산된 값이 있고(메모 포함), 재계산을 원치 않는 경우 스킵
    if (a값 && aNote && (aNote.includes("✅") || aNote.includes("계산완료"))) {
       Logger.log("  행 " + i + ": 이미 계산됨(메모) → 스킵");
       처리결과.push([a값]); // 기존 값 유지
       성공카운트++;
       continue;
    }

    Logger.log("  행 " + i + ": AP=" + ap값 + ", BA=" + ba값 + ", A=" + a값);

    var 최종가격 = 0;
    var 계산성공 = false;
    var isManualBase = false; // 수동 A값을 베이스로 했는지 여부

    // 1. 기본 단가 계산 시도 (AP값이 있을 때)
    if (ap값 && ap값.toString().trim() !== "") {
      var 제품타입 = 추출_제품타입(ap값);
      var 공급가 = 찾기_공급가(단가표데이터, 제품타입, as값, aq값);

      if (공급가 !== null) {
        최종가격 = 공급가 * ar값;
        계산성공 = true;
        // 자동 계산 시에는 메모 초기화 (새로운 계산이므로)
        aInRange.clearNote();
      }
    }
    
    // 2. 계산 실패/스킵 시 & 22~34행 & A값 존재 시 -> A값을 베이스로 사용
    if (!계산성공) {
       if (i >= 22 && i <= 34 && a값 !== "" && a값 !== null) {
          최종가격 = Number(a값); 
          if (!isNaN(최종가격)) {
             계산성공 = true;
             isManualBase = true; // 수동 베이스임
             Logger.log("    📌 기본 계산 없음 → 기존 A값 유지: " + 최종가격);
          }
       }
    }

    if (계산성공) {
      // 3. 추가 금액 로직
      var extraAdded = false;

      // A) BA열에 문자가 있고, "없음"이 아니면 5,500 추가
      // [수정] "단종", "단종예정", "없음" 제외
      var baStr = ba값 ? ba값.toString().trim() : "";
      if (baStr !== "" && baStr !== "없음" && baStr !== "단종" && baStr !== "단종예정") {
         if (i >= 22 && i <= 34) {
            Logger.log("    🚫 22~34행 구간이므로 BA값 5,500원 추가 제외");
         } else {
            최종가격 += 5500;
            Logger.log("    📌 BA값 '" + ba값 + "' 있음 → 5,500원 추가");
         }
      }
      
      // B) 22행~34행이고 AW값 매칭 시 추가 금액 가산
      if (i >= 22 && i <= 34 && aw값 && aw값.toString().trim() !== "" && 추가금액정보) {
          
          // 이미 추가금이 반영된 상태인지 확인 (메모 체크) - 디버깅 위해 잠시 비활성화
          // 이미 추가금이 반영된 상태인지 확인 (메모 체크)
          // [수정] 주석 해제 및 강화: 수동 베이스이고 메모가 있으면 스킵
          if (isManualBase && aNote && aNote.includes("✅")) {
             Logger.log("    ⚠️ 이미 추가금 반영됨 (메모확인) → 추가 계산 스킵");
          } else {
             // 추가 금액 계산 실행
              var keyword = aw값.toString().trim();
              var addedPrice = 0;
              var matchedCol = -1;
              
              outerLoop:
              for (var r = 0; r < 추가금액데이터.length; r++) {
                 var rowData = 추가금액데이터[r];

                 for (var c = 0; c < 5; c++) { // V,W,X,Y,Z
                    var cellText = rowData[c] ? rowData[c].toString().toUpperCase().trim() : "";
                    var kwUpper = keyword.toUpperCase(); // keyword는 이미 trim됨 (line 299)
                    
                    // [수정] 양방향 포함 관계 체크 (디버그툴과 동일하게)
                    // Row 23 디버깅
                    if (i === 23 && c === 0 && r < 3) {
                        Logger.log("    [ROW 23 DEBUG] TargetKW='" + kwUpper + "', Cell='" + cellText + "' -> Match? " + (cellText && (cellText.includes(kwUpper) || kwUpper.includes(cellText))));
                    }
                    
                    if (cellText && (cellText.includes(kwUpper) || kwUpper.includes(cellText))) {
                       matchedCol = c;
                       Logger.log("    🔥 [매칭성공] 행 " + i + ": Cell='" + cellText + "' matches '" + kwUpper + "'");
                       break outerLoop;
                    }
                 }
              }
              
              if (matchedCol !== -1) {
                 var priceVal = 추가금액정보[matchedCol]; 
                 if (typeof priceVal === 'number') {
                    addedPrice = priceVal;
                    최종가격 += addedPrice;
                    extraAdded = true; // 추가됨 표시
                    Logger.log("    📌 AW값 매칭 [" + keyword + "] → " + addedPrice + "원 추가");
                 }
              }
      }
      }
      
      // C) [신규] 문짝 추가 금액 로직 (22~34행 & AQ값 "Y" 포함 & AW/AX 키워드 매칭)
      if (i >= 22 && i <= 34 && 문짝가격데이터.length > 0) {
          var aqStr = aq값 ? aq값.toString().toUpperCase() : ""; // 대소문자 무시
          if (aqStr.includes("Y")) {
             var targetStr = (aw값 ? aw값.toString() : "") + " " + (ax값 ? ax값.toString() : "");
             Logger.log("    🔎 문짝 옵션 검사 (행 " + i + "): AQ='" + aqStr + "', Target='" + targetStr + "'");
             
             var doorAddedPrice = 0;
             var matchedKw = "";
             
             // 문짝가격데이터 순회
             for(var d=0; d<문짝가격데이터.length; d++) {
                var entry = 문짝가격데이터[d];
                // 키워드 중 하나라도 포함되면
                for(var k=0; k<entry.keywords.length; k++) {
                   var kw = entry.keywords[k].toString().toUpperCase().trim();
                   var targetUpper = targetStr.toUpperCase();
                   
                   if(kw && targetUpper.includes(kw)) {
                      doorAddedPrice = entry.price;
                      matchedKw = kw;
                      break;
                   }
                }
                if(doorAddedPrice > 0) break; // 매칭되면 중단 (우선순위: 위쪽 데이터?? 일단 첫매칭)
             }
             
             if(doorAddedPrice > 0) {
                // [신규] AV값 >= 2166 일 때만 추가
                var avNum = Number(av값) || 0;
                if (avNum >= 2166) {
                   최종가격 += doorAddedPrice;
                   Logger.log("    📌 문짝 옵션 매칭 [" + matchedKw + "] (AQ='Y', AV=" + avNum + " >= 2166) → " + doorAddedPrice + "원 추가");
                } else {
                   Logger.log("    ⚠️ 문짝 옵션 매칭 했지만 AV=" + avNum + " < 2166 → 추가 안함");
                }
             }
          }
      }

      처리결과.push([최종가격]);
      성공카운트++;
      Logger.log("    ✅ 최종: " + 최종가격);
      
      // 결과 출력 후 메모 업데이트 (수동 베이스이고, 추가금이 더해졌을 때만)
      if (isManualBase && extraAdded) {
         발주서시트.getRange("A" + i).setValue(최종가격).setNote("✅추가금반영됨"); 
         // setValues로 한꺼번에 처리하기 어려우므로 개별 처리하거나,
         // 처리결과 배열을 사용하되 메모는 별도로 루프 밖에서 처리해야 함.
         // 여기서는 루프 안에서 즉시 메모 설정 (성능 영향 적음)
      } else if (!isManualBase) {
         // 자동 계산인 경우 값만 저장 (setValues가 나중에 처리함)
         // 하지만 루프 밖 setValues와 충돌할 수 있으므로, 
         // 메모 설정이 필요한 경우만 여기서 setValue 하고, 
         // 처리결과 배열에는 빈값을 넣어서 덮어쓰기 방지? 
         // 아니면 setValues를 전체적으로 하고, 메모만 따로?
         
         // 수정 전략: loop 끝난 후 setValues는 모든 값을 덮어쓴다.
         // 따라서 여기서 setValue를 하면 loop 후 setValues에 의해 덮어써짐.
         // '처리결과' 배열에 최종가격을 넣는 것은 맞음.
         // 메모 설정은 '단가계산 완료' 후 별도로 하거나, range 객체를 미리 잡아두고 해야함.
         
         // 여기서는 메모 설정만 예약하거나 즉시 실행.
         // setValues가 실행되면 값은 바뀌고 메모는 유지됨.
         if (isManualBase && extraAdded) {
             aInRange.setNote("✅추가금반영됨");
         }
      }

    } else {
      처리결과.push([""]);
      실패카운트++;
      Logger.log("    ❌ 실패 (빈값 처리)");
    }
  }

  // [디버그] 23행 결과 확인 (인덱스 11 = 23-12)
  Logger.log("🔍 처리결과 배열 (23행 = index 11): " + JSON.stringify(처리결과[11]));
  Logger.log("🔍 처리결과 전체 길이: " + 처리결과.length);
  
  발주서시트.getRange("A" + 시작행 + ":A" + 끝행).setValues(처리결과);

  Logger.log("\n단가계산 완료 - 성공: " + 성공카운트 + "개, 실패: " + 실패카운트 + "개");

  return {
    성공: 성공카운트,
    실패: 실패카운트
  };
}

/**
 * AP값에서 "*방" 앞 문자 추출
 */
function 추출_제품타입(ap값) {
  if (!ap값 || ap값 === "") {
    return "";
  }

  var 문자열 = ap값.toString();
  var 방패턴 = /(\d+)방/;
  var 매칭 = 문자열.match(방패턴);

  if (매칭) {
    var 방위치 = 문자열.indexOf(매칭[0]);
    var 결과 = 문자열.substring(0, 방위치);

    if (결과.endsWith("ㅣ")) {
      결과 = 결과.substring(0, 결과.length - 1);
    }

    return 결과;
  }

  return 문자열;
}

/**
 * 제품타입 키워드 정규화 ("형" 제거)
 */
function 정규화_키워드(키워드) {
  if (!키워드) return "";

  var 정규화 = 키워드.toString().trim();

  if (정규화.endsWith("형")) {
    정규화 = 정규화.substring(0, 정규화.length - 1);
  }

  return 정규화;
}

/**
 * 단가표에서 공급가 찾기
 */
function 찾기_공급가(단가표데이터, 제품타입, 사이즈, 방향) {
  if (!제품타입 || !사이즈 || !방향) {
    return null;
  }

  var 제품타입_정규화 = 제품타입.toString().trim();
  var 사이즈_정규화 = 사이즈.toString().trim();
  var 방향_정규화 = 방향.toString().trim();

  for (var i = 0; i < 단가표데이터.length; i++) {
    var 행 = 단가표데이터[i];
    var c열 = 행[0] ? 행[0].toString().trim() : "";
    var d열 = 행[1] ? 행[1].toString().trim() : "";
    var e열 = 행[2] ? 행[2].toString().trim() : "";
    var f열 = 행[3];

    if (!c열 && !d열 && !e열) {
      continue;
    }

    var 제품타입키워드 = 제품타입_정규화.split('ㅣ');
    var c열정규화 = c열.replace(/형/g, '');

    var c열포함 = 제품타입키워드.every(function (키워드) {
      var 키워드정규화 = 정규화_키워드(키워드);
      return c열정규화.includes(키워드정규화);
    });
    var d열일치 = d열 === 사이즈_정규화;
    var e열일치 = e열 === 방향_정규화;

    if (c열포함 && d열일치 && e열일치) {
      if (typeof f열 === 'number') {
        return f열;
      } else if (f열 && !isNaN(f열)) {
        return Number(f열);
      } else {
        return null;
      }
    }
  }

  return null;
}

// ============================================
// 2. 코드생성 (BC~BF열 출력)
// ============================================

/**
 * 코드생성 실행 (UI용)
 */
function 생성_품목코드_문틀() {
  try {
    var 결과 = 생성_품목코드_문틀_내부();

    Logger.log("\n✅ 코드생성 완료!");
    Logger.log("  성공: " + 결과.성공 + "개");
    Logger.log("  실패: " + 결과.실패 + "개");
    Logger.log("  출력: BC12~BF35\n");

  } catch (e) {
    Logger.log("❌ 코드생성 오류: " + e.message);
  }
}

/**
 * 코드생성 내부 함수
 */
function 생성_품목코드_문틀_내부() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 발주서시트 = ss.getSheetByName("영림발주서");

  if (!발주서시트) {
    throw new Error('"영림발주서" 시트를 찾을 수 없습니다.');
  }

  var 시작행 = 12;
  var 끝행 = 35;

  var AT전체 = 발주서시트.getRange("AT" + 시작행 + ":AT" + 끝행).getValues();
  var AV전체 = 발주서시트.getRange("AV" + 시작행 + ":AV" + 끝행).getValues();

  var 처리할행목록 = [];
  for (var i = 0; i < AT전체.length; i++) {
    var at값 = AT전체[i][0];
    var av값 = AV전체[i][0];

    if (at값 || av값) {
      처리할행목록.push(시작행 + i);
    }
  }

  if (처리할행목록.length === 0) {
    Logger.log("⚠️ 처리할 데이터가 없습니다 (AT, AV 모두 비어있음)");
    return {
      성공: 0,
      실패: 0
    };
  }

  Logger.log("========================================");
  Logger.log("📦 코드생성 시작 (Ver11 로직)");
  Logger.log("처리할 행 수: " + 처리할행목록.length + "개");
  Logger.log("========================================\n");

  var 전체행수 = 끝행 - 시작행 + 1;
  var 품목명결과 = new Array(전체행수).fill([""]);
  var 품목코드결과 = new Array(전체행수).fill([""]);
  var 빈칸결과 = new Array(전체행수).fill([""]);
  var 단위결과 = new Array(전체행수).fill([""]);

  var 성공카운트 = 0;
  var 실패카운트 = 0;

  for (var idx = 0; idx < 처리할행목록.length; idx++) {
    var i = 처리할행목록[idx];
    var 배열인덱스 = i - 시작행;

    Logger.log("  행 " + i + " 처리중 (" + (idx + 1) + "/" + 처리할행목록.length + ")");

    var aw값 = 발주서시트.getRange("AW" + i).getValue();
    var ax값 = 발주서시트.getRange("AX" + i).getValue();
    var ap값 = 발주서시트.getRange("AP" + i).getValue();
    var as값 = 발주서시트.getRange("AS" + i).getValue();
    var at값 = 발주서시트.getRange("AT" + i).getValue();
    var av값 = 발주서시트.getRange("AV" + i).getValue();
    var aq값 = 발주서시트.getRange("AQ" + i).getValue();

    var as숫자 = 추출_숫자_from문자열(as값);
    var 규격숫자들 = [as숫자, Number(at값) || 0, Number(av값) || 0];
    var 최대규격 = Math.max.apply(null, 규격숫자들);

    if (최대규격 <= 999) {
      Logger.log("    ⚠️ 규격 최대값 ≤ 999 (부품) → 생성 안 함");
      실패카운트++;
      continue;
    }

    try {
      var 품목명 = 생성_품목명(ap값, aw값, ax값, as값, at값, av값, aq값, i); // i=행번호 전달
    } catch (e) {
      Logger.log("    ❌ 품목명 생성 실패: " + e.message);
      실패카운트++;
      continue;
    }

    try {
      var 품목코드 = 생성_품목코드_NEW(ap값, aw값, ax값, as값, at값, av값, aq값, i); // NEW 함수 호출
    } catch (e) {
      Logger.log("    ❌ 품목코드 생성 실패: " + e.message);
      품목명결과[배열인덱스] = [품목명];
      // 문짝 구간(22~34)이면 단위 '짝', 아니면 '틀'
    if (i >= 22 && i <= 34) {
       단위결과[배열인덱스] = ["짝"];
    } else {
       단위결과[배열인덱스] = ["틀"];
    }
      실패카운트++;
      continue;
    }

    품목명결과[배열인덱스] = [품목명];
    품목코드결과[배열인덱스] = [품목코드];
    빈칸결과[배열인덱스] = [""];
    // 문짝 구간(22~34)이면 단위 '짝', 아니면 '틀'
    if (i >= 22 && i <= 34) {
       단위결과[배열인덱스] = ["짝"];
    } else {
       단위결과[배열인덱스] = ["틀"];
    }
    성공카운트++;
    Logger.log("    ✅ 성공");
  }

  발주서시트.getRange("BC" + 시작행 + ":BC" + 끝행).setValues(품목명결과);
  발주서시트.getRange("BD" + 시작행 + ":BD" + 끝행).setValues(품목코드결과);
  발주서시트.getRange("BE" + 시작행 + ":BE" + 끝행).setValues(빈칸결과);
  발주서시트.getRange("BF" + 시작행 + ":BF" + 끝행).setValues(단위결과);

  Logger.log("\n코드생성 완료 - 성공: " + 성공카운트 + "개, 실패: " + 실패카운트 + "개");

  return {
    성공: 성공카운트,
    실패: 실패카운트
  };
}

/**
 * 문자열에서 숫자만 추출
 */
function 추출_숫자_from문자열(값) {
  if (!값) return 0;

  var 문자열 = 값.toString();
  var 매칭 = 문자열.match(/\d+/);

  return 매칭 ? Number(매칭[0]) : 0;
}

/**
 * 품목명 생성
 */
function 생성_품목명(ap값, aw값, ax값, as값, at값, av값, aq값, row) { // row 추가
  var 품명 = ap값 ? ap값.toString() : "";
  var 타입 = 구분_품목타입(품명, row); // row 전달

  // === 문짝 로직 (shop.gs 이식) ===
  if (타입 === 'DOOR') {
     var 회사명 = "영림"; // 기본값
     if (품명.includes("영림")) 회사명 = "영림";
     
     // 1. 색상 전처리 (shop.gs: preprocessColorForProductName)
     // door.gs의 색상_전처리 함수 활용하되, 회사명 체크
     var 색상 = 색상_전처리(aw값, ax값);
     
     var 최종색상 = "";
     // 색상에 이미 회사명이 있으면 그대로, 없으면 붙임
     // shop.gs의 shouldAddCompanyPrefix 로직 단순화
     if (색상.startsWith("영림")) {
        최종색상 = 색상;
     } else {
        최종색상 = 회사명 + 색상;
     }

     // 2. 품명 전처리 (shop.gs: preprocessItemNameForProductName)
     var 최종품명 = 품명_전처리_문짝(품명, as값 + "*" + at값 + "*" + av값); // spec 문자열 조합해서 전달

     // 3. 규격 전처리 (shop.gs: preprocessSpecForProductName)
     // 여기선 as, at, av, aq 조합
     var 식기표시 = "";
     if (aq값) {
        if (aq값.toString().includes("3방") || aq값.toString().includes("식기무")) 식기표시 = "식기무";
        else if (aq값.toString().includes("식기유")) 식기표시 = "식기유";
     }
     // 기본 규격 문자열
     var 규격문자열 = as값 + "*" + at값 + "*" + av값;
     if (식기표시) 규격문자열 += 식기표시; // shop.gs는 / N 등을 식기무로 바꿈
     
     // shop.gs 스타일 규격 전처리 (필요하다면)
     var 최종규격 = 규격문자열; // shop.gs는 specRaw를 preprocessSpecForProductName 하는데, 여기선 일단 조합
     
     return 최종색상 + " " + 최종품명 + " " + 최종규격;
  }

  // === 기존 문틀 로직 ===
  var 회사명 = "영림";
  if (ap값 && ap값.toString().includes("영림")) {
    회사명 = "영림";
  }

  var 색상 = 색상_전처리(aw값, ax값);
  var 품명 = 품명_전처리(ap값);
  var as숫자 = 추출_숫자_from문자열(as값);
  var 식기표시 = aq값 && aq값.toString().includes("3방") ? "식기무" : "식기유";
  var 규격 = as숫자 + "*" + at값 + "*" + av값 + 식기표시;

  var 최종색상 = "";
  if (색상 && 색상.toString().indexOf("영림") === 0) {
    최종색상 = 색상;
    Logger.log("    색상이 '영림'으로 시작 → 회사명 추가 안 함");
  } else {
    최종색상 = 회사명 + 색상;
    Logger.log("    색상이 '영림'으로 시작 안 함 → 회사명 추가");
  }

  return 최종색상 + " " + 품명 + " " + 규격;
}

/**
 * 색상 전처리
 */
function 색상_전처리(aw값, ax값) {
  var 색상1 = aw값 ? aw값.toString().trim() : "";
  var 색상2 = ax값 ? ax값.toString().trim() : "";

  var 조합 = "";
  if (색상1 && 색상2) {
    조합 = 색상1 + " " + 색상2;
  } else if (색상1) {
    조합 = 색상1;
  } else if (색상2) {
    조합 = 색상2;
  }

  if (!조합) return "";

  var 패턴1 = /^(영림\d+)\s+(PS\d+|[A-Z]+\d+)$/i;
  if (패턴1.test(조합)) {
    var 결과 = 조합.replace(/\s+/g, '');
    Logger.log("    색상 패턴1: " + 조합 + " → " + 결과);
    return 결과;
  }

  var 패턴2 = /^(영림\d+)\s+[가-힣]+$/;
  if (패턴2.test(조합)) {
    var 매칭 = 조합.match(/영림\d+/);
    var 결과 = 매칭 ? 매칭[0] : 조합;
    Logger.log("    색상 패턴2: " + 조합 + " → " + 결과);
    return 결과;
  }

  if (/^[가-힣\s]+$/.test(조합)) {
    var 결과 = 조합.replace(/\s+/g, '');
    Logger.log("    색상 패턴3 (한글만): " + 조합 + " → " + 결과);
    return 결과;
  }

  var 결과 = 조합.replace(/영림|우딘|예림/g, '').trim();
  Logger.log("    색상 기본 처리: " + 조합 + " → " + 결과);

  return 결과;
}

/**
 * 품명 전처리
 */
function 품명_전처리(ap값) {
  if (!ap값) return "";

  var 품명 = ap값.toString();

  품명 = 품명.replace(/^영림ㅣ/, '');
  품명 = 품명.replace(/ㅣ/g, ' ');
  품명 = 품명.replace(/문틀/g, '');
  품명 = 품명.replace(/형/g, '');
  품명 = 품명.replace(/\d+바/g, '');
  품명 = 품명.replace(/\(식기[XO]\)/g, '');
  품명 = 품명.trim().replace(/\s+/g, ' ');

  Logger.log("    품명 전처리: " + ap값 + " → " + 품명);

  return 품명;
}

/**
 * 품목코드 생성 (NEW)
 */
function 생성_품목코드_NEW(ap값, aw값, ax값, as값, at값, av값, aq값, row) { // row 추가
  // 1. 브랜드/색상코드
  var 브랜드색상코드 = 브랜드색상코드_생성(aw값, ax값);
  Logger.log("    [코드생성] 브랜드색상: " + 브랜드색상코드);

  // 2. 타입 확인 (문틀 vs 문짝)
  var 품명 = ap값 ? ap값.toString() : "";
  var 타입 = 구분_품목타입(품명, row); // row 전달
  
  Logger.log("    [코드생성] 행: " + row + ", 타입: " + 타입 + ", 품명: " + 품명);
  
  var 중간코드 = "";
  var 규격코드 = "";

  if (타입 === 'DOOR') {
     // 문짝: 모델코드 사용
     중간코드 = 모델코드_생성(품명);
     if (!중간코드) {
        중간코드 = ""; 
        Logger.log("    ⚠️ 문짝 모델코드 생성 실패 (빈값)");
     } else {
        Logger.log("    ✅ 문짝 모델코드 생성: " + 중간코드);
     }
     
     // 문짝 규격: 숫자만 연결 (식기표시 제외) - shop.gs 스타일 (359002100)
     // door.gs는 as, at, av가 분리되어 있으므로 그대로 연결
     var as숫자 = 추출_숫자_from문자열(as값);
     규격코드 = "" + as숫자 + at값 + av값;
  } else {
     // 문틀: 플래그코드 사용
     중간코드 = 플래그코드_생성(ap값);
     
     // 문틀 규격: 숫자 + 식기표시(Y/N)
     규격코드 = 규격코드_생성(as값, at값, av값, aq값);
  }

  Logger.log("    [코드생성] 최종조합: " + 브랜드색상코드 + " + " + 중간코드 + " + " + 규격코드);

  return 브랜드색상코드 + 중간코드 + 규격코드;
}

/**
 * 브랜드/색상 코드 생성
 */
function 브랜드색상코드_생성(aw값, ax값) {
  var 색상1 = aw값 ? aw값.toString().trim() : "";
  var 색상2 = ax값 ? ax값.toString().trim() : "";

  var 조합 = "";
  if (색상1 && 색상2) {
    조합 = 색상1 + 색상2;
  } else if (색상1) {
    조합 = 색상1;
  } else if (색상2) {
    조합 = 색상2;
  }

  if (!조합) throw new Error("색상 정보 없음");

  var 브랜드 = "Y";

  var 패턴1 = /영림(\d+)PS\d+/;
  var 매칭1 = 조합.match(패턴1);
  if (매칭1) {
    return 브랜드 + 매칭1[1];
  }

  var 패턴2 = /PS([A-Z]+\d+)/i;
  var 매칭2 = 조합.match(패턴2);
  if (매칭2) {
    return 브랜드 + "S" + 매칭2[1];
  }

  var 패턴3 = /영림(\d+)/;
  var 매칭3 = 조합.match(패턴3);
  if (매칭3) {
    return 브랜드 + 매칭3[1];
  }

  if (/^[가-힣]+$/.test(조합)) {
    var 한글2자 = 조합.substring(0, 2);
    return 브랜드 + 한글2자;
  }

  var 패턴5 = /(\d+)/;
  var 매칭5 = 조합.match(패턴5);
  if (매칭5) {
    return 브랜드 + 매칭5[1];
  }

  throw new Error("색상 코드 생성 실패: " + 조합);
}

/**
 * 플래그 코드 생성
 */
function 플래그코드_생성(ap값) {
  if (!ap값) throw new Error("AP값 없음");

  var 플래그문자열 = ap값.toString();

  플래그문자열 = 플래그문자열.replace(/^영림ㅣ/, '');
  var 키워드들 = 플래그문자열.split('ㅣ');

  키워드들 = 키워드들.map(function (k) {
    return k.replace(/형/g, '').trim();
  });

  Logger.log("    플래그 키워드: " + JSON.stringify(키워드들));

  var 연동숫자 = null;
  for (var i = 0; i < 키워드들.length; i++) {
    var 매칭 = 키워드들[i].match(/(\d+)연동/);
    if (매칭) {
      연동숫자 = 매칭[1];
      break;
    }
  }

  var 상위코드 = "";
  var 상위맵 = {
    "발포": "B",
    "방염": "F",
    "비방염": "N",
    "알루미늄": "A"
  };

  for (var key in 상위맵) {
    if (키워드들.indexOf(key) !== -1) {
      상위코드 = 상위맵[key];
      break;
    }
  }

  if (연동숫자 && (상위코드 === "F" || 상위코드 === "N" || 상위코드 === "A")) {
    return 상위코드 + 연동숫자 + "C";
  }

  var 하위코드 = "";
  var 하위맵 = {
    "슬림": "S",
    "와이드": "W",
    "분리": "D",
    "히든": "H",
    "미서기": "L"
  };

  for (var key in 하위맵) {
    if (키워드들.indexOf(key) !== -1) {
      하위코드 += 하위맵[key];
    }
  }

  return 상위코드 + 하위코드;
}

/**
 * 규격 코드 생성
 */
function 규격코드_생성(as값, at값, av값, aq값) {
  var as숫자 = 추출_숫자_from문자열(as값);
  var at숫자 = at값 ? at값.toString() : "";
  var av숫자 = av값 ? av값.toString() : "";

  var 식기표시 = "";
  if (aq값) {
    var aq문자열 = aq값.toString();
    if (aq문자열.includes("3방")) {
      식기표시 = "N";
    } else if (aq문자열.includes("4방")) {
      식기표시 = "Y";
    }
  }

  return as숫자 + at숫자 + av숫자 + 식기표시;
}

// ============================================
// [신규] shop.gs 이식 함수들 (문짝 로직)
// ============================================

/**
 * 품목 타입 구분 (shop.gs: classifyTarget)
 */
function 구분_품목타입(itemName, row) {
  // 1. [강제 로직] 행 번호 기반 강제 분류
  if (row) {
     if (row >= 22 && row <= 34) {
        return 'DOOR';
     }
  }

  var itemStr = itemName ? itemName.toString().trim() : "";
  
  // 문틀 키워드
  var frameKeywords = ['문틀', '발포', '분리형', '스토퍼'];
  var hasFrame = frameKeywords.some(function(kw) { return itemStr.includes(kw); });
  
  // 문짝 키워드
  var doorKeywords = ['문짝', 'ABS', '도어', 'M/D', '민무늬', '탈공', '미서기', '미닫이'];
  var hasDoor = doorKeywords.some(function(kw) { return itemStr.includes(kw); });
  
  // 문짝 패턴 (TA, YS, YA 등) - 대소문자 무시
  var doorPatterns = /YS-[A-Z0-9]+|YA-[A-Z0-9]+|YAT-[A-Z0-9]+|EZ-[A-Z0-9]+|LS-[A-Z0-9]+|YM-[A-Z0-9]+|YAL-[A-Z0-9]+|YV-[A-Z0-9]+|YFL-[A-Z0-9]+|SW-[A-Z0-9]+|TD-[A-Z0-9]+|SL-[A-Z0-9]+|TA-[A-Z0-9]+/i;
  var hasDoorPattern = doorPatterns.test(itemStr);
  
  var hasYeondong = /\d+연동/.test(itemStr);
  
  // 우선순위: 문틀 > 문짝
  if (hasFrame) {
    return 'FRAME';
  }
  
  if (hasDoor || hasDoorPattern || hasYeondong) {
    return 'DOOR';
  }
  
  return 'NONE';
}

/**
 * 모델코드 생성 (shop.gs: generateModelCode)
 * 예: TA-04 -> TA04
 */
function 모델코드_생성(itemName) {
  var itemStr = itemName.toString().trim();
  
  // 영문-패턴 -> 영문+숫자만 추출 (한글 제거)
  // 예: TA-04 -> TA04
  var patternMatch = itemStr.match(/([A-Z]+)-([A-Z0-9]+)/);
  
  if (patternMatch) {
    var prefix = patternMatch[1];  // 영문 부분
    var suffix = patternMatch[2];  // 하이픈 뒤 부분
    
    // 한글이 나타나면 그 전까지만 추출
    var hangulIndex = suffix.search(/[가-힣]/);
    if (hangulIndex !== -1) {
      suffix = suffix.substring(0, hangulIndex);
    }
    
    return prefix + suffix;
  }
  
  // 2순위: 탈공 -> 탈
  if (itemStr.includes('탈공')) {
    return '탈';
  }
  
  // 3순위: M/D + 민무늬 -> MD
  if (itemStr.includes('M/D') && itemStr.includes('민무늬')) {
    return 'MD';
  }
  
  // 4순위: *도어 -> 도어 앞 문자
  var doorMatch = itemStr.match(/(\S+)도어/);
  if (doorMatch && doorMatch[1]) {
    return doorMatch[1];
  }
  
  return '';
}

/**
 * 문짝용 품명 전처리 (shop.gs: preprocessItemNameForProductName)
 */
function 품명_전처리_문짝(itemName, spec) {
  var itemStr = itemName.toString().trim();
  
  // "문틀" 삭제
  itemStr = itemStr.replace(/문틀/g, '').trim();
  
  // "(식기X)" 또는 "(식기O)" 삭제
  itemStr = itemStr.replace(/\(식기[XO]\)/g, '').trim();
  
  // 규격 첫번째 숫자 + 붙어있는 문자 패턴 제거
  var specStr = spec ? spec.toString().trim() : "";
  var firstNumberMatch = specStr.match(/^(\d+)/);
  
  if (firstNumberMatch) {
    var firstNumber = firstNumberMatch[1];
    // 예: "35" -> "35바" 등 제거
    var numberPatternRegex = new RegExp(firstNumber + '[가-힣]+', 'g');
    itemStr = itemStr.replace(numberPatternRegex, '').trim();
  }
  
  return itemStr;
}

// ============================================
// 3. 전체 실행 (단가계산 + 코드생성)
// ============================================

/**
 * 전체 실행
 */
function 전체_실행() {
  Logger.log("\n");
  Logger.log("╔════════════════════════════════════════╗");
  Logger.log("║       🚀 전체 실행 시작                ║");
  Logger.log("╚════════════════════════════════════════╝");
  Logger.log("");

  var 단가계산_성공 = false;
  var 단가계산_결과 = { 성공: 0, 실패: 0 };
  var 코드생성_성공 = false;
  var 코드생성_결과 = { 성공: 0, 실패: 0 };
  var 오류메시지 = [];

  // STEP 1: 단가계산
  try {
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("📌 STEP 1: 단가계산 실행");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

    var 결과1 = 계산_영림발주서_가격_내부();

    if (결과1) {
      단가계산_성공 = true;
      단가계산_결과 = 결과1;
      Logger.log("\n✅ STEP 1 완료: 단가계산 성공");
      Logger.log("   성공: " + 결과1.성공 + "개, 실패: " + 결과1.실패 + "개");
    }

  } catch (e) {
    Logger.log("\n❌ STEP 1 오류: " + e.message);
    Logger.log("   스택: " + e.stack);
    오류메시지.push("단가계산 오류: " + e.message);
  }

  // STEP 2: 코드생성
  try {
    Logger.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("📌 STEP 2: 코드생성 실행");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

    var 결과2 = 생성_품목코드_문틀_내부();

    if (결과2) {
      코드생성_성공 = true;
      코드생성_결과 = 결과2;
      Logger.log("\n✅ STEP 2 완료: 코드생성 성공");
      Logger.log("   성공: " + 결과2.성공 + "개, 실패: " + 결과2.실패 + "개");
    }

  } catch (e) {
    Logger.log("\n❌ STEP 2 오류: " + e.message);
    Logger.log("   스택: " + e.stack);
    오류메시지.push("코드생성 오류: " + e.message);
  }

  // 최종 결과
  Logger.log("\n");
  Logger.log("╔════════════════════════════════════════╗");
  Logger.log("║       ✅ 전체 실행 완료                ║");
  Logger.log("╚════════════════════════════════════════╝");
  Logger.log("");
  Logger.log("📊 실행 결과:");
  Logger.log("  1. 단가계산: " + (단가계산_성공 ? "✅ 성공" : "❌ 실패"));
  if (단가계산_성공) {
    Logger.log("     성공: " + 단가계산_결과.성공 + "개, 실패: " + 단가계산_결과.실패 + "개");
  }
  Logger.log("  2. 코드생성: " + (코드생성_성공 ? "✅ 성공" : "❌ 실패"));
  if (코드생성_성공) {
    Logger.log("     성공: " + 코드생성_결과.성공 + "개, 실패: " + 코드생성_결과.실패 + "개");
  }

  if (오류메시지.length > 0) {
    Logger.log("\n⚠️ 오류 발생:");
    for (var i = 0; i < 오류메시지.length; i++) {
      Logger.log("  • " + 오류메시지[i]);
    }
  }

  Logger.log("\n로그 확인: 보기 > 로그");
}

// ============================================
// 4. 버튼 관리
// ============================================

/**
 * 시트에 버튼 만들기
 */
function 시트에_버튼_만들기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 시트 = ss.getSheetByName("영림발주서");

  if (!시트) {
    SpreadsheetApp.getUi().alert('오류: "영림발주서" 시트를 찾을 수 없습니다.');
    return;
  }

  var 기존버튼 = 시트_버튼_찾기();
  if (기존버튼) {
    var 응답 = SpreadsheetApp.getUi().alert(
      '버튼이 이미 존재합니다',
      '기존 버튼을 삭제하고 새로 만들까요?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (응답 === SpreadsheetApp.getUi().Button.YES) {
      시트_버튼_삭제();
    } else {
      return;
    }
  }

  var 셀범위 = 시트.getRange("BC1:BE1");
  셀범위.merge();
  셀범위.setHorizontalAlignment("center");
  셀범위.setVerticalAlignment("middle");
  셀범위.setValue("📦 품목코드 생성 (BC1 클릭 또는 체크박스 사용)");
  셀범위.setBackground("#4285f4");
  셀범위.setFontColor("#ffffff");
  셀범위.setFontWeight("bold");
  셀범위.setFontSize(11);

  var 체크박스셀 = 시트.getRange("BC2");
  체크박스셀.insertCheckboxes();

  var 설명셀 = 시트.getRange("BD2:BE2");
  설명셀.merge();
  설명셀.setValue("← 체크하면 자동 실행");
  설명셀.setFontSize(10);
  설명셀.setFontColor("#666666");

  SpreadsheetApp.getUi().alert(
    '✅ 실행 버튼 생성 완료!\n\n' +
    '사용 방법:\n' +
    '1. BC2 체크박스 체크 → 자동 실행\n' +
    '2. 또는 메뉴: 🔧 영림발주서 > 📦 코드생성\n\n' +
    '* BC1 셀은 버튼 표시용입니다.\n' +
    '* BC2 체크박스를 사용하세요!'
  );

  Logger.log("✅ 실행 버튼 생성 완료 (BC1~BE1, BC2 체크박스)");
}

/**
 * 시트에서 버튼 찾기
 */
function 시트_버튼_찾기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 시트 = ss.getSheetByName("영림발주서");

  if (!시트) return null;

  try {
    var 셀 = 시트.getRange("BC1");
    var 값 = 셀.getValue();

    if (값 && 값.toString().includes("품목코드 생성")) {
      return 셀;
    }
  } catch (e) {
    Logger.log("버튼 찾기 오류: " + e.message);
  }

  return null;
}

/**
 * 실행 버튼 삭제
 */
function 시트_버튼_삭제() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 시트 = ss.getSheetByName("영림발주서");

  if (!시트) {
    SpreadsheetApp.getUi().alert('오류: "영림발주서" 시트를 찾을 수 없습니다.');
    return;
  }

  try {
    var 셀범위 = 시트.getRange("BC1:BE1");
    if (셀범위.isPartOfMerge()) {
      셀범위.breakApart();
    }
    셀범위.clear();
    셀범위.setBackground(null);
    셀범위.setFontColor(null);
    셀범위.setFontWeight("normal");
    셀범위.setFontSize(10);

    var 체크박스셀 = 시트.getRange("BC2");
    체크박스셀.clear();

    var 설명셀 = 시트.getRange("BD2:BE2");
    if (설명셀.isPartOfMerge()) {
      설명셀.breakApart();
    }
    설명셀.clear();

    SpreadsheetApp.getUi().alert('✅ 실행 버튼이 삭제되었습니다.');
    Logger.log("✅ 실행 버튼 삭제 완료");

  } catch (e) {
    SpreadsheetApp.getUi().alert('오류: ' + e.message);
    Logger.log("❌ 버튼 삭제 오류: " + e.message);
  }
}



/**
 * [디버그용] 저장된 색상 매핑 데이터 확인
 */
function debugColorSync() {
  try {
    var jsonMap = PropertiesService.getScriptProperties().getProperty("COLOR_MAP");
    if (!jsonMap) {
      SpreadsheetApp.getUi().alert("❌ 저장된 데이터가 없습니다.\n'🔄 색상 데이터 업데이트'를 먼저 실행해주세요.");
      return;
    }
    
    var map = JSON.parse(jsonMap);
    var keys = Object.keys(map);
    var sample = keys.slice(0, 10).map(function(k) { return k + " -> " + map[k]; }).join("\n");
    
    var msg = "✅ 데이터 확인됨 (총 " + keys.length + "개)\n\n[샘플 10개]\n" + sample;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("오류: " + e.message);
  }
}

/**
 * 2. 체크박스 및 색상 자동 연동 처리
 */
/**
 * 2. 체크박스 및 색상 자동 연동 처리
 * 통합된 onEdit 함수 (AX, BA, AQ 자동입력 모두 포함)
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    var val = range.getValue();
    
    // [디버그] 진입 확인
    Logger.log("onEdit 진입: 시트=" + sheet.getName() + ", 행=" + row + ", 열=" + col + ", 값=" + val);

    // 영림발주서 시트에서만 동작
    if (sheet.getName() !== "영림발주서") {
       Logger.log(">> 시트 불일치로 종료");
       return;
    }

    // A. 체크박스 실행 (BC2)
    if (range.getA1Notation() === "BC2" && val === true) {
       e.source.toast("🚀 품목코드 생성 시작...", "알림");
       생성_품목코드_문틀();
       range.setValue(false);
       e.source.toast("✅ 품목코드 생성 완료", "알림");
       return; 
    }

    // B. 메모 초기화 (재계산 트리거용)
    // 22~34행의 A열(1) 또는 AW열(49) 수정 시 → A열의 노트(계산완료마킹) 초기화
    if (row >= 22 && row <= 34) {
       if (col === 1 || col === 49) { 
          sheet.getRange(row, 1).clearNote();
       }
    }
    
    // C. AW열(49) 입력 시 자동완성 로직 통합
    if (col === 49) {
       var inputKey = val ? val.toString().trim() : "";
       if (!inputKey) return; // 값이 없으면 중단

       var props = PropertiesService.getScriptProperties();
       var axValue = "";

       // 1. AX열 색상코드 자동완성 (범위: 12행 ~ 35행)
       if (row >= 12 && row <= 35) {
          Logger.log("[onEdit] AX 자동완성 진입: H" + row + " (값: " + inputKey + ")");
          
          var colorMapJson = props.getProperty("COLOR_MAP");
          if (colorMapJson) {
             var map = JSON.parse(colorMapJson);
             var result = map[inputKey];
             
             // 1차 검색 실패 시: 공백 제거 후 재검색
             if (!result) {
                var keyNoSpace = inputKey.replace(/\s+/g, '');
                result = map[keyNoSpace];
             }
             
             // 부분 일치 검색 (정확한 매칭이 없을 때만)
             if (!result) {
                for (var key in map) {
                  if (key.includes(inputKey) || inputKey.includes(key)) {
                    result = map[key];
                    break;
                  }
                }
             }
             
             if (result) {
               sheet.getRange(row, 50).setValue(result); // AX열 (50)
               // e.source.toast("AX" + row + ": " + result, "색상 자동완성");
               axValue = result; // 이후 로직에서 사용
             }
          }
       }
       
       // 2. BA열 가스켓/상태 자동완성 (범위: 12행 ~ 20행)
       if (row >= 12 && row <= 20) {
           var gasketMapJson = props.getProperty("GASKET_COLOR_MAP");
           if (gasketMapJson) {
              var infoMap = JSON.parse(gasketMapJson);
              var targetKeywords = [inputKey];
              if (axValue) targetKeywords.push(axValue); // AX값도 검색 키워드에 추가
              
              var foundInfo = null;
              var foundGroup = "";
              
              var groupNames = ['M', 'P', 'S'];
              var bestMatch = null;
              var bestMatchKey = "";
              var bestMatchGroup = "";
              
              for (var gi = 0; gi < groupNames.length; gi++) {
                 var groupName = groupNames[gi];
                 var groupMap = infoMap[groupName];
                 if (!groupMap) continue;
                 
                 for (var dbKey in groupMap) {
                    for (var k = 0; k < targetKeywords.length; k++) {
                       var uKey = targetKeywords[k];
                       
                       // 정확 일치
                       if (dbKey === uKey) {
                          foundInfo = groupMap[dbKey];
                          foundGroup = groupName;
                          break;
                       }
                       // 부분 일치 (가장 긴 매칭 우선)
                       if (dbKey.includes(uKey) || uKey.includes(dbKey)) {
                          if (!bestMatch || dbKey.length > bestMatchKey.length) {
                             bestMatch = groupMap[dbKey];
                             bestMatchKey = dbKey;
                             bestMatchGroup = groupName;
                          }
                       }
                    }
                    if (foundInfo) break;
                 }
                 if (foundInfo) break;
              }
              
              if (!foundInfo && bestMatch) {
                 foundInfo = bestMatch;
                 foundGroup = bestMatchGroup;
              }
              
              if (foundInfo) {
                 var baCell = sheet.getRange(row, 53); // BA열
                 if (foundInfo.isDiscontinued) {
                    baCell.setValue(foundInfo.status);
                    e.source.toast("BA" + row + ": " + foundInfo.status + " (그룹: " + foundGroup + ")", "정보");
                 } else {
                    baCell.setValue(foundInfo.gasketColor);
                    e.source.toast("BA" + row + ": " + foundInfo.gasketColor + " (그룹: " + foundGroup + ") [5,500원 추가]", "성공");
                 }
              }
           }
       }
       
       // 3. AQ열 문짝 방향 자동완성 (범위: 22행 ~ 34행)
       if (row >= 22 && row <= 34) {
           var avValue = sheet.getRange(row, 48).getValue(); // AV열
           var avNum = Number(avValue) || 0;
           var aqCell = sheet.getRange(row, 43); // AQ열
           
           if (avNum <= 2165) {
               // AQ 빈칸 처리
               aqCell.setValue("");
           } else {
               // AV >= 2166: AG/AH 매칭 확인
               var testSheet = e.source.getSheetByName("테스트");
               if (testSheet) {
                   var agData = testSheet.getRange("AG1:AG300").getValues();
                   var ahData = testSheet.getRange("AH1:AH300").getValues();
                   
                   var searchKeys = [inputKey.toUpperCase()]; // AW값
                   if (axValue) searchKeys.push(axValue.toUpperCase()); // AX값
                   
                   var matchedAH = false;
                   var matchedAG = false;
                   
                   // AH 열 매칭 (Y)
                   for (var i = 0; i < ahData.length; i++) {
                      var cellUpper = ahData[i][0] ? ahData[i][0].toString().toUpperCase() : "";
                      if (!cellUpper) continue;
                      for (var k = 0; k < searchKeys.length; k++) {
                         if (cellUpper.includes(searchKeys[k])) {
                            matchedAH = true; break;
                         }
                      }
                      if (matchedAH) break;
                   }
                   
                   // AG 열 매칭 (N) - AH 매칭 실패 시
                   if (!matchedAH) {
                       for (var i = 0; i < agData.length; i++) {
                          var cellUpper = agData[i][0] ? agData[i][0].toString().toUpperCase() : "";
                          if (!cellUpper) continue;
                          for (var k = 0; k < searchKeys.length; k++) {
                             if (cellUpper.includes(searchKeys[k])) {
                                matchedAG = true; break;
                             }
                          }
                          if (matchedAG) break;
                       }
                   }
                   
                   if (matchedAH) {
                      aqCell.setValue("Y");
                      e.source.toast("AQ: Y (자동입력)", "성공");
                   } else if (matchedAG) {
                      aqCell.setValue("N");
                      e.source.toast("AQ: N (자동입력)", "성공");
                   }
               }
           }
       }
    } // End if (col === 49)

  } catch (error) {
    Logger.log("❌ onEdit 오류: " + error.message);
  }
}

// ... (중략: setDropdowns 등 기존 함수들) ...

/**
 * 색상 데이터 매핑 업데이트 (외부 시트 → 스크립트 속성 저장)
 * 메뉴에서 실행하거나, 주기적으로 실행 필요
 */
/**
 * 색상 데이터 및 품목 정보 업데이트
 * 1. 색상 매핑 (외부 V1:Y150) -> COLOR_MAP
 * 2. 품목 정보 (로컬 '테스트' M1:U300) -> PRODUCT_INFO_MAP
 */
// (구 updateColorCodeMap 삭제됨)

function setDropdowns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('영림발주서');

  for (var row = 12; row <= 30; row++) {
    var cell = sheet.getRange('AS' + row);
    var column = String.fromCharCode(67 + row - 12); // C, D, E, ...
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('필터').getRange(column + ':' + column))
      .build();
    cell.setDataValidation(rule);
  }

  SpreadsheetApp.getUi().alert('드롭다운 설정 완료!');
}

/**
 * 외부 시트 접근 및 A43 값 확인
 */
function checkExternalSheetAccess() {
  var sheetId = '1bd7fjyjumFC0RZU56VO3Qdr4Z3gwbS5T04smwe6FKSA';
  
  Logger.log("🔍 외부 시트 접근 및 A43 값 확인...");
  
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    Logger.log("✅ 접근 성공: " + ss.getName());
    
    // 첫 번째 시트의 A43 값 읽기
    var sheet = ss.getSheets()[0];
    var value = sheet.getRange("A43").getValue();
    
    Logger.log("----------------------------------------");
    Logger.log("📄 시트명: " + sheet.getName());
    Logger.log("📍 A43 값: " + value);
    Logger.log("----------------------------------------");
    
    SpreadsheetApp.getUi().alert("성공!\n\n시트명: " + sheet.getName() + "\nA43 값: " + value);
    
  } catch (e) {
    Logger.log("❌ 오류 발생: " + e.message);
    SpreadsheetApp.getUi().alert("오류 발생: " + e.message);
  }
}

/**
 * 외부 시트의 배경색 분석 및 데이터 찾기
 */
function analyzeExternalSheetColors() {
  var sheetId = '1bd7fjyjumFC0RZU56VO3Qdr4Z3gwbS5T04smwe6FKSA';
  
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheets()[0]; // 첫 번째 시트
    
    // 분석할 범위 설정 (예: A1에서 F50까지)
    var maxRow = 50;
    var maxCol = 10; // J열까지
    var range = sheet.getRange(1, 1, maxRow, maxCol);
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();
    
    var colorReport = {};
    var foundCount = 0;
    
    // 데이터 스캔
    for (var i = 0; i < maxRow; i++) {
      for (var j = 0; j < maxCol; j++) {
        var color = backgrounds[i][j];
        var value = values[i][j];
        
        // 흰색(#ffffff)이 아니고 값이 있는 경우만 체크
        if (color !== '#ffffff' && value !== "") {
          if (!colorReport[color]) {
            colorReport[color] = [];
          }
          // 처음 5개까지만 값 저장 (예시용)
          if (colorReport[color].length < 5) {
             var cellAddress = String.fromCharCode(65 + j) + (i + 1);
             colorReport[color].push("[" + cellAddress + "] " + value);
          }
          foundCount++;
        }
      }
    }
    
    // 결과 메시지 생성
    var message = "🎨 색상 분석 결과 (" + foundCount + "개 발견)\n\n";
    
    if (Object.keys(colorReport).length === 0) {
      message += "흰색(#ffffff) 이외의 배경색을 가진 셀을 찾지 못했습니다.";
    } else {
      for (var code in colorReport) {
        message += "■ 색상코드: " + code + "\n";
        message += "   찾은 값들: " + colorReport[code].join(", ") + "\n\n";
      }
    }
    
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("오류: " + e.message);
  }
}

/**
 * [테스트용] 색상 데이터 파싱 로직 검증
 */
function testColorParsingLogic() {
  var sheetId = '1bd7fjyjumFC0RZU56VO3Qdr4Z3gwbS5T04smwe6FKSA';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  
  // V1:Z150 데이터 읽기
  var rangeValues = sheet.getRange("V1:Z150").getValues();
  var pairs = [];
  
  Logger.log("🔍 파싱 로직 테스트 시작 (V1:Z150)...");
  
  for (var r = 0; r < rangeValues.length; r++) {
    for (var c = 0; c < rangeValues[r].length; c++) {
      var cellVal = rangeValues[r][c];
      if (!cellVal || cellVal.toString().trim() === "") continue;
      
      var text = cellVal.toString().trim();
      var code1 = "", code2 = "";
      
      // 1. 괄호가 있는 경우: '영림숫자' + '괄호 안 내용'
      if (text.includes("(")) {
        var matchNumbers = text.match(/(영림\d+)/); // 영림+숫자 추출
        var matchParens = text.match(/\(([^)]+)\)/); // 괄호 안 내용 추출
        
        if (matchNumbers && matchParens) {
          code1 = matchNumbers[1];
          code2 = matchParens[1];
        } else {
             Logger.log("⚠️ 괄호 패턴 매칭 실패: " + text);
             continue; // 매칭 안되면 건너뜀
        }
      } 
      // 2. 괄호가 없고 공백이 있는 경우: '첫단어' + '나머지'
      else if (text.includes(" ")) {
        var parts = text.split(" ");
        code1 = parts[0];
        code2 = parts.slice(1).join(" ");
      } 
      // 3. 단독 코드인 경우 (무시)
      else {
        // Logger.log("  Pass (단독): " + text);
        continue;
      }
      
      if (code1 && code2) {
        pairs.push(code1 + " ↔ " + code2 + " (원본: " + text + ")");
      }
    }
  }
  
  Logger.log("✅ 추출된 쌍 (" + pairs.length + "개):");
  Logger.log(pairs.slice(0, 30).join("\n")); // 처음 30개만 로그 출력
  if (pairs.length > 30) Logger.log("...외 " + (pairs.length - 30) + "개 더 있음");
  
  SpreadsheetApp.getUi().alert("분석 완료! 로그를 확인하세요.\n추출된 쌍 개수: " + pairs.length);
}

/**
 * 입력값 및 출력값 초기화
 * 범위: 12행 ~ 35행
 * 대상: A열(금액/메모), AP열~BG열(입력데이터)
 */
function 초기화_영림발주서() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("영림발주서");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ '영림발주서' 시트를 찾을 수 없습니다.");
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('⚠️ 초기화 확인', '정말 입력/출력 데이터를 모두 지우시겠습니까?\n(12행~35행: A열, AP열~BG열)', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    // 1. A열 초기화 (값 + 메모) - 12~20, 22~34 (21행 헤더 제외)
    sheet.getRange("A12:A20").clearContent().clearNote();
    sheet.getRange("A22:A34").clearContent().clearNote();
    
    // 2. AR열 ~ BF열 초기화 (값) - 12~20, 22~34 (21행 헤더 제외)
    // AR(44) ~ BF(58)
    sheet.getRange("AR12:BF20").clearContent();
    sheet.getRange("AR22:BF34").clearContent();
    
    ss.toast("모든 데이터가 초기화되었습니다.", "완료");
  } else {
    ss.toast("취소되었습니다.", "알림");
  }
}
/**
 * onEdit 트리거 함수
 */

// Removed duplicate onEdit at 2253



/**
 * 가스켓 색상 데이터 업데이트
 * 테스트 시트의 M~U열 데이터를 읽어서 그룹별로 저장
 */
function updateGasketColorMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testSheet = ss.getSheetByName("테스트");
  
  if (!testSheet) {
    SpreadsheetApp.getUi().alert("❌ '테스트' 시트를 찾을 수 없습니다.");
    return;
  }
  
  var rangeData = testSheet.getRange("M1:U300").getValues();
  var gasketMapM = {};
  var gasketMapP = {};
  var gasketMapS = {};
  
  var groups = [
    {nameIdx: 0, gasketIdx: 1, statusIdx: 2, group: 'M', map: gasketMapM},
    {nameIdx: 3, gasketIdx: 4, statusIdx: 5, group: 'P', map: gasketMapP},
    {nameIdx: 6, gasketIdx: 7, statusIdx: 8, group: 'S', map: gasketMapS}
  ];
  
  var totalCount = 0;
  
  for (var i = 0; i < rangeData.length; i++) {
    var row = rangeData[i];
    for (var g = 0; g < groups.length; g++) {
      var grp = groups[g];
      var colorName = row[grp.nameIdx];
      var gasketColor = row[grp.gasketIdx];
      var status = row[grp.statusIdx];
      if (colorName && colorName.toString().trim() !== "") {
        var key = colorName.toString().trim();
        var gasket = gasketColor ? gasketColor.toString().trim() : "";
        var statusText = status ? status.toString().trim() : "";
        grp.map[key] = {
          gasketColor: gasket, 
          status: statusText, 
          isDiscontinued: (statusText === "단종" || statusText === "단종예정"), 
          group: grp.group
        };
        totalCount++;
      }
    }
  }
  
  var gasketMap = {M: gasketMapM, P: gasketMapP, S: gasketMapS};
  PropertiesService.getScriptProperties().setProperty("GASKET_COLOR_MAP", JSON.stringify(gasketMap));
  SpreadsheetApp.getUi().alert("✅ 업데이트 완료! 총 " + totalCount + "건");
}


/**
 * 디버그: 특정 키워드 검색 과정 확인
 */
function debugSearchProcess() {
  var searchKey = "영림116"; // 검색할 키워드
  
  var props = PropertiesService.getScriptProperties();
  var jsonMap = props.getProperty("GASKET_COLOR_MAP");
  
  if (!jsonMap) {
    SpreadsheetApp.getUi().alert("데이터 없음! updateGasketColorMap() 먼저 실행하세요.");
    return;
  }
  
  var gasketMap = JSON.parse(jsonMap);
  var log = "=== 검색 디버그: " + searchKey + " ===\n\n";
  
  // 각 그룹별 데이터 확인
  var groupNames = ['M', 'P', 'S'];
  
  for (var gi = 0; gi < groupNames.length; gi++) {
    var groupName = groupNames[gi];
    var groupMap = gasketMap[groupName];
    
    log += "📁 " + groupName + " 그룹:\n";
    
    if (!groupMap) {
      log += "  (데이터 없음)\n";
      continue;
    }
    
    // searchKey가 포함된 항목 찾기
    var found = false;
    for (var key in groupMap) {
      if (key.includes(searchKey) || searchKey.includes(key)) {
        var info = groupMap[key];
        log += "  ✅ 매칭! 키: " + key + "\n";
        log += "     gasketColor: " + info.gasketColor + "\n";
        log += "     status: " + info.status + "\n";
        log += "     isDiscontinued: " + info.isDiscontinued + "\n";
        found = true;
      }
    }
    
    if (!found) {
      log += "  (매칭 항목 없음)\n";
    }
    
    log += "\n";
  }
  
  Logger.log(log);
  SpreadsheetApp.getUi().alert(log);
}


/**
 * 직접 테스트: 영림116 검색
 */
function testSearch영림116() {
  var searchKey = "영림116";
  
  var props = PropertiesService.getScriptProperties();
  var jsonMap = props.getProperty("GASKET_COLOR_MAP");
  
  if (!jsonMap) {
    SpreadsheetApp.getUi().alert("GASKET_COLOR_MAP 없음! updateGasketColorMap() 실행 필요");
    return;
  }
  
  var infoMap = JSON.parse(jsonMap);
  var targetKeywords = [searchKey];
  
  var foundInfo = null;
  var foundGroup = "";
  
  var groupNames = ['M', 'P', 'S'];
  var bestMatch = null;
  var bestMatchKey = "";
  var bestMatchGroup = "";
  
  for (var gi = 0; gi < groupNames.length; gi++) {
    var groupName = groupNames[gi];
    var groupMap = infoMap[groupName];
    
    if (!groupMap) continue;
    
    for (var dbKey in groupMap) {
      for (var k = 0; k < targetKeywords.length; k++) {
        var uKey = targetKeywords[k];
        
        // 1. 정확히 일치하면 즉시 사용
        if (dbKey === uKey) {
          foundInfo = groupMap[dbKey];
          foundGroup = groupName;
          break;
        }
        
        // 2. 부분 일치: 더 긴 키가 우선
        if (dbKey.includes(uKey) || uKey.includes(dbKey)) {
          if (!bestMatch || dbKey.length > bestMatchKey.length) {
            bestMatch = groupMap[dbKey];
            bestMatchKey = dbKey;
            bestMatchGroup = groupName;
          }
        }
      }
      if (foundInfo) break;
    }
    if (foundInfo) break;
  }
  
  // 정확히 일치하는 것이 없으면 가장 긴 부분 일치 사용
  if (!foundInfo && bestMatch) {
    foundInfo = bestMatch;
    foundGroup = bestMatchGroup;
  }
  
  var msg = "=== 검색 결과: " + searchKey + " ===\n\n";
  
  if (foundInfo) {
    msg += "✅ 매칭 성공!\n";
    msg += "그룹: " + foundGroup + "\n";
    msg += "gasketColor: " + foundInfo.gasketColor + "\n";
    msg += "status: " + foundInfo.status + "\n";
    msg += "isDiscontinued: " + foundInfo.isDiscontinued + "\n\n";
    
    if (foundInfo.isDiscontinued) {
      msg += "→ BA 출력: " + foundInfo.status + " (단종)";
    } else {
      msg += "→ BA 출력: " + foundInfo.gasketColor + " [5,500원 추가]";
    }
  } else {
    msg += "❌ 매칭 실패";
  }
  
  SpreadsheetApp.getUi().alert(msg);
}


/**
 * onEdit 트리거 - AW열 입력 시 BA열 자동 입력
 */

// Removed duplicate onEdit at 2513


/**
 * 색상코드 매핑 업데이트
 * 테스트 시트에서 색상 매핑 데이터를 읽어 스크립트 속성에 저장
 */
function updateColorCodeMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testSheet = ss.getSheetByName("테스트");
  
  if (!testSheet) {
    SpreadsheetApp.getUi().alert("❌ '테스트' 시트를 찾을 수 없습니다.");
    return;
  }
  
  // V1:Z300 영역 스캔 (넓은 범위)
  var data = testSheet.getRange("V1:Z300").getValues();
  
  var colorMap = {};
  var count = 0;
  
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var cellVal = data[r][c];
      if (!cellVal || cellVal.toString().trim() === "") continue;
      
      var text = cellVal.toString().trim();
      var code1 = "", code2 = "";
      
      // 1. 괄호 패턴: "영림116 (중백색)"
      if (text.includes("(") && text.includes(")")) {
         var matchNumbers = text.match(/(영림\s*\d+)/);
         var matchParens = text.match(/\(([^)]+)\)/);
         if (matchNumbers && matchParens) {
            code1 = matchNumbers[1].replace(/\s+/g, ''); 
            code2 = matchParens[1].trim(); 
         }
      }
      // 2. 공백 패턴: "PS010 중백색"
      else if (text.includes(" ")) {
         var parts = text.split(" ");
         if (parts.length >= 2) {
            code1 = parts[0].trim();
            code2 = parts.slice(1).join(" ").trim();
         }
      }
      
      if (code1 && code2) {
         colorMap[code1] = code2; 
         count++;
      }
    }
  }
  

  PropertiesService.getScriptProperties().setProperty("COLOR_MAP", JSON.stringify(colorMap));
  
  var msg = "✅ 색상코드 업데이트 완료!\n- 범위: 테스트 시트 V1:Z300\n- 추출된 항목: " + count + "개";
  if (count === 0) msg += "\n⚠️ 데이터를 찾지 못했습니다.\nV~Z열에 '영림116 (중백색)' 또는 '코드 색상' 형식의 데이터가 필요합니다.";
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * [디버그] COLOR_MAP 데이터 및 검색 로직 확인
 */
function debugColorMapCheck() {
  var props = PropertiesService.getScriptProperties();
  var jsonMap = props.getProperty("COLOR_MAP");
  
  if (!jsonMap) {
    var msg = "❌ COLOR_MAP 데이터가 없습니다.\n'updateColorCodeMap()' 함수를 실행하여 데이터를 업데이트해주세요.";
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }
  
  var map = JSON.parse(jsonMap);
  var keys = Object.keys(map);
  var count = keys.length;
  
  var log = "=== COLOR_MAP 점검 ===\n";
  log += "총 데이터 수: " + count + "개\n";
  
  if (count > 0) {
    log += "첫 번째 데이터: " + keys[0] + " -> " + map[keys[0]] + "\n";
  }
  
  // 테스트 검색
  var testKey = "영림116"; // 테스트해볼 키워드
  var result = map[testKey];
  
  log += "\n[검색 테스트: '" + testKey + "']\n";
  if (result) {
    log += "✅ 정확히 일치: " + result + "\n";
  } else {
    var keyNoSpace = testKey.replace(/\s+/g, '');
    result = map[keyNoSpace];
    if (result) {
       log += "✅ 공백제거 일치: " + result + "\n";
    } else {
       // 부분 일치 확인
       var partial = "";
       for (var k in map) {
         if (k.includes(testKey) || testKey.includes(k)) {
           partial = map[k] + " (Key: " + k + ")";
           break;
         }
       }
       if (partial) {
          log += "✅ 부분 일치: " + partial + "\n";
       } else {
          log += "❌ 검색 실패\n";
       }
    }
  }
  
  Logger.log(log);
  SpreadsheetApp.getUi().alert(log);
}

