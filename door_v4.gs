/**
 * 영림발주서 통합 스크립트 v4 - 양방향 색상 매핑 지원 버전
 * 
 * 기능:
 * 1. 단가계산: A12~A35에 가격 출력
 * 2. 코드생성: BC12~BF35에 품목명/코드 출력
 * 3. 전체실행: 1+2 순차 실행
 * 4. 데이터 업데이트: 색상코드 및 가스켓 정보를 한 번에 동기화 (양방향 매핑 지원)
 * 
 * 메뉴:
 * - 🔄 데이터 업데이트 (색상/가스켓 전체)
 * - 💰 단가계산
 * - 📦 코드생성
 * - 🚀 전체
 * - 🧹 입력/출력 초기화
 */

// ============================================
// 메뉴 및 초기화
// ============================================

/**
 * 시트 열 때 자동 실행 - 메뉴 생성
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 영림발주서 v4')
    .addItem('🔄 데이터 업데이트 (전체 필독)', 'menu_updateAllData')
    .addSeparator()
    .addItem('💰 단가계산', '계산_영림발주서_가격')
    .addItem('📦 코드생성', '생성_품목코드_문틀')
    .addItem('🚀 전체 (단가+코드)', '전체_실행')
    .addSeparator()
    .addItem('🎨 실행 버튼 만들기', '시트에_버튼_만들기')
    .addItem('🗑️ 실행 버튼 삭제', '시트_버튼_삭제')
    .addSeparator()
    .addItem('🧹 입력/출력 초기화', '초기화_영림발주서')
    .addItem('📋 로그 보기 안내', '로그보기')
    .addToUi();
}

/**
 * 전체 데이터 업데이트 (색상 + 가스켓)
 */
function menu_updateAllData() {
  var ui = SpreadsheetApp.getUi();
  ui.showModelessDialog(HtmlService.createHtmlOutput('<p>데이터 업데이트 중입니다... 잠시만 기다려주세요.</p>').setWidth(300).setHeight(100), '🔄 진행 중');
  
  try {
    updateColorCodeMap();
    updateGasketColorMap();
    
    ui.alert('✅ 업데이트 완료', 
             '1. 색상 매핑 데이터 업데이트 완료 (양방향 지원)\n' + 
             '2. 가스켓 색상 데이터 업데이트 완료\n\n' +
             '이제 AW열에 값을 입력하면 실시간으로 AX/BA열이 채워집니다.', 
             ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ 업데이트 오류', e.message, ui.ButtonSet.OK);
  }
}

/**
 * [테스트] 23행만 실제 계산 로직으로 실행하고 결과를 팝업으로 표시
 */
function testRow23Calculation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 발주서시트 = ss.getSheetByName("영림발주서");
  var 단가표시트 = ss.getSheetByName("영림문틀단가표");
  var 테스트시트 = ss.getSheetByName("테스트");
  
  var log = "=== 23행 실제 로직 테스트 ===\n";
  var i = 23; // 고정
  
  // ===== 데이터 로드 =====
  var 단가표데이터 = 단가표시트.getRange("C6:F500").getValues();
  
  var 추가금액정보 = null;
  var 추가금액데이터 = [];
  var 문짝가격데이터 = [];
  
  if (테스트시트) {
     var 전체범위 = 테스트시트.getRange("V3:Z100");
     var 전체값 = 전체범위.getValues();
     추가금액정보 = 전체값[0];
     추가금액데이터 = 전체값.slice(1);
     
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
  
  // ===== 계산 로직 =====
  var 최종가격 = 0;
  var 계산성공 = false;
  
  if (ap값 && ap값.toString().trim() !== "") {
    var 제품타입 = 추출_제품타입(ap값);
    var 공급가 = 찾기_공급가(단가표데이터, 제품타입, as값, aq값);
    if (공급가 !== null) {
      최종가격 = 공급가 * ar값;
      계산성공 = true;
    }
  }
  
  if (!계산성공 && i >= 22 && i <= 34 && a값 !== "" && a값 !== null) {
      최종가격 = Number(a값);
      if (!isNaN(최종가격)) 계산성공 = true;
  }
  
  if (계산성공 && i >= 22 && i <= 34) {
    // AW 매칭
    if (aw값 && aw값.toString().trim() !== "" && 추가금액정보) {
        var keyword = aw값.toString().trim().toUpperCase();
        var matchedCol = -1;
        
        outerLoop:
        for (var r = 0; r < 추가금액데이터.length; r++) {
           for (var c = 0; c < 5; c++) {
              var cellText = 추가금액데이터[r][c] ? 추가금액데이터[r][c].toString().toUpperCase().trim() : "";
              if (cellText && (cellText.includes(keyword) || keyword.includes(cellText))) {
                 matchedCol = c; break outerLoop;
              }
           }
        }
        if (matchedCol !== -1) 최종가격 += (추가금액정보[matchedCol] || 0);
    }
    
    // Door 매칭
    var aqStr = aq값 ? aq값.toString().toUpperCase() : "";
    if (aqStr.includes("Y") && 문짝가격데이터.length > 0) {
       var targetUpper = ((aw값 ? aw값.toString() : "") + " " + (ax값 ? ax값.toString() : "")).toUpperCase();
       for(var d=0; d<문짝가격데이터.length; d++) {
          var entry = 문짝가격데이터[d];
          for(var k=0; k<entry.keywords.length; k++) {
             var kw = entry.keywords[k].toString().toUpperCase().trim();
             if(kw && targetUpper.includes(kw)) {
                최종가격 += entry.price; break;
             }
          }
       }
    }
  }
  
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
  if (sheet.getName() !== "영림발주서") return SpreadsheetApp.getUi().alert("영림발주서 시트에서 실행해주세요.");
  
  var log = "=== 행 " + row + " 가격 계산 추적 ===\n";
  var ap = sheet.getRange("AP" + row).getValue();
  var aw = sheet.getRange("AW" + row).getValue();
  var ax = sheet.getRange("AX" + row).getValue();
  var aq = sheet.getRange("AQ" + row).getValue();
  var a = sheet.getRange("A" + row).getValue();
  
  var testSheet = ss.getSheetByName("테스트");
  if (!testSheet) return;
  var rawPart1 = testSheet.getRange("V3:Z100").getValues();
  var info1 = rawPart1[0];
  var data1 = rawPart1.slice(1);
  
  var keyword = aw ? aw.toString().trim().toUpperCase() : "";
  var part1Price = 0;
  if (keyword) {
    outer: for(var r=0; r<data1.length; r++) {
      for(var c=0; c<5; c++) {
        var cell = data1[r][c] ? data1[r][c].toString().toUpperCase().trim() : "";
        if (cell && (cell.includes(keyword) || keyword.includes(cell))) {
          if (typeof info1[c] === 'number') { part1Price = info1[c]; break outer; }
        }
      }
    }
  }
  
  var part2Price = 0;
  if ((aq ? aq.toString().toUpperCase() : "").includes("Y")) {
    var rawPart2 = testSheet.getRange("AD1:AF100").getValues();
    var target = ((aw?aw.toString():"") + " " + (ax?ax.toString():"")).toUpperCase();
    for(var i=0; i<rawPart2.length; i++) {
       if(!rawPart2[i][0]) continue;
       var kws = rawPart2[i][0].toString().split(',').map(function(s){ return s.trim().toUpperCase(); });
       if (kws.some(function(k){ return k && target.includes(k); })) {
         part2Price = Number(rawPart2[i][2]) || 0; break;
       }
    }
  }
  
  var base = Number(a) || 0;
  log += "결과: " + base + " (기본) + " + part1Price + " (AW추가) + " + part2Price + " (Door추가) = " + (base + part1Price + part2Price);
  SpreadsheetApp.getUi().alert(log);
}

/**
 * 로그 확인 안내
 */
function 로그보기() {
  SpreadsheetApp.getUi().alert('로그 확인 방법', '보기 > 로그 또는 Ctrl+Enter 누르기', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================
// 1. 단가계산 (A열 출력)
// ============================================

function 계산_영림발주서_가격() {
  try {
    var 결과 = 계산_영림발주서_가격_내부();
    SpreadsheetApp.getActiveSpreadsheet().toast("성공: " + 결과.성공 + ", 실패: " + 결과.실패, "✅ 단가계산 완료");
  } catch (e) {
    Logger.log("❌ 단가계산 오류: " + e.message);
  }
}

const CONFIG = {
  SHEET_NAME: "영림발주서",
  TEST_SHEET_NAME: "테스트",
  START_ROW: 12,
  END_ROW: 35,
  COLS: {
    AP: 42, AQ: 43, AR: 44, AS: 45, AT: 46, AW: 47, AX: 48, BA: 51
  }
};

/**
 * 병합된 AT셀 값을 파싱하여 너비/높이를 분리한다.
 * @param {*} atValue - AT셀 값 (예: "880*2090", 880, "", null)
 * @returns {{ width: number, height: number, raw: string }}
 */
function parseAT(atValue) {
  if (!atValue) return { width: 0, height: 0, raw: "" };
  var str = atValue.toString().trim();
  var parts = str.split("*");
  return {
    width:  Number(parts[0]) || 0,
    height: parts.length > 1 ? (Number(parts[parts.length - 1]) || 0) : 0,
    raw:    str
  };
}

function 계산_영림발주서_가격_내부() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 발주서시트 = ss.getSheetByName(CONFIG.SHEET_NAME);
  var 단가표시트 = ss.getSheetByName("영림문틀단가표");
  if (!발주서시트 || !단가표시트) throw new Error('시트를 찾을 수 없습니다.');

  var 단가표데이터 = 단가표시트.getRange("C6:F500").getValues();
  var testData = loadTestSheetData_Optimized(ss);
  
  var startRow = CONFIG.START_ROW;
  var numRows = CONFIG.END_ROW - startRow + 1;
  var aRange = 발주서시트.getRange(startRow, 1, numRows, 1);
  var aValues = aRange.getValues();
  var aNotes = aRange.getNotes();
  var dataValues = 발주서시트.getRange(startRow, CONFIG.COLS.AP, numRows, CONFIG.COLS.BA - CONFIG.COLS.AP + 1).getValues();
  
  var resultPrices = [], resultNotes = [];
  var 성공 = 0, 실패 = 0;

  for (var i = 0; i < numRows; i++) {
    var currentRow = startRow + i;
    var rowData = dataValues[i];
    var ap = rowData[0], aq = rowData[1], ar = rowData[2], as = rowData[3], atRaw = rowData[4],
        aw = rowData[5], ax = rowData[6], ba = rowData[9];
    var parsed = parseAT(atRaw);
    var at = parsed.width, av = parsed.height;
    var curP = aValues[i][0], curN = aNotes[i][0];

    if (!atRaw) { resultPrices.push([curP]); resultNotes.push([curN]); 실패++; continue; }
    if (curP && curN && (curN.includes("✅"))) { resultPrices.push([curP]); resultNotes.push([curN]); 성공++; continue; }
    
    var finalP = 0, success = false, manual = false;
    if (ap) {
        var pType = 추출_제품타입(ap);
        var sP = 찾기_공급가(단가표데이터, pType, as, aq);
        if (sP !== null) { finalP = sP * (Number(ar) || 0); success = true; curN = ""; }
    }
    if (!success && currentRow >= 22 && curP) {
        finalP = Number(curP);
        if (!isNaN(finalP)) { success = true; manual = true; }
    }
    
    if (success) {
        var extra = false;
        if (currentRow < 22 && ba && !["없음","단종","단종예정"].includes(ba.toString().trim())) finalP += 5500;
        if (currentRow >= 22 && aw) {
            var kw = aw.toString().trim().toUpperCase();
            var added = findPriceFromMap_Scan(kw, testData.additionalPriceMap, testData.additionalPriceInfo);
            if (added !== null) { finalP += added; extra = true; }
        }
        if (currentRow >= 22 && (aq?aq.toString().toUpperCase():"").includes("Y")) {
            var target = ((aw?aw.toString():"") + " " + (ax?ax.toString():"")).toUpperCase();
            var doorP = findDoorPrice_Scan(target, testData.doorPriceMap);
            if (doorP > 0 && av >= 2166) finalP += doorP;
        }
        resultPrices.push([finalP]);
        resultNotes.push([(manual && extra) ? "✅추가금반영됨" : curN]);
        성공++;
    } else {
        resultPrices.push([""]); resultNotes.push([""]); 실패++;
    }
  }
  aRange.setValues(resultPrices);
  aRange.setNotes(resultNotes);
  return { 성공: 성공, 실패: 실패 };
}

function loadTestSheetData_Optimized(ss) {
    try {
        var sheet = ss.getSheetByName(CONFIG.TEST_SHEET_NAME);
        if (!sheet) return { additionalPriceInfo: [], additionalPriceMap: {}, doorPriceMap: {} };
        
        var rawRange = sheet.getRange("V3:Z300");
        var raw = rawRange.getValues();
        
        var addMap = {};
        // raw[0] is interpreted as additionalPriceInfo (headers)
        var headers = (raw && raw.length > 0) ? raw[0] : [];
        
        for (var r = 1; r < raw.length; r++) {
           for (var c = 0; c < 5; c++) {
               if (raw[r][c]) {
                   var key = raw[r][c].toString().toUpperCase().trim();
                   if (key) addMap[key] = c;
               }
           }
        }
        
        var doorData = sheet.getRange("AD1:AF50").getValues();
        var doorMap = {};
        for (var i = 0; i < doorData.length; i++) {
            if (doorData[i][0] && doorData[i][2] !== undefined && doorData[i][2] !== "") {
                var price = Number(doorData[i][2]) || 0;
                doorData[i][0].toString().split(',').forEach(function(k){ 
                    var doorKey = k.trim().toUpperCase();
                    if (doorKey) doorMap[doorKey] = price; 
                });
            }
        }
        return { additionalPriceInfo: headers, additionalPriceMap: addMap, doorPriceMap: doorMap };
    } catch (err) {
        Logger.log("Error in loadTestSheetData_Optimized: " + err.message);
        return { additionalPriceInfo: [], additionalPriceMap: {}, doorPriceMap: {} };
    }
}

function findPriceFromMap_Scan(target, map, infoArr) {
    if (!target || !map || !infoArr) return null;
    var idx = map[target];
    if (idx !== undefined && infoArr[idx] !== undefined) return infoArr[idx];
    
    for (var key in map) { 
        if (key && (key.includes(target) || target.includes(key))) {
            var sIdx = map[key];
            if (sIdx !== undefined && infoArr[sIdx] !== undefined) return infoArr[sIdx];
        }
    }
    return null;
}

function findDoorPrice_Scan(target, map) {
    for (var key in map) { if (target.includes(key)) return map[key]; }
    return 0;
}

function 추출_제품타입(ap값) {
  if (!ap값) return "";
  // 모든 브랜드 접두사 패턴 제거 (영림/우딘/예림 + 기호 또는 공백)
  var s = ap값.toString().replace(/^(영림|우딘|예림)[^가-힣a-zA-Z0-9\s]?\s*/, '').trim();
  
  var m = s.match(/(\d+)방/);
  if (m) {
    var res = s.substring(0, s.indexOf(m[0]));
    // 끝에 남은 특수문자나 구분자 제거
    return res.replace(/[^가-힣a-zA-Z0-9\s]$/, '').trim();
  }
  return s;
}

function 정규화_키워드(k) {
  var s = k ? k.toString().trim() : "";
  return s.endsWith("형") ? s.slice(0, -1) : s;
}

function 찾기_공급가(data, type, size, direction) {
  var t = type ? type.toString().trim() : "";
  var s = size ? size.toString().replace(/바$/, '').trim() : "";
  var d = direction ? direction.toString().trim() : "";
  
  if (!t || !s || !d) return null;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowType = row[0] ? row[0].toString().trim().replace(/형/g, '') : "";
    var rowSize = row[1] ? row[1].toString().replace(/바$/, '').trim() : "";
    var rowDir = row[2] ? row[2].toString().trim() : "";
    
    // 타입(제품명), 규격(숫자부분), 인필(방수) 3가지 조건 대조
    if (t.split('ㅣ').every(function(kw){ return rowType.includes(정규화_키워드(kw)); }) && 
        rowSize === s && 
        rowDir === d) {
      return Number(row[3]) || null;
    }
  }
  return null;
}

// ============================================
// 2. 코드생성 (BC~BF열 출력)
// ============================================

function 생성_품목코드_문틀() {
  try {
    var 결과 = 생성_품목코드_문틀_내부();
    SpreadsheetApp.getActiveSpreadsheet().toast("성공: " + 결과.성공 + ", 실패: " + 결과.실패, "✅ 코드생성 완료");
  } catch (e) {
    Logger.log("❌ 코드생성 오류: " + e.message);
  }
}

function 생성_품목코드_문틀_내부() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('시트 없음');
  var start = CONFIG.START_ROW, num = CONFIG.END_ROW - start + 1;
  var data = sheet.getRange(start, CONFIG.COLS.AP, num, CONFIG.COLS.BA - CONFIG.COLS.AP + 1).getValues();
  
  var names = [], codes = [], empty = [], units = [];
  var 성공 = 0, 실패 = 0;

  for (var i = 0; i < num; i++) {
    var row = data[i], rIdx = start + i;
    
    // 조건 1: 병합셀(AT)이 비어 있는 경우
    if (!row[4]) { 
      names.push([""]); codes.push([""]); empty.push([""]); units.push([""]); 
      실패++; continue; 
    }
    
    // 조건 2: 숫자가 특정 임계값(499) 이하인 경우
    var _p = parseAT(row[4]);
    if (Math.max(추출_숫자_from문자열(row[3]), _p.width, _p.height) <= 499) { 
      names.push([""]); codes.push([""]); empty.push([""]); units.push([""]); 
      실패++; continue; 
    }

    try {
      var _p2 = parseAT(row[4]);
      var n = 생성_품목명(row[0], row[5], row[6], row[3], _p2.width, _p2.height, row[1], rIdx);
      var c = 생성_품목코드_NEW(row[0], row[5], row[6], row[3], _p2.width, _p2.height, row[1], rIdx);
      names.push([n]); codes.push([c]); empty.push([""]); units.push([rIdx >= 22 ? "짝" : "틀"]); 성공++;
    } catch (e) {
      names.push([""]); codes.push([""]); empty.push([""]); units.push([""]); 
      실패++;
    }
  }
  
  sheet.getRange(start, 53, num, 1).setValues(names);
  sheet.getRange(start, 54, num, 1).setValues(codes);
  sheet.getRange(start, 55, num, 1).setValues(empty);
  sheet.getRange(start, 56, num, 1).setValues(units);
  return { 성공: 성공, 실패: 실패 };
}

function 추출_숫자_from문자열(v) { return v ? (v.toString().match(/\d+/) ? Number(v.toString().match(/\d+/)[0]) : 0) : 0; }

function 생성_품목명(ap, aw, ax, as, at, av, aq, row) {
  var p = ap ? ap.toString() : "", t = 구분_품목타입(p, row);
  if (t === 'DOOR') {
     var co = p.includes("영림") ? "영림" : "영림";
     var color = 색상_전처리(aw, ax);
     var finalColor = color.startsWith("영림") ? color : co + color;
     var finalP = 품명_전처리_문짝(p, as + "*" + at + "*" + av);
     var sk = (aq && (aq.toString().includes("3방") || aq.toString().includes("식기무"))) ? "식기무" : (aq && aq.toString().includes("식기유") ? "식기유" : "");
     return finalColor + " " + finalP + " " + at + "*" + av + sk;
  }
  var co = (ap && ap.toString().includes("영림")) ? "영림" : "영림";
  var color = 색상_전처리(aw, ax);
  var spec = 추출_숫자_from문자열(as) + "*" + at + "*" + av + (aq && aq.toString().includes("3방") ? "식기무" : "식기유");
  return (color.startsWith("영림") ? color : co + color) + " " + 품명_전처리(ap) + " " + spec;
}

function 색상_전처리(aw, ax) {
  var s1 = aw ? aw.toString().trim() : "", s2 = ax ? ax.toString().trim() : "";
  var comb = (s1 && s2) ? s1 + " " + s2 : (s1 || s2 || "");
  if (!comb) return "";
  if (/^(영림\d+)\s+(PS\d+|[A-Z]+\d+)$/i.test(comb)) return comb.replace(/\s+/g, '');
  if (/^(영림\d+)\s+[가-힣]+$/.test(comb)) return comb.match(/영림\d+/)[0];
  if (/^[가-힣\s]+$/.test(comb)) return comb.replace(/\s+/g, '');
  return comb.replace(/영림|우딘|예림/g, '').trim();
}

function 품명_전처리(ap) {
  if (!ap) return "";
  return ap.toString().replace(/^영림ㅣ/, '').replace(/ㅣ/g, ' ').replace(/문틀|형/g, '').replace(/\d+바/g, '').replace(/\(식기[XO]\)/g, '').trim().replace(/\s+/g, ' ');
}

function 생성_품목코드_NEW(ap, aw, ax, as, at, av, aq, row) {
  var bc = 브랜드색상코드_생성(aw, ax), t = 구분_품목타입(ap, row), mid = "", spec = "";
  if (t === 'DOOR') { mid = 모델코드_생성(ap); spec = "" + at + av; }
  else { mid = 플래그코드_생성(ap); spec = 규격코드_생성(as, at, av, aq); }
  return bc + mid + spec;
}

function 브랜드색상코드_생성(aw, ax) {
  var s1 = aw ? aw.toString().trim() : "", s2 = ax ? ax.toString().trim() : "", c = (s1 && s2) ? s1 + s2 : (s1 || s2 || "");
  if (!c) throw new Error("색상 없음");
  var m1 = c.match(/영림(\d+)PS\d+/); if (m1) return "Y" + m1[1];
  var m2 = c.match(/PS([A-Z]+\d+)/i); if (m2) return "YS" + m2[1];
  var m3 = c.match(/영림(\d+)/); if (m3) return "Y" + m3[1];
  if (/^[가-힣]+$/.test(c)) return "Y" + c.substring(0, 2);
  var m5 = c.match(/(\d+)/); if (m5) return "Y" + m5[1];
  throw new Error("코드생성 실패");
}

function 플래그코드_생성(ap) {
  var s = ap.toString().replace(/^영림ㅣ/, '').split('ㅣ').map(function(k){ return k.replace(/형/g, '').trim(); });
  var yNum = null; s.forEach(function(k){ var m = k.match(/(\d+)연동/); if(m) yNum = m[1]; });
  var head = "", headMap = {"발포":"B","방염":"F","비방염":"N","알루미늄":"A"};
  for(var k in headMap) { if(s.indexOf(k)!==-1) { head = headMap[k]; break; } }
  if(yNum && ["F","N","A"].indexOf(head)!==-1) return head + yNum + "C";
  var tail = "", tailMap = {"슬림":"S","와이드":"W","분리":"D","히든":"H","미서기":"L"};
  for(var k in tailMap) { if(s.indexOf(k)!==-1) tail += tailMap[k]; }
  return head + tail;
}

function 규격코드_생성(as, at, av, aq) { return 추출_숫자_from문자열(as) + (at ? at.toString() : "") + (av ? av.toString() : "") + (aq && aq.toString().includes("3방") ? "N" : (aq && aq.toString().includes("4방") ? "Y" : "")); }

function 구분_품목타입(itemName, row) { 
  if (row >= 22 && row <= 34) return 'DOOR';
  var s = itemName ? itemName.toString() : "";
  if (['문틀','발포','분리형','스토퍼'].some(function(k){ return s.includes(k); })) return 'FRAME';
  if (['문짝','ABS','도어','M/D','민무늬','탈공','미서기','미닫이'].some(function(k){ return s.includes(k); }) || /YS-|YA-|YAT-|EZ-|LS-|YM-|YAL-|YV-|YFL-|SW-|TD-|SL-|TA-/i.test(s) || /\d+연동/.test(s)) return 'DOOR';
  return 'NONE';
}

function 모델코드_생성(itemName) {
  var s = itemName.toString().trim(), m = s.match(/([A-Z]+)-([A-Z0-9]+)/);
  if (m) { var suf = m[2], hIdx = suf.search(/[가-힣]/); return m[1] + (hIdx !== -1 ? suf.substring(0, hIdx) : suf); }
  if (s.includes('탈공')) return '탈';
  if (s.includes('M/D') && s.includes('민무늬')) return 'MD';
  var dM = s.match(/(\S+)도어/); return (dM && dM[1]) ? dM[1] : '';
}

function 품명_전처리_문짝(itemName, spec) {
  var s = itemName.toString().replace(/문틀/g, '').replace(/\(식기[XO]\)/g, '').trim();
  var fM = spec.toString().match(/^(\d+)/);
  if (fM) s = s.replace(new RegExp(fM[1] + '[가-힣]+', 'g'), '').trim();
  return s;
}

// ============================================
// 3. 전체 실행
// ============================================

function 전체_실행() {
  Logger.log("전체 실행 시작");
  계산_영림발주서_가격_내부();
  생성_품목코드_문틀_내부();
  SpreadsheetApp.getUi().alert("✅ 전체 실행 완료 (단가계산 + 코드생성)");
}

// ============================================
// 4. 버튼 관리
// ============================================

function 시트에_버튼_만들기() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  if (시트_버튼_찾기()) { if (SpreadsheetApp.getUi().alert('버튼 재생성', '기존 버튼을 삭제하고 새로 만들까요?', SpreadsheetApp.getUi().ButtonSet.YES_NO) === SpreadsheetApp.getUi().Button.YES) 시트_버튼_삭제(); else return; }
  
  var r = sheet.getRange("BA1:BC1");
  r.merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("📦 품목코드 생성 (BA2 클릭)").setBackground("#4285f4").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("BA2").insertCheckboxes();
  sheet.getRange("BB2:BC2").merge().setValue("← 체크하면 자동 실행").setFontColor("#666666").setFontSize(10);
  
  SpreadsheetApp.getUi().alert('✅ 버튼 생성 완료\n\nAW열 색상 입력 전 상단 메뉴 [🔄 데이터 업데이트]를 꼭 실행해주세요!');
}

function 시트_버튼_찾기() {
  try { var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME).getRange("BA1").getValue(); return (v && v.toString().includes("품목코드 생성")) ? true : false; } catch(e) { return false; }
}

function 시트_버튼_삭제() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  var r = s.getRange("BA1:BC1"); if(r.isPartOfMerge()) r.breakApart(); r.clear().setBackground(null);
  s.getRange("BA2").clear();
  var r2 = s.getRange("BB2:BC2"); if(r2.isPartOfMerge()) r2.breakApart(); r2.clear();
}

/**
 * 전용 업데이트 기능들
 */
function updateColorCodeMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var test = ss.getSheetByName("테스트");
  if (!test) return;
  var data = test.getRange("V1:Z300").getValues();
  var map = {}, count = 0;
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 5; c++) {
      var t = data[r][c] ? data[r][c].toString().trim() : "";
      if (!t) continue;
      var c1 = "", c2 = "";
      if (t.includes("(") && t.includes(")")) {
         var mN = t.match(/(영림\s*\d+)/), mP = t.match(/\(([^)]+)\)/);
         if (mN && mP) { c1 = mN[1].replace(/\s+/g, ''); c2 = mP[1].trim(); }
      } else if (t.includes(" ")) {
         var p = t.split(" "); if (p.length >= 2) { c1 = p[0].trim(); c2 = p.slice(1).join(" ").trim(); }
      }
      if (c1 && c2) { 
        map[c1] = c2; 
        map[c2] = c1; // 역방향 매핑 추가
        count++; 
      }
    }
  }
  PropertiesService.getScriptProperties().setProperty("COLOR_MAP", JSON.stringify(map));
  Logger.log("색상 맵 업데이트 완료: " + count + "개 항목 (양방향 포함)");
}

function updateGasketColorMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var test = ss.getSheetByName("테스트");
  if (!test) return;
  var data = test.getRange("M1:U300").getValues();
  var mapM = {}, mapP = {}, mapS = {};
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    [[0,1,2,'M',mapM],[3,4,5,'P',mapP],[6,7,8,'S',mapS]].forEach(function(g){
      if (r[g[0]]) {
        var k = r[g[0]].toString().trim(), st = (r[g[2]]||"").toString().trim();
        g[4][k] = { gasketColor: (r[g[1]]||"").toString().trim(), status: st, isDiscontinued: ["단종","단종예정"].includes(st), group: g[3] };
      }
    });
  }
  PropertiesService.getScriptProperties().setProperty("GASKET_COLOR_MAP", JSON.stringify({M:mapM, P:mapP, S:mapS}));
}

function onEdit(e) {
  try {
    var r = e.range, s = r.getSheet();
    var sheetName = s.getName();
    var rowS = r.getRow(), colS = r.getColumn(), val = r.getValue();
    
    // 디버그 로그 1: 트리거 발생 정보
    Logger.log("[onEdit] 발생 - 시트: " + sheetName + ", 행: " + rowS + ", 열: " + colS + ", 값: " + val);
    
    if (sheetName !== CONFIG.SHEET_NAME) {
      Logger.log("[onEdit] 종료 - 대상 시트 아님 (" + CONFIG.SHEET_NAME + " 필요)");
      return;
    }
    
    if (r.getA1Notation() === "BA2" && val === true) { 
      생성_품목코드_문틀(); r.setValue(false); return; 
    }
    
    if (colS === 1 || colS === CONFIG.COLS.AW || colS === CONFIG.COLS.AT) { 
      for(var i=rowS; i<=r.getLastRow(); i++) if(i>=12 && i<=35) s.getRange(i,1).clearNote(); 
    }
    
    // AT(병합셀) 또는 AW(색상) 수정 시 실행
    if (colS === CONFIG.COLS.AW || colS === CONFIG.COLS.AT) {
       Logger.log("[onEdit] 처리 대상 열(AT/AW) 수정됨");
       var props = PropertiesService.getScriptProperties();
       var cMap = JSON.parse(props.getProperty("COLOR_MAP")||"{}"), gMap = JSON.parse(props.getProperty("GASKET_COLOR_MAP")||"{}");
       
       for (var i = rowS; i <= r.getLastRow(); i++) {
          if (i < 12 || i > 35) continue;
          
          var k = s.getRange(i, CONFIG.COLS.AW).getValue().toString().trim(), axV = "";
          var atValRaw = s.getRange(i, CONFIG.COLS.AT).getValue();
          var _pEdit = parseAT(atValRaw);
          var avV = _pEdit.height;
          
          Logger.log("[onEdit] 행 " + i + " 처리중 - 색상(AW): '" + k + "', 높이: " + avV + " (원본AT: " + atValRaw + ")");
          
          // 1. AX열 (색상코드) 자동완성 (AW 수정 시)
          if (colS === CONFIG.COLS.AW && k) {
             var res = cMap[k] || cMap[k.replace(/\s+/g,'')];
             if (!res) for (var mK in cMap) if (mK.includes(k) || k.includes(mK)) { res = cMap[mK]; break; }
             if (res) { s.getRange(i, CONFIG.COLS.AX).setValue(res); axV = res; }
          } else {
             axV = s.getRange(i, CONFIG.COLS.AX).getValue().toString().trim();
          }

          // 2. BA열 가스켓 (12-20행)
          if (i <= 20 && k) {
             var targets = [k]; if(axV) targets.push(axV);
             var found = null;
             ['M','P','S'].forEach(function(g){
                if(found) return;
                for(var dK in gMap[g]) {
                   targets.forEach(function(t){ if(found) return; if(dK===t||dK.includes(t)||t.includes(dK)) found=gMap[g][dK]; });
                }
             });
             if(found) s.getRange(i, CONFIG.COLS.BA).setValue(found.isDiscontinued ? found.status : found.gasketColor);
          }
          
          // 3. AQ열 인필 방향 (22-35행)
          if (i >= 22) {
             if (avV >= 2166 && k) {
                var test = e.source.getSheetByName("테스트");
                if (test) {
                   var ag = test.getRange("AG1:AG300").getValues(), ah = test.getRange("AH1:AH300").getValues();
                   var sk = [k.toUpperCase()]; if(axV) sk.push(axV.toUpperCase());
                   
                   var isY = ah.some(function(row){ 
                     var v = (row[0]||"").toString().toUpperCase().trim(); 
                     return v && sk.some(function(key){ return v.includes(key) || key.includes(v); }); 
                   });
                   var isN = !isY && ag.some(function(row){ 
                     var v = (row[0]||"").toString().toUpperCase().trim(); 
                     return v && sk.some(function(key){ return v.includes(key) || key.includes(v); }); 
                   });
                   
                   Logger.log("[onEdit] 행 " + i + " 매칭 결과 - isY: " + isY + ", isN: " + isN);
                   if (isY) s.getRange(i, CONFIG.COLS.AQ).setValue("Y"); 
                   else if (isN) s.getRange(i, CONFIG.COLS.AQ).setValue("N");
                   else {
                      Logger.log("[onEdit] 행 " + i + " 매칭 실패 (Y/N 대상 아님)");
                      s.getRange(i, CONFIG.COLS.AQ).setValue(""); 
                   }
                } else {
                   Logger.log("[onEdit] '테스트' 시트 못찾음");
                }
             } else {
                // 높이가 2166 미만이거나 색상이 없는 경우 -> 'N' 표시
                Logger.log("[onEdit] 행 " + i + " 높이(" + avV + ") < 2166 또는 색상 없음 -> N 설정");
                s.getRange(i, CONFIG.COLS.AQ).setValue("N");
             }
          }
       }
    }
  } catch(err) { Logger.log("onEdit Error: " + err.message); }
}

function 초기화_영림발주서() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (SpreadsheetApp.getUi().alert('⚠️ 초기화', '데이터를 모두 지우겠습니까?', SpreadsheetApp.getUi().ButtonSet.YES_NO) === SpreadsheetApp.getUi().Button.YES) {
    s.getRange("A12:A20").clearContent().clearNote(); s.getRange("A22:A35").clearContent().clearNote();
    s.getRange("AR12:BD20").clearContent(); s.getRange("AR22:BD35").clearContent();
    s.getRange("AQ22:AQ35").clearContent();
  }
}

/**
 * AP12:20 영역의 데이터 유효성 검사(드롭다운)를 복구합니다.
 * '필터' 시트 또는 '영림문틀단가표' 시트의 데이터를 사용합니다.
 */
function setDropdowns_AP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var filterSheet = ss.getSheetByName('필터');
  var priceSheet = ss.getSheetByName('영림문틀단가표');
  
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert('오류: "영림발주서" 시트를 찾을 수 없습니다.');
    return;
  }
  
  var range = sheet.getRange("AP12:AP20");
  
  // 후보 1: 필터 시트의 AP열
  var sourceRange = null;
  if (filterSheet) {
    var filterValues = filterSheet.getRange("AP:AP").getValues().filter(function(r) { return r[0] !== ""; });
    if (filterValues.length > 0) {
      sourceRange = filterSheet.getRange("AP:AP");
    }
  }
  
  // 후보 2: 영림문틀단가표의 C열
  if (!sourceRange && priceSheet) {
    var priceValues = priceSheet.getRange("C6:C500").getValues().filter(function(r) { return r[0] !== ""; });
    if (priceValues.length > 0) {
      sourceRange = priceSheet.getRange("C6:C500");
    }
  }

  if (!sourceRange) {
    ui.alert('❌ 드롭다운 목록으로 사용할 데이터를 찾지 못했습니다.\n"필터" 시트의 AP열이나 "영림문틀단가표" 시트의 C열을 확인해 주세요.');
    return;
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange)
    .setAllowInvalid(true) 
    .build();
    
  range.setDataValidation(rule);
  
  ui.alert('✅ AP12:20 드롭다운 설정 완료!\n사용된 소스: ' + sourceRange.getSheet().getName() + ' 시트');
}

/**
 * 드롭다운 소스 후보를 분석하여 팝업으로 보여줍니다.
 */
function findCorrectAPSource() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var filterSheet = ss.getSheetByName('필터');
  var priceSheet = ss.getSheetByName('영림문틀단가표');
  var report = "--- 드롭다운 후보 분석 ---\n\n";

  if (filterSheet) {
    report += "[필터 시트]\n";
    var cols = ["AP", "C", "D", "E"];
    cols.forEach(function(col) {
      try {
        var vals = filterSheet.getRange(col + "1:" + col + "100").getValues().filter(function(r) { return r[0] !== ""; });
        report += "- " + col + "열: " + vals.length + "개 항목 (샘플: " + (vals.length > 0 ? vals[0][0] : "없음") + ")\n";
      } catch(e) {}
    });
  }

  if (priceSheet) {
    report += "\n[영림문틀단가표]\n";
    try {
      var vals = priceSheet.getRange("C6:C100").getValues().filter(function(r) { return r[0] !== ""; });
      report += "- C열 (C6:C100): " + vals.length + "개 항목 (샘플: " + (vals.length > 0 ? vals[0][0] : "없음") + ")\n";
    } catch(e) {}
  }

  SpreadsheetApp.getUi().alert(report + "\n이 중 AP12:20에 나타나야 할 목록이 포함된 열을 알려주세요.");
}

/**
 * AP12 셀의 현재 유효성 검사 상태를 로그로 출력합니다.
 */
function inspectValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var range = sheet.getRange("AP12");
  var rule = range.getDataValidation();
  
  var log = "--- AP12 유효성 검사 상태 ---\n";
  if (!rule) {
    log += "규칙 없음\n";
  } else {
    log += "유형: " + rule.getCriteriaType() + "\n";
    log += "잘못된 데이터 허용: " + rule.getAllowInvalidData() + "\n";
    log += "도움말 텍스트: " + rule.getHelpText() + "\n";
    var values = rule.getCriteriaValues();
    log += "기준 값: " + (values ? values.join(", ") : "없음") + "\n";
  }
  
  SpreadsheetApp.getUi().alert(log);
  Logger.log(log);
}

/**
 * AP12:20 영역의 모든 유효성 검사를 완전히 제거합니다.
 */
function clearValidation_AP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var range = sheet.getRange("AP12:AP20");
  range.clearDataValidations();
  SpreadsheetApp.getUi().alert('🗑️ AP12:20 유효성 검사가 모두 삭제되었습니다. 이제 자유롭게 입력이 가능한지 확인해 보세요.');
}

/**
 * [진단] 12행(A12) 가격 계산 프로세스 상세 추적
 */
function debugRow12() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var priceSheet = ss.getSheetByName("영림문틀단가표");
  
  if (!sheet || !priceSheet) return SpreadsheetApp.getUi().alert("필요한 시트를 찾을 수 없습니다.");
  
  var log = "=== [A12 계산 진단 - 정밀 버전] ===\n\n";
  var row = 12;
  
  // 1. 입력값 읽기
  var ap = sheet.getRange("AP" + row).getValue();
  var aq = sheet.getRange("AQ" + row).getValue();
  var ar = sheet.getRange("AR" + row).getValue();
  var as = sheet.getRange("AS" + row).getValue();
  var atRaw = sheet.getRange("AT" + row).getValue();
  var _pDbg = parseAT(atRaw);
  var at = _pDbg.width, av = _pDbg.height;
  
  log += "[1. 입력 시트 데이터]\n";
  var apStr = (ap || "").toString();
  log += "- 제품명(AP): " + (apStr || "(비어있음)") + "\n";
  
  var codes = [];
  for (var c=0; c<apStr.length; c++) codes.push(apStr.charCodeAt(c).toString(16).toUpperCase());
  log += "  (제품명 HEX 코드: " + codes.join(" ") + ")\n";
  
  log += "- 인필(AQ): " + (aq || "(비어있음)") + "\n";
  log += "- 수량(AR): " + ar + "\n";
  log += "- 규격(AS): " + (as || "(비어있음)") + (typeof as === 'number' ? " (Number)" : " (String)") + "\n";
  log += "- 규격(AT): " + atRaw + " → 너비: " + at + ", 높이: " + av + "\n\n";
  
  if (!ap || !atRaw) {
    log += "❌ 오류: 필수 입력값 부족\n";
    return SpreadsheetApp.getUi().alert(log);
  }

  // 2. 단가표 매칭 시도
  log += "[2. 단가표 매칭 분석]\n";
  var type = 추출_제품타입(ap);
  log += "- 추출된 제품타입: [" + type + "]\n";
  
  var searchSize = as ? as.toString().trim() : "";
  var searchDir = aq ? aq.toString().trim() : "";
  log += "- 검색 조건: 타입=[" + type + "], 규격=[" + searchSize + "], 인필=[" + searchDir + "]\n\n";
  
  var priceData = priceSheet.getRange("C6:F500").getValues();
  
  // 단가표 샘플 5개 출력
  log += "[3. 단가표 상단 5개 데이터 샘플]\n";
  for (var i=0; i<Math.min(5, priceData.length); i++) {
    log += "  " + (i+6) + "행: " + priceData[i].map(function(v){ return "["+v+"]"; }).join(", ") + "\n";
  }
  log += "...\n\n";

  var supplyPrice = 찾기_공급가(priceData, type, as, aq);
  
  if (supplyPrice === null) {
    log += "❌ 실패: 단가표에서 일치하는 항목 없음\n";
    
    // 왜 실패했는지 추측을 위한 정밀 검사 (첫 100행 중 타입이라도 맞는게 있는지)
    var partialMatch = false;
    for (var j=0; j<Math.min(100, priceData.length); j++) {
      var rowC = priceData[j][0] ? priceData[j][0].toString().trim() : "";
      if (rowC.includes(type)) {
        log += "  🔍 힌트: " + (j+6) + "행에서 타입 '" + type + "'은 발견됨. \n";
        log += "     하지만 규격[" + priceData[j][1] + "] 또는 인필[" + priceData[j][2] + "]이 일치하지 않음.\n";
        partialMatch = true;
        break;
      }
    }
    if (!partialMatch) log += "  🔍 힌트: 단가표 C열에서 '" + type + "' 성분을 포함한 행을 찾지 못했습니다.\n";

  } else {
    log += "✅ 성공: 기본 단가 -> " + supplyPrice + "\n";
  }
  
  SpreadsheetApp.getUi().alert(log);
}
