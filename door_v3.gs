/**
 * 영림발주서 통합 스크립트 v3 - 데이터 업데이트 강화 버전
 * 
 * 기능:
 * 1. 단가계산: A12~A35에 가격 출력
 * 2. 코드생성: BC12~BF35에 품목명/코드 출력
 * 3. 전체실행: 1+2 순차 실행
 * 4. 데이터 업데이트: 색상코드 및 가스켓 정보를 한 번에 동기화
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
  ui.createMenu('🔧 영림발주서 v3')
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
             '1. 색상 매핑 데이터 업데이트 완료\n' + 
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
    AP: 42, AQ: 43, AR: 44, AS: 45, AT: 46, AV: 48, AW: 49, AX: 50, BA: 53
  }
};

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
    var ap = rowData[0], aq = rowData[1], ar = rowData[2], as = rowData[3], at = rowData[4], av = rowData[6], aw = rowData[7], ax = rowData[8], ba = rowData[11];
    var curP = aValues[i][0], curN = aNotes[i][0];

    if (!at && !av) { resultPrices.push([curP]); resultNotes.push([curN]); 실패++; continue; }
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
            if (doorP > 0 && (Number(av)||0) >= 2166) finalP += doorP;
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
    var sheet = ss.getSheetByName(CONFIG.TEST_SHEET_NAME);
    if (!sheet) return { additionalPriceInfo: [], additionalPriceMap: {}, doorPriceMap: {} };
    var raw = sheet.getRange("V3:Z300").getValues();
    var addMap = {};
    for (var r = 1; r < raw.length; r++) {
       for (var c = 0; c < 5; c++) {
           if (raw[r][c]) addMap[raw[r][c].toString().toUpperCase().trim()] = c;
       }
    }
    var doorData = sheet.getRange("AD1:AF50").getValues();
    var doorMap = {};
    for (var i = 0; i < doorData.length; i++) {
        if (doorData[i][0] && doorData[i][2]) {
            doorData[i][0].toString().split(',').forEach(function(k){ doorMap[k.trim().toUpperCase()] = Number(doorData[i][2]) || 0; });
        }
    }
    return { additionalPriceInfo: raw[0], additionalPriceMap: addMap, doorPriceMap: doorMap };
}

function findPriceFromMap_Scan(target, map, infoArr) {
    if (map[target] !== undefined) return infoArr[map[target]];
    for (var key in map) { if (key.includes(target) || target.includes(key)) return infoArr[map[key]]; }
    return null;
}

function findDoorPrice_Scan(target, map) {
    for (var key in map) { if (target.includes(key)) return map[key]; }
    return 0;
}

function 추출_제품타입(ap값) {
  var s = ap값 ? ap값.toString() : "";
  var m = s.match(/(\d+)방/);
  if (m) {
    var res = s.substring(0, s.indexOf(m[0]));
    return res.endsWith("ㅣ") ? res.slice(0,-1) : res;
  }
  return s;
}

function 정규화_키워드(k) {
  var s = k ? k.toString().trim() : "";
  return s.endsWith("형") ? s.slice(0, -1) : s;
}

function 찾기_공급가(data, type, size, direction) {
  var t = type ? type.toString().trim() : "", s = size ? size.toString().trim() : "", d = direction ? direction.toString().trim() : "";
  if (!t || !s || !d) return null;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var c = row[0] ? row[0].toString().trim().replace(/형/g, '') : "";
    if (t.split('ㅣ').every(function(kw){ return c.includes(정규화_키워드(kw)); }) && row[1].toString().trim() === s && row[2].toString().trim() === d) {
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
    if (!row[4] && !row[6]) { names.push([""]); codes.push([""]); empty.push([""]); units.push([""]); 실패++; continue; }
    if (Math.max(추출_숫자_from문자열(row[3]), Number(row[4])||0, Number(row[6])||0) <= 999) { 실패++; continue; }

    try {
      var n = 생성_품목명(row[0], row[7], row[8], row[3], row[4], row[6], row[1], rIdx);
      var c = 생성_품목코드_NEW(row[0], row[7], row[8], row[3], row[4], row[6], row[1], rIdx);
      names.push([n]); codes.push([c]); empty.push([""]); units.push([rIdx >= 22 ? "짝" : "틀"]); 성공++;
    } catch (e) {
       실패++;
    }
  }
  sheet.getRange(start, 55, num, 1).setValues(names);
  sheet.getRange(start, 56, num, 1).setValues(codes);
  sheet.getRange(start, 57, num, 1).setValues(empty);
  sheet.getRange(start, 58, num, 1).setValues(units);
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
     return finalColor + " " + finalP + " " + as + "*" + at + "*" + av + sk;
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
  if (t === 'DOOR') { mid = 모델코드_생성(ap); spec = "" + 추출_숫자_from문자열(as) + at + av; }
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
  
  var r = sheet.getRange("BC1:BE1");
  r.merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("📦 품목코드 생성 (BC2 클릭)").setBackground("#4285f4").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("BC2").insertCheckboxes();
  sheet.getRange("BD2:BE2").merge().setValue("← 체크하면 자동 실행").setFontColor("#666666").setFontSize(10);
  
  SpreadsheetApp.getUi().alert('✅ 버튼 생성 완료\n\nAW열 색상 입력 전 상단 메뉴 [🔄 데이터 업데이트]를 꼭 실행해주세요!');
}

function 시트_버튼_찾기() {
  try { var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME).getRange("BC1").getValue(); return (v && v.toString().includes("품목코드 생성")) ? true : false; } catch(e) { return false; }
}

function 시트_버튼_삭제() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  var r = s.getRange("BC1:BE1"); if(r.isPartOfMerge()) r.breakApart(); r.clear().setBackground(null);
  s.getRange("BC2").clear();
  var r2 = s.getRange("BD2:BE2"); if(r2.isPartOfMerge()) r2.breakApart(); r2.clear();
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
    if (s.getName() !== CONFIG.SHEET_NAME) return;
    var rowS = r.getRow(), colS = r.getColumn(), val = r.getValue();
    if (r.getA1Notation() === "BC2" && val === true) { 생성_품목코드_문틀(); r.setValue(false); return; }
    if (colS === 1 || colS === CONFIG.COLS.AW) { for(var i=rowS; i<=r.getLastRow(); i++) if(i>=12 && i<=35) s.getRange(i,1).clearNote(); }
    if (colS === CONFIG.COLS.AW) {
       var props = PropertiesService.getScriptProperties();
       var cMap = JSON.parse(props.getProperty("COLOR_MAP")||"{}"), gMap = JSON.parse(props.getProperty("GASKET_COLOR_MAP")||"{}");
       for (var i = rowS; i <= r.getLastRow(); i++) {
          if (i < 12 || i > 35) continue;
          var k = s.getRange(i, CONFIG.COLS.AW).getValue().toString().trim(), axV = "";
          if (!k) continue;
          var res = cMap[k] || cMap[k.replace(/\s+/g,'')];
          if (!res) for (var mK in cMap) if (mK.includes(k) || k.includes(mK)) { res = cMap[mK]; break; }
          if (res) { s.getRange(i, CONFIG.COLS.AX).setValue(res); axV = res; }
          if (i <= 20) {
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
          if (i >= 22) {
             var avV = Number(s.getRange(i, CONFIG.COLS.AV).getValue()) || 0;
             if (avV >= 2166) {
                var test = e.source.getSheetByName("테스트");
                if (test) {
                   var ag = test.getRange("AG1:AG300").getValues(), ah = test.getRange("AH1:AH300").getValues(), sk = [k.toUpperCase()]; if(axV) sk.push(axV.toUpperCase());
                   var isY = ah.some(function(row){ var v = (row[0]||"").toString().toUpperCase(); return v && sk.some(function(key){ return v.includes(key); }); });
                   var isN = !isY && ag.some(function(row){ var v = (row[0]||"").toString().toUpperCase(); return v && sk.some(function(key){ return v.includes(key); }); });
                   if (isY) s.getRange(i, CONFIG.COLS.AQ).setValue("Y"); else if (isN) s.getRange(i, CONFIG.COLS.AQ).setValue("N");
                }
             } else s.getRange(i, CONFIG.COLS.AQ).setValue("");
          }
       }
    }
  } catch(err) { Logger.log(err.message); }
}

function 초기화_영림발주서() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (SpreadsheetApp.getUi().alert('⚠️ 초기화', '데이터를 모두 지우겠습니까?', SpreadsheetApp.getUi().ButtonSet.YES_NO) === SpreadsheetApp.getUi().Button.YES) {
    s.getRange("A12:A20").clearContent().clearNote(); s.getRange("A22:A35").clearContent().clearNote();
    s.getRange("AR12:BF20").clearContent(); s.getRange("AR22:BF35").clearContent();
    s.getRange("AQ22:AQ35").clearContent();
  }
}
