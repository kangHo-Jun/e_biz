# 직접 테스트 함수 추가
test_code = '''

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
  
  var msg = "=== 검색 결과: " + searchKey + " ===\\n\\n";
  
  if (foundInfo) {
    msg += "✅ 매칭 성공!\\n";
    msg += "그룹: " + foundGroup + "\\n";
    msg += "gasketColor: " + foundInfo.gasketColor + "\\n";
    msg += "status: " + foundInfo.status + "\\n";
    msg += "isDiscontinued: " + foundInfo.isDiscontinued + "\\n\\n";
    
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
'''

with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'a', encoding='utf-8') as f:
    f.write(test_code)

print("✅ 테스트 함수 추가 완료!")
