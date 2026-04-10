# 디버그 함수 추가
debug_code = '''

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
  var log = "=== 검색 디버그: " + searchKey + " ===\\n\\n";
  
  // 각 그룹별 데이터 확인
  var groupNames = ['M', 'P', 'S'];
  
  for (var gi = 0; gi < groupNames.length; gi++) {
    var groupName = groupNames[gi];
    var groupMap = gasketMap[groupName];
    
    log += "📁 " + groupName + " 그룹:\\n";
    
    if (!groupMap) {
      log += "  (데이터 없음)\\n";
      continue;
    }
    
    // searchKey가 포함된 항목 찾기
    var found = false;
    for (var key in groupMap) {
      if (key.includes(searchKey) || searchKey.includes(key)) {
        var info = groupMap[key];
        log += "  ✅ 매칭! 키: " + key + "\\n";
        log += "     gasketColor: " + info.gasketColor + "\\n";
        log += "     status: " + info.status + "\\n";
        log += "     isDiscontinued: " + info.isDiscontinued + "\\n";
        found = true;
      }
    }
    
    if (!found) {
      log += "  (매칭 항목 없음)\\n";
    }
    
    log += "\\n";
  }
  
  Logger.log(log);
  SpreadsheetApp.getUi().alert(log);
}
'''

with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'a', encoding='utf-8') as f:
    f.write(debug_code)

print("✅ 디버그 함수 추가 완료!")
