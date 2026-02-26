# onEdit 함수 추가
onedit_code = '''

/**
 * onEdit 트리거 - AW열 입력 시 BA열 자동 입력
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    var val = range.getValue();
    
    // 영림발주서 시트의 AW열(49) 입력 시
    if (sheet.getName() === "영림발주서" && col === 49 && row >= 12 && row <= 20) {
      if (!val || val.toString().trim() === "") return;
      
      var inputKey = val.toString().trim();
      var props = PropertiesService.getScriptProperties();
      var jsonMap = props.getProperty("GASKET_COLOR_MAP");
      
      if (!jsonMap) {
        e.source.toast("데이터 없음. updateGasketColorMap() 실행 필요", "오류");
        return;
      }
      
      var infoMap = JSON.parse(jsonMap);
      var targetKeywords = [inputKey];
      
      var foundInfo = null;
      var foundGroup = "";
      var bestMatch = null;
      var bestMatchKey = "";
      var bestMatchGroup = "";
      var groupNames = ['M', 'P', 'S'];
      
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
  } catch (error) {
    Logger.log("❌ onEdit 오류: " + error.message);
  }
}
'''

with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'a', encoding='utf-8') as f:
    f.write(onedit_code)

print("✅ onEdit 함수 추가 완료!")
