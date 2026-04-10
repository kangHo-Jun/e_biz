
# onEdit 함수 추가
onedit_code = '''
/**
 * onEdit 트리거 함수
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    var val = range.getValue();
    
    if (sheet.getName() === "영림발주서" && col === 49 && row >= 12 && row <= 20) {
      if (!val || val.toString().trim() === "") return;
      
      var inputKey = val.toString().trim();
      var axValue = sheet.getRange(row, 50).getValue();
      var targetKeywords = [inputKey];
      if (axValue) targetKeywords.push(axValue.toString().trim());
      
      var props = PropertiesService.getScriptProperties();
      var jsonMap = props.getProperty("GASKET_COLOR_MAP");
      
      if (!jsonMap) {
        e.source.toast("데이터 없음. updateGasketColorMap() 실행 필요", "오류");
        return;
      }
      
      var gasketMap = JSON.parse(jsonMap);
      var foundInfo = null;
      var foundGroup = "";
      var groupNames = ['M', 'P', 'S'];
      var found = false;
      
      for (var gi = 0; gi < groupNames.length && !found; gi++) {
        var groupName = groupNames[gi];
        var groupMap = gasketMap[groupName];
        if (!groupMap) continue;
        for (var key in groupMap) {
          for (var k = 0; k < targetKeywords.length; k++) {
            var searchKey = targetKeywords[k];
            if (key.includes(searchKey) || searchKey.includes(key)) {
              foundInfo = groupMap[key];
              foundGroup = groupName;
              found = true;
              break;
            }
          }
          if (found) break;
        }
      }
      
      if (foundInfo) {
        var baCell = sheet.getRange(row, 53);
        var outputValue = foundInfo.isDiscontinued ? foundInfo.status : foundInfo.gasketColor;
        baCell.setValue(outputValue);
        var msg = "BA" + row + ": " + outputValue + " (그룹: " + foundGroup + ")";
        if (!foundInfo.isDiscontinued) msg += " [5,500원 추가 대상]";
        e.source.toast(msg, "자동입력");
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
