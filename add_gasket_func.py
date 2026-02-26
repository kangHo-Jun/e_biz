# updateGasketColorMap 함수 추가
code = '''

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
'''

with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'a', encoding='utf-8') as f:
    f.write(code)

print("✅ updateGasketColorMap 함수 추가 완료!")
