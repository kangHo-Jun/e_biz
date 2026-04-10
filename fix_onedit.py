import re

# 파일 읽기
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    content = f.read()

# 수정할 부분 찾기 (1703~1714행)
old_code = r'''                  var foundInfo = null;
                  
                  for \(var dbKey in infoMap\) \{
                    for \(var k = 0; k < targetKeywords\.length; k\+\+\) \{
                      var uKey = targetKeywords\[k\];
                      if \(dbKey\.includes\(uKey\) \|\| uKey\.includes\(dbKey\)\) \{
                        foundInfo = infoMap\[dbKey\];
                        break;
                      \}
                    \}
                    if \(foundInfo\) break;
                  \}'''

new_code = '''                  var foundInfo = null;
                  var foundGroup = "";
                  
                  // [수정] 그룹별 맵 구조에 맞게 검색
                  // infoMap = {M: {...}, P: {...}, S: {...}}
                  var groupNames = ['M', 'P', 'S'];
                  
                  outerSearch:
                  for (var gi = 0; gi < groupNames.length; gi++) {
                    var groupName = groupNames[gi];
                    var groupMap = infoMap[groupName];
                    
                    if (!groupMap) continue;
                    
                    for (var dbKey in groupMap) {
                      for (var k = 0; k < targetKeywords.length; k++) {
                        var uKey = targetKeywords[k];
                        if (dbKey.includes(uKey) || uKey.includes(dbKey)) {
                          foundInfo = groupMap[dbKey];
                          foundGroup = groupName;
                          break outerSearch;
                        }
                      }
                    }
                  }'''

# 정규식으로 치환
content = re.sub(old_code, new_code, content, flags=re.MULTILINE)

# 토스트 메시지에 groupInfo 추가
content = content.replace(
    'e.source.toast("품목 정보: " + foundInfo.statusText, "정보");',
    'e.source.toast("품목 정보: " + foundInfo.statusText + " (그룹: " + foundGroup + ")", "정보");'
)
content = content.replace(
    'e.source.toast("품목 정보: 단종 제품입니다.", "정보");',
    'e.source.toast("품목 정보: 단종 제품입니다. (그룹: " + foundGroup + ")", "정보");'
)
content = content.replace(
    'e.source.toast("품목 정보: " + foundInfo.value, "성공");',
    'e.source.toast("품목 정보: " + foundInfo.value + " (그룹: " + foundGroup + ")", "성공");'
)

# 로그 추가
content = content.replace(
    '                      // [수정] 상태 텍스트가 있으면 우선 사용',
    '''                      // 그룹 정보 로그
                      Logger.log("  [BA 자동입력] 행 " + row + ": 매칭됨 (그룹: " + foundGroup + ") → 상태: " + foundInfo.statusText);
                      
                      // 상태 텍스트가 있으면 우선 사용'''
)

# 파일 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.write(content)

print("✅ 수정 완료!")
