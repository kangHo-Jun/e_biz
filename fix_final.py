# UTF-8로 파일 읽기
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    content = f.read()

# 1694~1750행 부분을 완전히 새로 작성
# 먼저 해당 부분을 찾아서 제거하고 새로 삽입

# 파일을 라인 단위로 읽기
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 1694~1800행 사이의 중복/오류 부분을 찾아서 정리
# 간단하게: 1694행부터 "// 3. AQ열 자동 입력" 이 나올 때까지 모두 교체

start_idx = 1693  # "// 2. BA열 품목 정보 자동 입력" 줄
end_marker = "           // 3. AQ열 자동 입력"

# end_idx 찾기
end_idx = -1
for i in range(start_idx, min(len(lines), start_idx + 200)):
    if end_marker in lines[i]:
        end_idx = i
        break

if end_idx == -1:
    print("❌ 종료 마커를 찾을 수 없습니다!")
    exit(1)

print(f"교체 범위: {start_idx+1}행 ~ {end_idx}행")

# 새로운 코드 블록
new_code = '''           // 2. BA열 품목 정보 자동 입력 (신규 로직)
           // 12행~20행만 적용 (22~34행 제외, 헤더 21행 제외)
           if (row >= 12 && row <= 20) {
               var jsonInfoMap = props.getProperty("PRODUCT_INFO_MAP");
               if (jsonInfoMap) {
                  var infoMap = JSON.parse(jsonInfoMap);
                  var targetKeywords = [inputKey];
                  if (axValue) targetKeywords.push(axValue);
                  
                  var foundInfo = null;
                  var foundGroup = "";
                  
                  // [수정] 그룹별 맵 구조에 맞게 검색
                  // infoMap = {M: {...}, P: {...}, S: {...}}
                  var groupNames = ['M', 'P', 'S'];
                  
                  var found = false;
                  for (var gi = 0; gi < groupNames.length && !found; gi++) {
                    var groupName = groupNames[gi];
                    var groupMap = infoMap[groupName];
                    
                    if (!groupMap) continue;
                    
                    for (var dbKey in groupMap) {
                      for (var k = 0; k < targetKeywords.length; k++) {
                        var uKey = targetKeywords[k];
                        if (dbKey.includes(uKey) || uKey.includes(dbKey)) {
                          foundInfo = groupMap[dbKey];
                          foundGroup = groupName;
                          found = true;
                          break;
                        }
                      }
                      if (found) break;
                    }
                  }
                  
                   if (foundInfo) {
                      var baCell = sheet.getRange(row, 53); // BA열
                      
                      // 그룹 정보 로그
                      Logger.log("  [BA 자동입력] 행 " + row + ": 매칭됨 (그룹: " + foundGroup + ") → 상태: " + foundInfo.statusText);
                      
                      // 상태 텍스트가 있으면 우선 사용
                      if (foundInfo.statusText) {
                         baCell.setValue(foundInfo.statusText);
                         e.source.toast("품목 정보: " + foundInfo.statusText + " (그룹: " + foundGroup + ")", "정보");
                      } else if (foundInfo.isDiscontinued) {
                         baCell.setValue("단종");
                         e.source.toast("품목 정보: 단종 제품입니다. (그룹: " + foundGroup + ")", "정보");
                      } else {
                         baCell.setValue(foundInfo.value);
                         e.source.toast("품목 정보: " + foundInfo.value + " (그룹: " + foundGroup + ")", "성공");
                      }
                   } else {
                      // 정보가 없으면 BA열 비우기? (사용자 요청에 따라 결정)
                   }
               }
           }
           
'''

# 교체
lines[start_idx:end_idx] = [new_code]

# 파일 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✅ 수정 완료!")
print(f"총 라인 수: {len(lines)}")
