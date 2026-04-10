# door_v1.gs 파일 수정 스크립트
# 1. updateColorCodeMap() 함수 수정 (그룹별 맵 분리)
# 2. onEdit() 함수 수정 (그룹별 검색)

with open(r'c:\Users\DSAI\Desktop\이비즈\door_v1.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"원본 파일 라인 수: {len(lines)}")

# ============================================
# 1. updateColorCodeMap() 함수 수정
# ============================================

# "var rangeData = testSheet.getRange" 찾기
start_idx = -1
for i, line in enumerate(lines):
    if 'var rangeData = testSheet.getRange("M1:U300")' in line:
        start_idx = i
        break

if start_idx == -1:
    print("❌ updateColorCodeMap 시작점을 찾을 수 없습니다!")
    exit(1)

print(f"✅ updateColorCodeMap 시작: {start_idx + 1}행")

# 그룹 정의 부분 찾기 (var groups = [)
groups_idx = -1
for i in range(start_idx, min(len(lines), start_idx + 20)):
    if 'var groups = [' in lines[i]:
        groups_idx = i
        break

if groups_idx == -1:
    print("❌ groups 정의를 찾을 수 없습니다!")
    exit(1)

# groups 정의 교체 (3줄)
lines[groups_idx] = "      // [수정] 각 그룹별로 별도 맵 생성 (중복 제품명 문제 해결)\r\n"
lines.insert(groups_idx + 1, "      var infoMapM = {}; // M열 제품 (O열 상태)\r\n")
lines.insert(groups_idx + 2, "      var infoMapP = {}; // P열 제품 (R열 상태)\r\n")
lines.insert(groups_idx + 3, "      var infoMapS = {}; // S열 제품 (U열 상태)\r\n")
lines.insert(groups_idx + 4, "      \r\n")
lines.insert(groups_idx + 5, "      // 3개 그룹을 각각 순회\r\n")
lines.insert(groups_idx + 6, "      var groups = [\r\n")
lines.insert(groups_idx + 7, "        {keyIdx: 0, valIdx: 1, stsIdx: 2, group: 'M', map: infoMapM}, // M, N, O\r\n")
lines.insert(groups_idx + 8, "        {keyIdx: 3, valIdx: 4, stsIdx: 5, group: 'P', map: infoMapP}, // P, Q, R\r\n")
lines.insert(groups_idx + 9, "        {keyIdx: 6, valIdx: 7, stsIdx: 8, group: 'S', map: infoMapS}  // S, T, U\r\n")
lines.insert(groups_idx + 10, "      ];\r\n")

# 기존 groups 정의 3줄 삭제
del lines[groups_idx + 11:groups_idx + 14]

# infoMap[keyStr] = infoObj; 찾아서 grp.map[keyStr] = infoObj;로 교체
for i in range(groups_idx, min(len(lines), groups_idx + 100)):
    if 'infoMap[keyStr] = infoObj;' in lines[i]:
        lines[i] = lines[i].replace('infoMap[keyStr] = infoObj;', 'grp.map[keyStr] = infoObj;')
        lines[i - 1] = "             // [수정] 각 그룹의 맵에 별도로 저장 (덮어쓰기 방지)\r\n"
        print(f"✅ infoMap 저장 방식 수정: {i + 1}행")
        break

# } else { 다음에 infoMap 통합 코드 추가
else_idx = -1
for i in range(groups_idx, min(len(lines), groups_idx + 100)):
    if '} else {' in lines[i] and '테스트' in lines[i + 1]:
        else_idx = i
        break

if else_idx != -1:
    lines.insert(else_idx, "      \r\n")
    lines.insert(else_idx + 1, "      // [신규] 3개 맵을 하나의 객체로 통합하여 저장\r\n")
    lines.insert(else_idx + 2, "      infoMap = {\r\n")
    lines.insert(else_idx + 3, "        M: infoMapM,\r\n")
    lines.insert(else_idx + 4, "        P: infoMapP,\r\n")
    lines.insert(else_idx + 5, "        S: infoMapS\r\n")
    lines.insert(else_idx + 6, "      };\r\n")
    print(f"✅ infoMap 통합 코드 추가: {else_idx + 1}행")

# ============================================
# 2. onEdit() 함수 수정
# ============================================

# "for (var dbKey in infoMap) {" 찾기
onedit_idx = -1
for i, line in enumerate(lines):
    if 'for (var dbKey in infoMap) {' in line and i > 1600:  # onEdit 함수 내부
        onedit_idx = i
        break

if onedit_idx == -1:
    print("❌ onEdit infoMap 검색 부분을 찾을 수 없습니다!")
    exit(1)

print(f"✅ onEdit 검색 로직 시작: {onedit_idx + 1}행")

# "var foundInfo = null;" 부터 "if (foundInfo) {" 전까지 교체
# 역방향으로 "var foundInfo = null;" 찾기
foundinfo_idx = -1
for i in range(onedit_idx - 1, max(0, onedit_idx - 10), -1):
    if 'var foundInfo = null;' in lines[i]:
        foundinfo_idx = i
        break

# "if (foundInfo) {" 찾기
if_foundinfo_idx = -1
for i in range(onedit_idx, min(len(lines), onedit_idx + 20)):
    if 'if (foundInfo) {' in lines[i]:
        if_foundinfo_idx = i
        break

if foundinfo_idx != -1 and if_foundinfo_idx != -1:
    # 해당 구간 교체
    new_search_code = [
        "                  var foundInfo = null;\r\n",
        "                  var foundGroup = \"\";\r\n",
        "                  \r\n",
        "                  // [수정] 그룹별 맵 구조에 맞게 검색\r\n",
        "                  // infoMap = {M: {...}, P: {...}, S: {...}}\r\n",
        "                  var groupNames = ['M', 'P', 'S'];\r\n",
        "                  \r\n",
        "                  var found = false;\r\n",
        "                  for (var gi = 0; gi < groupNames.length && !found; gi++) {\r\n",
        "                    var groupName = groupNames[gi];\r\n",
        "                    var groupMap = infoMap[groupName];\r\n",
        "                    \r\n",
        "                    if (!groupMap) continue;\r\n",
        "                    \r\n",
        "                    for (var dbKey in groupMap) {\r\n",
        "                      for (var k = 0; k < targetKeywords.length; k++) {\r\n",
        "                        var uKey = targetKeywords[k];\r\n",
        "                        if (dbKey.includes(uKey) || uKey.includes(dbKey)) {\r\n",
        "                          foundInfo = groupMap[dbKey];\r\n",
        "                          foundGroup = groupName;\r\n",
        "                          found = true;\r\n",
        "                          break;\r\n",
        "                        }\r\n",
        "                      }\r\n",
        "                      if (found) break;\r\n",
        "                    }\r\n",
        "                  }\r\n",
        "                  \r\n",
    ]
    
    lines[foundinfo_idx:if_foundinfo_idx] = new_search_code
    print(f"✅ onEdit 검색 로직 교체: {foundinfo_idx + 1}~{if_foundinfo_idx}행")
    
    # 새로운 if_foundinfo_idx 계산
    new_if_idx = foundinfo_idx + len(new_search_code)
    
    # "if (foundInfo) {" 다음 줄에 로그 추가
    lines.insert(new_if_idx + 2, "                      // 그룹 정보 로그\r\n")
    lines.insert(new_if_idx + 3, "                      Logger.log(\"  [BA 자동입력] 행 \" + row + \": 매칭됨 (그룹: \" + foundGroup + \") → 상태: \" + foundInfo.statusText);\r\n")
    lines.insert(new_if_idx + 4, "                      \r\n")
    
    # toast 메시지에 groupInfo 추가
    for i in range(new_if_idx, min(len(lines), new_if_idx + 20)):
        if 'e.source.toast("품목 정보: " + foundInfo.statusText, "정보");' in lines[i]:
            lines[i] = lines[i].replace(', "정보");', ' + " (그룹: " + foundGroup + ")", "정보");')
        elif 'e.source.toast("품목 정보: 단종 제품입니다.", "정보");' in lines[i]:
            lines[i] = lines[i].replace('단종 제품입니다.", "정보");', '단종 제품입니다. (그룹: " + foundGroup + ")", "정보");')
        elif 'e.source.toast("품목 정보: " + foundInfo.value, "성공");' in lines[i]:
            lines[i] = lines[i].replace(', "성공");', ' + " (그룹: " + foundGroup + ")", "성공");')

# 파일 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door_v1.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f"\n✅ door_v1.gs 수정 완료!")
print(f"최종 라인 수: {len(lines)}")
