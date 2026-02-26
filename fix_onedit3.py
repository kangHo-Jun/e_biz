# UTF-8로 파일 읽기
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 1703~1730행 다시 교체
new_section = [
    "                  var foundInfo = null;\r\n",
    "                  var foundGroup = \"\";\r\n",
    "                  \r\n",
    "                  // [수정] 그룹별 맵 구조에 맞게 검색\r\n",
    "                  // infoMap = {M: {...}, P: {...}, S: {...}}\r\n",
    "                  var groupNames = ['M', 'P', 'S'];\r\n",
    "                  \r\n",
    "                  outerSearch:\r\n",
    "                  for (var gi = 0; gi < groupNames.length; gi++) {\r\n",
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
    "                          break outerSearch;\r\n",
    "                        }\r\n",
    "                      }\r\n",
    "                    }\r\n",
    "                  }\r\n",
]

# 1703~1730행 교체 (인덱스 1702~1729)
lines[1702:1730] = new_section

# 1732행 이후에 로그 추가 (foundInfo 체크 후)
# 현재 1732행이 "if (foundInfo) {" 이므로 그 다음에 추가
lines.insert(1733, "                      // 그룹 정보 로그\r\n")
lines.insert(1734, "                      Logger.log(\"  [BA 자동입력] 행 \" + row + \": 매칭됨 (그룹: \" + foundGroup + \") → 상태: \" + foundInfo.statusText);\r\n")
lines.insert(1735, "                      \r\n")

# 파일 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✅ 수정 완료!")
print(f"총 라인 수: {len(lines)}")
