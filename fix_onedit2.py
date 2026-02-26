# UTF-8로 파일 읽기
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 1703~1714행 교체 (0-indexed이므로 1702~1713)
new_lines = [
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

# 1703~1714행 교체 (인덱스 1702~1713)
lines[1702:1714] = new_lines

# 1719행에 로그 추가 (인덱스 1718)
lines[1718] = "                      // 그룹 정보 로그\r\n"
lines.insert(1719, "                      Logger.log(\"  [BA 자동입력] 행 \" + row + \": 매칭됨 (그룹: \" + foundGroup + \") → 상태: \" + foundInfo.statusText);\r\n")
lines.insert(1720, "                      \r\n")
lines.insert(1721, "                      // 상태 텍스트가 있으면 우선 사용\r\n")

# 파일 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✅ 라인 교체 완료!")
print(f"총 라인 수: {len(lines)}")
