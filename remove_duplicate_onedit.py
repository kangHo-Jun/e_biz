# 중복된 onEdit 함수 삭제 (1631~1878행)
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"원본 라인 수: {len(lines)}")

# 1631~1878행 삭제 (인덱스 1630~1877)
del lines[1630:1878]

print(f"삭제 후 라인 수: {len(lines)}")

# 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✅ 중복 onEdit 함수 삭제 완료!")
