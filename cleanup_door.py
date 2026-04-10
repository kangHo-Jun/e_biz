# 1628행부터 잘못된 코드 모두 삭제
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"원본 라인 수: {len(lines)}")

# 1628행부터 다음 함수 시작 전까지 찾기
# "function " 또는 "/**"로 시작하는 다음 줄 찾기
start_idx = 1627  # 1628행 (0-indexed)
end_idx = -1

for i in range(start_idx, len(lines)):
    line = lines[i].strip()
    if line.startswith('/**') and i > start_idx + 10:  # 최소 10줄 이후
        # 다음 줄이 function인지 확인
        if i + 1 < len(lines) and 'function' in lines[i + 1]:
            end_idx = i
            break

if end_idx == -1:
    print("❌ 종료 지점을 찾을 수 없습니다!")
    exit(1)

print(f"삭제 범위: {start_idx + 1}행 ~ {end_idx}행")

# 삭제
del lines[start_idx:end_idx]

print(f"삭제 후 라인 수: {len(lines)}")

# 저장
with open(r'c:\Users\DSAI\Desktop\이비즈\door.gs', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("✅ 잘못된 코드 삭제 완료!")
