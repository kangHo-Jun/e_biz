import json, os, time
from slack_reader import get_messages
from sheet_writer import write_to_sheet
from config import CHECK_INTERVAL

PROCESSED_FILE = os.path.join(os.path.dirname(__file__), "processed_ids.json")

def load_processed():
    if os.path.exists(PROCESSED_FILE):
        try:
            with open(PROCESSED_FILE, "r") as f:
                return set(json.load(f))
        except:
            return set()
    return set()

def save_processed(ids):
    try:
        with open(PROCESSED_FILE, "w") as f:
            json.dump(list(ids), f)
    except Exception as e:
        print(f"[처리 ID 저장 오류] {e}")

def main():
    print("[시작] 슬랙 자동화 모니터링 시작")
    processed = load_processed()

    while True:
        try:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 메시지 체크 중...")
            messages = get_messages()
            new_count = 0
            for msg in messages:
                if msg["ts"] not in processed:
                    write_to_sheet(msg["text"])
                    processed.add(msg["ts"])
                    new_count += 1
            
            if new_count > 0:
                save_processed(processed)
                print(f"[완료] {new_count}개의 새로운 메시지를 처리했습니다.")
            else:
                print("[알림] 새로운 메시지가 없습니다.")
                
            print(f"[대기] {CHECK_INTERVAL}초 후 다시 확인합니다...")
        except Exception as e:
            print(f"[오류 발생] {e}")
            
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    import sys
    from slack_reader import get_recent_messages
    
    if "--test" in sys.argv:
        print("[테스트] 최근 메시지 10개 가져오기")
        msgs = get_recent_messages(10)
        for i, m in enumerate(msgs):
            print(f"\n[{i+1}] {m['text'][:100]}")
        
        print("\n구글 시트에 입력할까요? (y/n)")
        if input().lower() == "y":
            for m in msgs:
                write_to_sheet(m["text"])
    else:
        main()
