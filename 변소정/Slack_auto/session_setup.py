from playwright.sync_api import sync_playwright
import json, os, time

COOKIE_FILE = os.path.join(os.path.dirname(__file__), "cookies.json")

def run_setup():
    print("="*50)
    print("슬랙 세션 설정 (시스템 실제 Chrome 사용)")
    print("="*50)
    print("1. 이제 PC에 설치된 실제 구글 크롬 브라우저가 열립니다.")
    print("2. 슬랙에 로그인해 주세요. (2단계 인증 포함)")
    print("3. 채널 메시지 화면까지 진입한 후 터미널을 확인해 주세요.")
    print("="*50)
    
    with sync_playwright() as p:
        try:
            # 내장 브라우저 대신 PC에 설치된 '진짜 크롬'을 실행합니다.
            # 이 방식은 슬랙의 보안 차단을 가장 확실하게 통과할 수 있습니다.
            browser = p.chromium.launch(headless=False, channel="chrome")
        except Exception as e:
            print(f"[알림] 실제 크롬 앱을 찾을 수 없어 내장 브라우저를 사용합니다.. ({e})")
            browser = p.chromium.launch(headless=False)
            
        context = browser.new_context(
            viewport={'width': 1280, 'height': 800},
            locale="ko-KR",
            timezone_id="Asia/Seoul"
        )
        
        page = context.new_page()
        page.goto("https://slack.com/signin")
        
        print("\n[대기 중] 브라우저에서 로그인을 완료해 주세요...")
        
        while True:
            try:
                # 슬랙 앱 실제 클라이언트 URL 감지
                if "app.slack.com/client" in page.url:
                    print(f"\n[감지] 로그인 성공 확인: {page.url}")
                    time.sleep(5)
                    break
                time.sleep(2)
            except:
                # 브라우저가 수동으로 닫힌 경우 등
                break
        
        # 쿠키 저장
        cookies = context.cookies()
        if cookies:
            with open(COOKIE_FILE, "w") as f:
                json.dump(cookies, f)
            print(f"\n[성공] {COOKIE_FILE} 저장 완료!")
        else:
            print("\n[실패] 쿠키를 가져오지 못했습니다.")
            
        browser.close()

if __name__ == "__main__":
    run_setup()
