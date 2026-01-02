import os
import time
import datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# GitHub Secrets에서 정보 가져오기
MY_EMAIL = os.environ.get('MY_EMAIL')
APP_PASSWORD = os.environ.get('APP_PASSWORD')
TO_EMAIL = MY_EMAIL # 받는 사람도 나 자신으로 설정

def run_agent():
    print("🚀 GitHub Action 에이전트 실행 시작...")

    # 1. 날짜 계산
    today = datetime.date.today()
    three_days_ago = today - datetime.timedelta(days=3)

    # 2. 헤드리스 브라우저 설정 (서버에는 모니터가 없으므로 필수)
    chrome_options = Options()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # Chrome 버전 이슈 방지를 위한 설정
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        # 3. 사이트 접속
        driver.get("https://knote.kr/EumSupyo.do")
        driver.implicitly_wait(10)
        
        # 4. 기간별 조회 클릭
        radio_btn = driver.find_element(By.XPATH, "//label[contains(text(), '기간별 조회')]/preceding-sibling::input")
        driver.execute_script("arguments[0].click();", radio_btn)
        time.sleep(1)

        # 5. 날짜 입력
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
        
        inputs[0].clear(); inputs[0].send_keys(str(three_days_ago.year))
        inputs[1].clear(); inputs[1].send_keys(str(three_days_ago.month).zfill(2))
        inputs[2].clear(); inputs[2].send_keys(str(three_days_ago.day).zfill(2))

        inputs[3].clear(); inputs[3].send_keys(str(today.year))
        inputs[4].clear(); inputs[4].send_keys(str(today.month).zfill(2))
        inputs[5].clear(); inputs[5].send_keys(str(today.day).zfill(2))

        # 6. 조회 클릭
        search_btn = driver.find_element(By.XPATH, "//a[contains(text(), '조회') or contains(@class, 'btn')]")
        search_btn.click()
        time.sleep(5) # 서버 속도를 고려해 대기 시간 넉넉히

        # 7. 데이터 크롤링
        table = driver.find_element(By.TAG_NAME, "table")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")

        data_list = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) > 1:
                data = {
                    "사업자번호": cols[0].text.strip(),
                    "법인명": cols[1].text.strip(),
                    "성명": cols[2].text.strip(),
                    "주소": cols[3].text.strip(),
                    "정지일": cols[4].text.strip()
                }
                data_list.append(data)
        
        print(f"📊 데이터 {len(data_list)}건 발견")

        # 8. 엑셀 저장 및 메일 전송
        file_name = f"stop_list_{today.strftime('%Y%m%d')}.xlsx"
        if data_list:
            df = pd.DataFrame(data_list)
            df.to_excel(file_name, index=False)
            send_email(file_name)
        else:
            print("데이터 없음 - 메일 미발송")

    except Exception as e:
        print(f"❌ 오류: {e}")
    finally:
        driver.quit()

def send_email(filename):
    msg = MIMEMultipart()
    msg['From'] = MY_EMAIL
    msg['To'] = TO_EMAIL
    msg['Subject'] = f"[자동알림] {datetime.date.today()} 당좌거래정지자"
    msg.attach(MIMEText("첨부파일 확인 바랍니다.", 'plain'))

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(MY_EMAIL, APP_PASSWORD)
    server.send_message(msg)
    server.quit()
    print("📧 전송 완료")

if __name__ == "__main__":
    run_agent()
