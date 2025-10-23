#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, time, re, json, base64
from pathlib import Path
from datetime import datetime, timedelta, timezone

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ------------------------------------------------------------
# 환경 설정
# ------------------------------------------------------------
WAIT = 5
ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"
SCHEDULE_URL = "https://live.ecomm-data.com/schedule/hs"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/19pcFwP2XOVEuHPsr9ITudLDSD1Tzg5RwsL3K6maIJ1U/edit?gid=0#gid=0"
WORKSHEET_NAME = "편성표RAW"

# ------------------------------------------------------------
# 유틸
# ------------------------------------------------------------
def make_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("user-agent=Mozilla/5.0 Chrome/122.0.0.0 Safari/537.36")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(60)
    return driver

def save_debug(driver, tag: str):
    ts = int(time.time())
    png = ARTIFACT_DIR / f"{ts}_{tag}.png"
    html = ARTIFACT_DIR / f"{ts}_{tag}.html"
    try:
        driver.save_screenshot(str(png))
        with open(html, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"[DEBUG] 저장: {png.name}, {html.name}")
    except Exception as e:
        print(f"[WARN] 디버그 저장 실패: {e}")

# ------------------------------------------------------------
# 로그인 + 세션 초과 팝업 처리
# ------------------------------------------------------------
def login_and_handle_session(driver):
    driver.get("https://live.ecomm-data.com")
    print("[STEP] 메인 페이지 진입 완료")

    login_link = WebDriverWait(driver, WAIT).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
    )
    driver.execute_script("arguments[0].click();", login_link)
    print("[STEP] 로그인 링크 클릭 완료")

    t0 = time.time()
    while "/user/sign_in" not in driver.current_url:
        if time.time() - t0 > WAIT:
            raise Exception("로그인 페이지 진입 실패 (타임아웃)")
        time.sleep(0.5)
    print("✅ 로그인 페이지 진입 완료:", driver.current_url)

    time.sleep(1)
    email_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='email']") if e.is_displayed()][0]
    pw_input    = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
    email_input.clear(); email_input.send_keys(ECOMM_ID)
    pw_input.clear(); pw_input.send_keys(ECOMM_PW)
    time.sleep(0.5)

    form = driver.find_element(By.TAG_NAME, "form")
    login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
    driver.execute_script("arguments[0].click();", login_button)
    print("✅ 로그인 시도!")

    time.sleep(2)
    try:
        session_items = [li for li in driver.find_elements(By.CSS_SELECTOR, "ul > li") if li.is_displayed()]
        if session_items:
            print(f"[INFO] 세션 초과: {len(session_items)}개 → 맨 아래 세션 선택 후 '종료 후 접속'")
            session_items[-1].click()
            time.sleep(1)
            close_btn = driver.find_element(By.XPATH, "//button[text()='종료 후 접속']")
            if close_btn.is_enabled():
                driver.execute_script("arguments[0].click();", close_btn)
                print("✅ '종료 후 접속' 버튼 클릭 완료")
                time.sleep(2)
        else:
            print("[INFO] 세션 초과 안내창 없음")
    except Exception as e:
        print("[WARN] 세션 처리 중 예외(무시):", e)

    time.sleep(2)
    curr = driver.current_url
    email_inputs = driver.find_elements(By.CSS_SELECTOR, "input[name='email']")
    if "/sign_in" in curr and any(e.is_displayed() for e in email_inputs):
        print("❌ 로그인 실패 (폼 그대로 존재함)")
        save_debug(driver, "login_fail")
        raise RuntimeError("로그인 실패")
    print("✅ 로그인 성공 판정! 현재 URL:", curr)
    save_debug(driver, "login_success")

# ------------------------------------------------------------
# 편성표 페이지 크롤링
# ------------------------------------------------------------
def crawl_schedule(driver):
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지 진입 완료")
    time.sleep(3)

    KST = timezone(timedelta(hours=9))
    yesterday = datetime.now(KST).date() - timedelta(days=1)
    date_text = str(yesterday.day)
    print(f"[STEP] 어제 날짜({yesterday}) 선택")

    date_button_xpath = f"//div[text()='{date_text}']"
    date_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, date_button_xpath))
    )
    driver.execute_script("arguments[0].click();", date_button)
    print("✅ 어제 날짜 클릭 완료")
    time.sleep(3)

    tables = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )

    all_data = []
    for table in tables:
        try:
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 7:
                    item = {
                        "방송시간": cols[1].text.strip(),
                        "방송정보": cols[2].text.strip(),
                        "분류": cols[3].text.strip(),
                        "판매량": cols[4].text.strip(),
                        "매출액": cols[5].text.strip(),
                        "상품수": cols[6].text.strip()
                    }
                    all_data.append(item)
        except Exception:
            continue

    df = pd.DataFrame(all_data)
    print(f"✅ 크롤링 완료: {len(df)}행")
    return df

# ------------------------------------------------------------
# 구글 시트 인증
# ------------------------------------------------------------
def gs_client_from_env():
    GSVC_JSON_B64 = os.environ.get("KEY1", "")
    if not GSVC_JSON_B64:
        raise RuntimeError("환경변수 KEY1이 비어있습니다.")
    svc_info = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = Credentials.from_service_account_info(svc_info, scopes=scope)
    return gspread.authorize(creds)

# ------------------------------------------------------------
# 숫자 변환
# ------------------------------------------------------------
def _to_int_kor(s):
    if not s or s == "-": return 0
    t = str(s).replace(",", "").replace(" ", "")
    unit = {"억": 100_000_000, "만": 10_000}
    for k, v in unit.items():
        if k in t:
            try:
                return int(float(t.split(k)[0]) * v)
            except:
                pass
    try:
        return int(float(t))
    except:
        return 0

# ------------------------------------------------------------
# ✅ 데이터 전처리 추가 (기준가치 시트 참조)
# ------------------------------------------------------------
def preprocess_dataframe(df, sh):
    print("🧮 데이터 전처리 시작")

    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)
    df["방송시간 절대시"] = 60.0
    df["분리송출구분"] = "일반"

    try:
        df["방송날짜"] = pd.to_datetime(df["방송시간"].str.split("\n").str[0], errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = df["방송시간"].str.split("\n").str[1]
    except Exception:
        pass

    try:
        df["일자"] = pd.to_datetime(df["방송날짜"]).dt.day.astype(str) + "일"
        df["시간대"] = pd.to_datetime(df["방송시작시간"]).dt.hour
    except Exception:
        df["일자"] = ""
        df["시간대"] = 0

    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=ref_values[0])
        ref_df["시간대"] = ref_df["시간대"].astype(int)
        ref_df["환산가치"] = ref_df["환산가치"].astype(float)
        df = df.merge(ref_df, on="시간대", how="left")
    except Exception as e:
        print("⚠️ 기준가치 시트 로드 실패:", e)
        df["환산가치"] = 1.0

    df["분리송출고려환산가치"] = df.apply(
        lambda x: x["환산가치"] * 0.5 if x["분리송출구분"] == "분리송출" else x["환산가치"], axis=1)
    df["주문효율 /h"] = df.apply(
        lambda x: x["매출액 환산수식"] / (x["방송시간 절대시"] * x["분리송출고려환산가치"])
        if x["방송시간 절대시"] and x["분리송출고려환산가치"] else 0, axis=1)

    print("✅ 데이터 전처리 완료")
    return df

# ------------------------------------------------------------
# 메인
# ------------------------------------------------------------
def main():
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")

    driver = make_driver()
    try:
        login_and_handle_session(driver)
        df = crawl_schedule(driver)

        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)

        # ✅ 데이터 전처리 수행
        df = preprocess_dataframe(df, sh)

        # ✅ RAW 시트 업로드
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df.columns))
        data_to_upload = [df.columns.tolist()] + df.astype(str).values.tolist()
        ws.clear()
        ws.update(values=data_to_upload, range_name="A1")
        print(f"✅ 편성표RAW 업로드 완료 (행수: {len(data_to_upload)})")

    except Exception as e:
        import traceback
        print("❌ 자동화 실패:", e)
        print(traceback.format_exc())
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
