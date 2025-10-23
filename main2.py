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
    except Exception as e:
        print(f"[WARN] 디버그 저장 실패: {e}")

# ------------------------------------------------------------
# 로그인 및 편성표 크롤링
# ------------------------------------------------------------
def login_and_handle_session(driver):
    driver.get("https://live.ecomm-data.com")
    login_link = WebDriverWait(driver, WAIT).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
    )
    driver.execute_script("arguments[0].click();", login_link)
    time.sleep(1)

    email_input = driver.find_element(By.NAME, "email")
    pw_input = driver.find_element(By.NAME, "password")
    email_input.send_keys(ECOMM_ID)
    pw_input.send_keys(ECOMM_PW)
    driver.find_element(By.XPATH, "//button[contains(text(),'로그인')]").click()
    time.sleep(3)
    print("✅ 로그인 완료")

def crawl_schedule(driver):
    driver.get(SCHEDULE_URL)
    time.sleep(3)
    KST = timezone(timedelta(hours=9))
    yesterday = datetime.now(KST).date() - timedelta(days=1)
    date_text = str(yesterday.day)
    date_button_xpath = f"//div[text()='{date_text}']"
    date_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, date_button_xpath))
    )
    driver.execute_script("arguments[0].click();", date_button)
    time.sleep(3)

    tables = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )

    all_data = []
    for table in tables:
        try:
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 7:
                    all_data.append({
                        "방송시간": cols[1].text.strip(),
                        "방송정보": cols[2].text.strip(),
                        "분류": cols[3].text.strip(),
                        "판매량": cols[4].text.strip(),
                        "매출액": cols[5].text.strip(),
                        "상품수": cols[6].text.strip(),
                    })
        except Exception:
            continue

    df = pd.DataFrame(all_data)
    print(f"✅ 편성표 크롤링 완료: {len(df)}행")
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
# 문자열→정수 변환
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
# 데이터 전처리 (엑셀 수식 반영)
# ------------------------------------------------------------
def preprocess_dataframe(df, sh):
    print("🧮 데이터 전처리 시작")

    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    def calc_end_time(row):
        try:
            start = datetime.strptime(str(row["방송시작시간"]), "%H:%M:%S")
            duration_min = 60
            return (start + timedelta(minutes=duration_min)).strftime("%H:%M:%S")
        except Exception:
            return ""
    df["종료시간"] = df.apply(calc_end_time, axis=1)
    df["방송시간 절대시"] = 60.0
    df["분리송출구분"] = "일반"

    try:
        df["일자"] = pd.to_datetime(df["방송날짜"]).dt.day.astype(str) + "일"
        df["시간대"] = pd.to_datetime(df["방송시작시간"]).dt.hour
    except Exception:
        df["일자"], df["시간대"] = "", 0

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
        lambda x: x["환산가치"] * 0.5 if str(x["분리송출구분"]).strip() == "분리송출" else x["환산가치"], axis=1)
    df["주문효율 /h"] = df.apply(
        lambda x: x["매출액 환산수식"] / (x["방송시간 절대시"] * x["분리송출고려환산가치"])
        if x["방송시간 절대시"] and x["분리송출고려환산가치"] else 0, axis=1)

    print("✅ 데이터 전처리 완료 (추가 컬럼 반영)")
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

        # ✅ 전처리 실행
        df = preprocess_dataframe(df, sh)

        # RAW 시트 업로드
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
