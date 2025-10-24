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
# 로그인
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
    print("✅ 로그인 페이지 진입 완료")

    email_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='email']") if e.is_displayed()][0]
    pw_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
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
            print("[INFO] 세션 초과 안내창 감지됨 → 종료 후 접속")
            session_items[-1].click()
            time.sleep(1)
            close_btn = driver.find_element(By.XPATH, "//button[text()='종료 후 접속']")
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(2)
    except Exception as e:
        print("[WARN] 세션 처리 중 예외:", e)

    time.sleep(2)
    if "/sign_in" in driver.current_url:
        save_debug(driver, "login_fail")
        raise RuntimeError("로그인 실패")
    print("✅ 로그인 성공!")

# ------------------------------------------------------------
# 편성표 크롤링
# ------------------------------------------------------------
def crawl_schedule(driver):
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지로 이동 완료")

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
    columns = ['방송시간', '방송정보', '분류', '판매량', '매출액', '상품수']

    for table in tables:
        try:
            rows = table.find_elements(By.TAG_NAME, "tr")
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

    df = pd.DataFrame(all_data, columns=columns)
    print(f"총 {len(df)}개 편성표 추출 완료")
    return df

# ------------------------------------------------------------
# Google Sheets 인증
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
# 전처리 함수 (간략화 버전)
# ------------------------------------------------------------
def preprocess_dataframe(df_raw):
    df = df_raw.copy()
    if "매출액" in df.columns:
        df["매출액 환산수식"] = df["매출액"].str.replace("[^0-9]", "", regex=True).astype(float)
    else:
        df["매출액 환산수식"] = 0
    df["일자"] = datetime.now().strftime("%d일")
    df["시간대"] = "0"
    df["환산가치"] = 0.6
    df["종료시간"] = "01:00"
    df["방송시간 절대시"] = "01:00"
    df["분리송출구분"] = "일반"
    df["분리송출고려환산가치"] = 0.6
    df["주문효율 /h"] = df["매출액 환산수식"] / 0.6
    return df

# ------------------------------------------------------------
# 서식 적용 (수정된 완전 작동 버전)
# ------------------------------------------------------------
def apply_formatting(sh, new_ws, ins_ws, data_row_count):
    try:
        reqs = []
        col_count = 18
        row_count = data_row_count

        # 전체 테두리
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": col_count},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"}
            }
        })

        # 열 너비
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": col_count},
                "properties": {"pixelSize": 100},
                "fields": "pixelSize"
            }
        })

        # 헤더 스타일
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"
            }
        })

        # ✅ 핵심 수정: gspread batch_update에 리스트 직접 전달
        sh.batch_update(reqs)

        print(f"✅ 서식 적용 완료 (행 {row_count})")
    except Exception as e:
        print("⚠️ 서식 적용 실패:", e)

# ------------------------------------------------------------
# 메인
# ------------------------------------------------------------
def main():
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")
            print("✅ 로컬 KEY1 설정 완료")

    driver = make_driver()
    try:
        login_and_handle_session(driver)
        df_raw = crawl_schedule(driver)
        df_processed = preprocess_dataframe(df_raw)

        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)

        # RAW 시트 업로드
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df_processed.columns))

        data_to_upload = [df_processed.columns.tolist()] + df_processed.astype(str).values.tolist()
        ws.clear()
        ws.update("A1", data_to_upload)
        print("✅ RAW 시트 업로드 완료")

        # 날짜 시트 생성
        title = datetime.now().strftime("%m/%d")
        new_ws = sh.add_worksheet(title=title, rows=len(data_to_upload)+1, cols=len(df_processed.columns))
        new_ws.update("A1", data_to_upload)
        print("✅ 날짜 시트 생성 완료")

        # INS_전일 시트
        try:
            ins_ws = sh.worksheet("INS_전일")
        except gspread.exceptions.WorksheetNotFound:
            ins_ws = sh.add_worksheet(title="INS_전일", rows=3, cols=3)

        # ✅ 시트 안정화 및 서식 적용
        time.sleep(1)
        new_ws = sh.worksheet(title)
        apply_formatting(sh, new_ws, ins_ws, len(data_to_upload)+1)

    except Exception as e:
        print("❌ 오류:", e)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
