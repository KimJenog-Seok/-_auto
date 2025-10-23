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
# 드라이버 설정
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
    except Exception as e:
        print(f"[WARN] 디버그 저장 실패: {e}")

# ------------------------------------------------------------
# 로그인 (기존 성공본 유지)
# ------------------------------------------------------------
def login_and_handle_session(driver):
    driver.get("https://live.ecomm-data.com")
    print("[STEP] 메인 페이지 진입 완료")

    login_link = WebDriverWait(driver, WAIT).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
    )
    driver.execute_script("arguments[0].click();", login_link)

    # 로그인 페이지 진입 대기
    t0 = time.time()
    while "/user/sign_in" not in driver.current_url:
        if time.time() - t0 > WAIT:
            raise Exception("로그인 페이지 진입 실패 (타임아웃)")
        time.sleep(0.5)
    print("✅ 로그인 페이지 진입 완료:", driver.current_url)

    time.sleep(1)
    email_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='email']") if e.is_displayed()][0]
    pw_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
    email_input.clear(); email_input.send_keys(ECOMM_ID)
    pw_input.clear(); pw_input.send_keys(ECOMM_PW)

    form = driver.find_element(By.TAG_NAME, "form")
    login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
    driver.execute_script("arguments[0].click();", login_button)
    print("✅ 로그인 시도!")

    # 세션 초과 처리
    time.sleep(2)
    try:
        session_items = [li for li in driver.find_elements(By.CSS_SELECTOR, "ul > li") if li.is_displayed()]
        if session_items:
            print("[INFO] 세션 초과 감지 → '종료 후 접속'")
            session_items[-1].click()
            time.sleep(1)
            close_btn = driver.find_element(By.XPATH, "//button[text()='종료 후 접속']")
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(2)
    except Exception as e:
        print("[WARN] 세션 처리 중 예외(무시):", e)

    print("✅ 로그인 성공 완료")

# ------------------------------------------------------------
# 편성표 크롤링
# ------------------------------------------------------------
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
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 7:
                    all_data.append({
                        "방송시간": cols[1].text.strip(),
                        "방송정보": cols[2].text.strip(),
                        "분류": cols[3].text.strip(),
                        "판매량": cols[4].text.strip(),
                        "매출액": cols[5].text.strip(),
                        "상품수": cols[6].text.strip()
                    })
        except Exception:
            continue

    df = pd.DataFrame(all_data)
    print(f"✅ 크롤링 완료 ({len(df)}행)")
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
# ✅ 데이터 전처리
# ------------------------------------------------------------
def preprocess_dataframe(df, sh):
    print("🧮 데이터 전처리 시작")

    # --- 날짜/시간 분리 ---
    split_result = df["방송시간"].str.split("\n", n=1, expand=True)
    df["방송날짜"] = pd.to_datetime(split_result[0].str.strip(), errors="coerce").dt.strftime("%Y-%m-%d")
    df["방송시작시간"] = split_result[1].str.strip()
    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # --- 기준가치 시트 참조 (기준시간 × N일 구조) ---
    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=[c.strip() for c in ref_values[0]])
        ref_df["기준시간"] = ref_df["기준시간"].astype(str).str.strip()

        df["일자"] = pd.to_datetime(df["방송날짜"]).dt.day.astype(str) + "일"
        df["시간대"] = pd.to_datetime(df["방송시작시간"], format="%H:%M", errors="coerce").dt.hour.astype(str)

        def lookup_value(row):
            h = row["시간대"]
            d = row["일자"]
            try:
                val = ref_df.loc[ref_df["기준시간"] == h, d].values
                if len(val) > 0:
                    return float(val[0])
            except Exception:
                pass
            return 0.0

        df["환산가치"] = df.apply(lookup_value, axis=1)
        print("✅ 기준가치 시트 매핑 완료")

    except Exception as e:
        print("⚠️ 기준가치 시트 로드 오류:", e)
        df["환산가치"] = 0.0

    # --- 종료시간 계산 (회사명 없이 순차기준 + 2시간 제한) ---
    df_sorted = df.sort_values("방송시작시간").reset_index()
    start_times = pd.to_datetime(df_sorted["방송시작시간"], format="%H:%M", errors="coerce")
    next_times = start_times.shift(-1)

    end_times = []
    for i, st in enumerate(start_times):
        if pd.isna(st):
            end_times.append(pd.NaT)
            continue
        et = next_times.iloc[i]
        if pd.isna(et) or et <= st:
            et = st + timedelta(hours=0, minutes=90)
        if (et - st) > timedelta(hours=2):
            et = st + timedelta(hours=2)
        end_times.append(et)

    df_sorted["종료시간"] = [t.strftime("%H:%M") if pd.notna(t) else "" for t in end_times]
    df = df_sorted.drop(columns=["index"])

    # --- 방송시간 절대시 (HH:MM 포맷, 2시간 초과 시 2시간 고정) ---
    def fmt_duration(start, end):
        if pd.isna(start) or pd.isna(end):
            return "00:00"
        delta = end - start
        if delta < timedelta(0):
            delta = timedelta(0)
        if delta > timedelta(hours=2):
            delta = timedelta(hours=2)
        total_min = int(delta.total_seconds() // 60)
        hh = total_min // 60
        mm = total_min % 60
        return f"{hh:02d}:{mm:02d}"

    df["방송시간 절대시"] = [
        fmt_duration(s, e)
        for s, e in zip(start_times, end_times)
    ]

    # --- 분리송출 구분 및 환산가치 나누기 ---
    grp_counts = df.groupby(["방송시작시간"]).transform("size")
    df["분리송출구분"] = grp_counts.apply(lambda x: "분리송출" if x > 1 else "일반")
    df["분리송출고려환산가치"] = df["환산가치"] / grp_counts.clip(lower=1)

    # --- 주문효율/h = 매출액 환산수식 ÷ 분리송출고려환산가치 ---
    def safe_eff(sales, adj_val):
        try:
            if adj_val and float(adj_val) != 0.0:
                return float(sales) / float(adj_val)
        except:
            pass
        return 0

    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)

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
        df = preprocess_dataframe(df, sh)

        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df.columns))

        data_to_upload = [df.columns.tolist()] + df.astype(str).values.tolist()
        ws.clear()
        ws.update(values=data_to_upload, range_name="A1")
        print(f"✅ 편성표RAW 업로드 완료 ({len(df)}행)")

    except Exception as e:
        import traceback
        print("❌ 자동화 실패:", e)
        print(traceback.format_exc())
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
