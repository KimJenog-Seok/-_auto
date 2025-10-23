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
# 환경 설정 (원본 유지)
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
# 유틸 (원본 유지)
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
# 로그인 + 세션 팝업 처리 (원본 성공 로직 유지)
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
    pw_input    = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
    email_input.clear(); email_input.send_keys(ECOMM_ID)
    pw_input.clear(); pw_input.send_keys(ECOMM_PW)

    form = driver.find_element(By.TAG_NAME, "form")
    login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
    driver.execute_script("arguments[0].click();", login_button)
    print("✅ 로그인 시도!")

    # 동시 세션 종료 후 접속
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
# 편성표 크롤링 (원본 유지)
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
                        "방송시간": cols[1].text.strip(),  # "YYYY.MM.DD\nHH:MM"
                        "방송정보": cols[2].text.strip(),  # "... 회사명"
                        "분류":     cols[3].text.strip(),
                        "판매량":   cols[4].text.strip(),
                        "매출액":   cols[5].text.strip(),
                        "상품수":   cols[6].text.strip()
                    })
        except Exception:
            continue

    df = pd.DataFrame(all_data)
    print(f"✅ 크롤링 완료 ({len(df)}행)")
    return df

# ------------------------------------------------------------
# Google Sheets 인증 (원본 유지)
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
# 숫자 변환 (원본 유지)
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
# 회사명 추출 (방송정보 말미의 플랫폼명 제거/추출)
# ------------------------------------------------------------
PLATFORM_MAP = {
    "CJ온스타일":"Live","CJ온스타일 플러스":"TC","GS홈쇼핑":"Live","GS홈쇼핑 마이샵":"TC",
    "KT알파쇼핑":"TC","NS홈쇼핑":"Live","NS홈쇼핑 샵플러스":"TC","SK스토아":"TC",
    "공영쇼핑":"Live","롯데원티비":"TC","롯데홈쇼핑":"Live","쇼핑엔티":"TC",
    "신세계쇼핑":"TC","현대홈쇼핑":"Live","현대홈쇼핑 플러스샵":"TC","홈앤쇼핑":"Live",
}
PLATFORMS_BY_LEN = sorted(PLATFORM_MAP.keys(), key=len, reverse=True)

def split_company_from_broadcast(text):
    if not text:
        return text, "", ""
    t = text.rstrip()
    for key in PLATFORMS_BY_LEN:
        pattern = r"\s*" + re.escape(key) + r"\s*$"
        if re.search(pattern, t):
            cleaned = re.sub(pattern, "", t).rstrip()
            return cleaned, key, PLATFORM_MAP[key]
    return text, "", ""

# ------------------------------------------------------------
# ✅ 데이터 전처리 (엑셀 수식 동등 구현)
#   - 환산가치 매핑(문자열 기준)
#   - 종료시간/방송시간(최대 2시간 캡)
#   - 분리송출 판정 + 분할
#   - 주문효율 산식
# ------------------------------------------------------------
def preprocess_dataframe(df, sh):
    print("🧮 데이터 전처리 시작")

    # --- 0) 방송날짜/시작시간 분리 ---
    split_result = df['방송시간'].str.split('\n', n=1, expand=True)
    df['방송날짜']    = pd.to_datetime(split_result[0].str.strip(), errors="coerce").dt.strftime("%Y-%m-%d")
    df['방송시작시간'] = split_result[1].str.strip()

    # --- 1) 매출액 환산 ---
    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # --- 2) 회사명 추출(분리송출 판정을 위해 필요) ---
    #    (출력 컬럼에는 포함하지 않지만 계산에는 사용)
    tmp_company = []
    for txt in df["방송정보"].astype(str).tolist():
        _, company, _ = split_company_from_broadcast(txt)
        tmp_company.append(company)
    df["_회사명_TMP"] = tmp_company

    # --- 3) 환산가치 매핑 (기준가치 시트: 문자열 키 매칭 + 공백 제거) ---
    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=[c.strip() for c in ref_values[0]])
        ref_df["시간대"]   = ref_df["시간대"].astype(str).str.strip()
        ref_df["환산가치"] = pd.to_numeric(ref_df["환산가치"], errors="coerce").fillna(method="ffill").fillna(0)
        df["시간대"]       = pd.to_datetime(df["방송시작시간"], format="%H:%M", errors="coerce").dt.hour.astype(str)
        df = df.merge(ref_df, on="시간대", how="left")
        df["환산가치"] = df["환산가치"].fillna(0.0)
    except Exception as e:
        print("⚠️ 기준가치 시트 로드 오류:", e)
        df["환산가치"] = 0.0

    # --- 4) 종료시간 계산 ---
    # 동일 회사 다음 방송의 시작시각을 종료시각으로.
    # (없으면 24:30 고정; 마지막 슬롯 보정. 전체 다음 방송(타사)도 후보가 될 수 있음)
    # 날짜는 동일하다고 가정.
    day = pd.to_datetime(df["방송날짜"]).dt.date.iloc[0] if len(df) else datetime.now().date()

    def to_dt(hhmm):
        try:
            h, m = map(int, str(hhmm).split(":"))
            return datetime.combine(day, datetime.min.time()) + timedelta(hours=h, minutes=m)
        except Exception:
            return pd.NaT

    df["_start_dt"] = df["방송시작시간"].apply(to_dt)

    # 전체 스케줄 기준 "다음 방송 시작시각"
    df_sorted_all = df.sort_values("_start_dt").reset_index()
    df_sorted_all["_next_any"] = df_sorted_all["_start_dt"].shift(-1)
    next_any_map = dict(zip(df_sorted_all["index"], df_sorted_all["_next_any"]))
    df["_next_any"] = df.index.map(next_any_map)

    # 회사별 기준 "다음 방송 시작시각"
    df_sorted_co = df.sort_values(["_회사명_TMP", "_start_dt"]).reset_index()
    df_sorted_co["_next_same"] = df_sorted_co.groupby("_회사명_TMP")["_start_dt"].shift(-1)
    next_same_map = dict(zip(df_sorted_co["index"], df_sorted_co["_next_same"]))
    df["_next_same"] = df.index.map(next_same_map)

    # 종료시각 결정: 우선 같은 회사의 다음 방송, 없으면 전체 다음 방송, 그것도 없으면 24:30
    def decide_end(row):
        end_dt = row["_next_same"]
        if pd.isna(end_dt):
            end_dt = row["_next_any"]
        if pd.isna(end_dt):
            # 마지막 슬롯: 24:30 (= 다음날 00:30)
            end_dt = datetime.combine(day, datetime.min.time()) + timedelta(days=1, hours=0, minutes=30)
        # 최대 2시간 제한
        if not pd.isna(row["_start_dt"]) and (end_dt - row["_start_dt"]) > timedelta(hours=2):
            end_dt = row["_start_dt"] + timedelta(hours=2)
        return end_dt

    df["_end_dt"] = df.apply(decide_end, axis=1)

    # 종료시각 텍스트: 24:30은 특수 표기, 그 외는 HH:MM
    def format_end(end_dt):
        # 다음날 00:30은 24:30으로 표기
        if isinstance(end_dt, datetime):
            base = datetime.combine(day, datetime.min.time())
            if (end_dt - base) == timedelta(days=1, minutes=30):
                return "24:30"
            return end_dt.strftime("%H:%M")
        return ""
    df["종료시간"] = df["_end_dt"].apply(format_end)

    # --- 5) 방송시간 절대시 (종료-시작, HH:MM 포맷 / 2시간 cap 반영) ---
    def fmt_duration(start_dt, end_dt):
        if pd.isna(start_dt) or pd.isna(end_dt):
            return "00:00"
        delta = end_dt - start_dt
        if delta < timedelta(0):
            delta = timedelta(0)
        if delta > timedelta(hours=2):
            delta = timedelta(hours=2)
        total_min = int(delta.total_seconds() // 60)
        hh = total_min // 60
        mm = total_min % 60
        return f"{hh:02d}:{mm:02d}"
    df["방송시간 절대시"] = df.apply(lambda r: fmt_duration(r["_start_dt"], r["_end_dt"]), axis=1)

    # --- 6) 분리송출 판정 + 분할 (COUNTIFS 동등: 회사명+방송시작시간) ---
    grp_cnt = df.groupby(["_회사명_TMP", "방송시작시간"]).transform("size")
    df["분리송출구분"] = grp_cnt.apply(lambda x: "분리송출" if x > 1 else "일반")
    split_counts = grp_cnt.clip(lower=1)  # 최소 1
    df["분리송출고려환산가치"] = df["환산가치"] / split_counts

    # --- 7) 주문효율 /h = 매출액 환산수식 ÷ 분리송출고려환산가치 ---
    def safe_eff(sales, adj_val):
        try:
            if adj_val and float(adj_val) != 0.0:
                return float(sales) / float(adj_val)
        except:
            pass
        return 0
    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)

    # 표시용 부가 컬럼
    df["일자"] = pd.to_datetime(df["방송날짜"], errors="coerce").dt.day.astype("Int64").astype(str) + "일"

    # 업로드 컬럼 순서 구성
    # A:I(원본) + [매출액 환산수식(H 옆에 '종료시간' 추가로 한 칸씩 밀기)]
    # 최종: 방송날짜, 방송시작시간, 방송정보, 분류, 판매량, 매출액, 상품수,
    #       매출액 환산수식, 종료시간, 방송시간 절대시, 분리송출구분, 일자, 시간대, 환산가치, 분리송출고려환산가치, 주문효율 /h
    df["시간대"] = df["시간대"].astype(str)  # 이미 위에서 설정
    final_cols = [
        "방송날짜","방송시작시간","방송정보","분류","판매량","매출액","상품수",
        "매출액 환산수식","종료시간","방송시간 절대시","분리송출구분",
        "일자","시간대","환산가치","분리송출고려환산가치","주문효율 /h"
    ]
    # 누락 컬럼 보호
    for c in final_cols:
        if c not in df.columns:
            df[c] = ""
    df = df[final_cols]

    # 내부 계산용 임시 컬럼 정리
    drop_cols = [c for c in ["_회사명_TMP","_start_dt","_end_dt","_next_any","_next_same"] if c in df.columns]
    df = df.drop(columns=drop_cols, errors="ignore")

    print("✅ 데이터 전처리 완료")
    return df

# ------------------------------------------------------------
# 메인 (원본 흐름 유지 + 전처리 삽입)
# ------------------------------------------------------------
def main():
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")

    driver = make_driver()
    try:
        # 1) 로그인/세션 처리
        login_and_handle_session(driver)

        # 2) 크롤링
        df = crawl_schedule(driver)

        # 3) 구글시트 핸들
        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)

        # 4) ✅ 전처리
        df = preprocess_dataframe(df, sh)

        # 5) RAW 업로드
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
