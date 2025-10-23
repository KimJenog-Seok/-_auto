#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ====== 표준/외부 모듈 ======
import os, sys, time, re, json, base64
from pathlib import Path
from datetime import datetime, timedelta, timezone

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ====== Selenium ======
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ------------------------------------------------------------
# 환경 설정 (main2.py 원본 유지)
# ------------------------------------------------------------
WAIT = 5
ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

# 로그인 계정 (요청에 따라 하드코딩 유지)
ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"

# 편성표 URL로 변경
SCHEDULE_URL = "https://live.ecomm-data.com/schedule/hs"

# 구글 시트 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/19pcFwP2XOVEuHPsr9ITudLDSD1Tzg5RwsL3K6maIJ1U/edit?gid=0#gid=0"
WORKSHEET_NAME = "편성표RAW"

# ------------------------------------------------------------
# 유틸 (main2.py 원본 유지)
# ------------------------------------------------------------
def make_driver():
    """GitHub Actions/서버/로컬 공용 크롬 드라이버 (Headless)."""
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox") # <- 'add_JArgument' 오타 수정됨
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("user-agent=Mozilla/5.0 Chrome/122.0.0.0 Safari/537.36")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opts)
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
        )
    except Exception:
        pass
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
# 로그인 + 세션 초과 팝업 처리 (main2.py 원본 유지)
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

    # 세션 초과 팝업 처리
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

    # 성공 판정
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
# 편성표 페이지 크롤링 (main2.py 원본 유지)
# ------------------------------------------------------------
def crawl_schedule(driver):
    # '편성표' 메뉴 클릭
    schedule_link = WebDriverWait(driver, WAIT).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'편성표')]"))
    )
    driver.execute_script("arguments[0].click();", schedule_link)
    print("[STEP] '편성표' 메뉴 클릭 완료")
    time.sleep(2)

    # URL에 이미 홈쇼핑 구분값('/hs')이 있으므로, '홈쇼핑' 버튼 클릭 로직은 건너뜀
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지로 직접 이동 완료")
    time.sleep(2)

    # '하루 전 날짜' 클릭
    KST = timezone(timedelta(hours=9))
    yesterday = datetime.now(KST).date() - timedelta(days=1)
    date_text = str(yesterday.day)
    print(f"[STEP] 어제 날짜({yesterday.strftime('%Y-%m-%d')}) 선택 (UI 표시: '{date_text}')")

    date_button_xpath = f"//div[text()='{date_text}']"
    date_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, date_button_xpath))
    )
    driver.execute_script("arguments[0].click();", date_button)
    print("✅ '하루 전 날짜' 클릭 완료")
    time.sleep(3)

    # 데이터 크롤링 (모든 테이블에서 데이터 추출)
    tables = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )

    all_data = []
    # 챗봇 코드와 달리 원본 컬럼명 유지
    columns = ['방송시간', '방송정보', '분류', '판매량', '매출액', '상품수']

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
                        "분류":   cols[3].text.strip(),
                        "판매량":  cols[4].text.strip(),
                        "매출액":  cols[5].text.strip(),
                        "상품수":  cols[6].text.strip()
                    }
                    all_data.append(item)
                else:
                    print(f"[WARN] 불완전한 데이터 행 발견 (열 개수: {len(cols)})")
                    continue
        except Exception as e:
            print(f"[WARN] 테이블 처리 중 오류 발생 (무시): {e}")
            continue

    df = pd.DataFrame(all_data, columns=columns)
    print(df.head())
    print(f"총 {len(df)}개 편성표 정보 추출 완료")
    return df

# ------------------------------------------------------------
# Google Sheets 인증 (main2.py 원본 유지)
# ------------------------------------------------------------
def gs_client_from_env():
    GSVC_JSON_B64 = os.environ.get("KEY1", "")
    if not GSVC_JSON_B64:
        raise RuntimeError("환경변수 KEY1이 비어있습니다(Base64 인코딩된 서비스계정 JSON 필요).")
    try:
        svc_info = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    except Exception as e:
        print("[WARN] 서비스계정 Base64 디코딩 실패:", e)
        raise

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = Credentials.from_service_account_info(svc_info, scopes=scope)
    return gspread.authorize(creds)

# ------------------------------------------------------------
# 플랫폼 매핑 및 유틸 (main2.py 원본 유지)
# ------------------------------------------------------------
PLATFORM_MAP = {
    "CJ온스타일":"Live","CJ온스타일 플러스":"TC","GS홈쇼핑":"Live","GS홈쇼핑 마이샵":"TC",
    "KT알파쇼핑":"TC","NS홈쇼핑":"Live","NS홈쇼핑 샵플러스":"TC","SK스토아":"TC",
    "공영쇼핑":"Live","롯데원티비":"TC","롯데홈쇼핑":"Live","쇼핑엔티":"TC",
    "신세계쇼핑":"TC","현대홈쇼핑":"Live","현대홈쇼핑 플러스샵":"TC","홈앤쇼핑":"Live",
}
PLATFORMS_BY_LEN = sorted(PLATFORM_MAP.keys(), key=len, reverse=True)

def make_yesterday_title_kst():
    KST = timezone(timedelta(hours=9))
    today = datetime.now(KST).date()
    yday = today - timedelta(days=1)
    return f"{yday.month}/{yday.day}"

def unique_sheet_title(sh, base):
    title = base; n = 1
    while True:
        try:
            sh.worksheet(title)
            n += 1; title = f"{base}-{n}"
        except gspread.exceptions.WorksheetNotFound:
            return title

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

def _to_int_kor(s):
    if s is None:
        return 0
    t = str(s).strip()
    if t == "" or t == "-":
        return 0
    t = t.replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        return int(float(t))
    unit_map = {"억": 100_000_000, "만": 10_000}
    m = re.fullmatch(r"(-?\d+(?:\.\d+)?)(억|만)", t)
    if m:
        return int(float(m.group(1)) * unit_map[m.group(2)])
    total = 0; rest = t
    if "억" in rest:
        parts = rest.split("억")
        try: total += int(float(parts[0]) * unit_map["억"])
        except: pass
        rest = parts[1] if len(parts) > 1 else ""
    if "만" in rest:
        parts = rest.split("만")
        try: total += int(float(parts[0]) * unit_map["만"])
        except: pass
        rest = parts[1] if len(parts) > 1 else ""
    if re.fullmatch(r"-?\d+", rest):
        total += int(rest)
    if total == 0:
        nums = re.findall(r"-?\d+", t)
        return int(nums[0]) if nums else 0
    return total

def format_sales(v):
    try: v = int(v)
    except: return str(v)
    return f"{v/100_000_000:.2f}억"

def format_num(v):
    try: v = int(v)
    except: return str(v)
    return f"{v:,}"

def _agg_two(df, group_cols):
    g = (df.groupby(group_cols, dropna=False)
            .agg(매출합=("매출액_int","sum"),
                 판매량합=("판매량_int","sum"))
            .reset_index()
            .sort_values("매출합", ascending=False))
    return g

def _format_df_table(df):
    d = df.copy()
    d["매출합"] = d["매출합"].apply(format_sales)
    d["판매량합"] = d["판매량합"].apply(format_num)
    return [d.columns.tolist()] + d.astype(str).values.tolist()

# ------------------------------------------------------------
# ★ 신규: 데이터 전처리 (챗봇 코드 통합) ★
# ------------------------------------------------------------
def preprocess_dataframe(df_raw, sh):
    """크롤링한 원본 DF를 받아 18개열의 최종 DF로 전처리합니다."""
    print("🧮 데이터 전처리 시작")
    df = df_raw.copy()

    # 0) 방송날짜/시작시간 분리
    split_result = df["방송시간"].str.split("\n", n=1, expand=True)
    if len(split_result.columns) == 2:
        df["방송날짜"]     = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = split_result[1].str.strip()
    else:
        df["방송날짜"]     = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = ""
    
    # 어제 날짜 확인 (종료시간 계산용)
    try:
        day = pd.to_datetime(df["방송날짜"].iloc[0]).date()
    except Exception:
        KST = timezone(timedelta(hours=9))
        day = datetime.now(KST).date() - timedelta(days=1)


    # 1) 방송정보에서 회사명/홈쇼핑구분 분리 (상품명 = 클린된 방송정보)
    titles, companies, kinds = [], [], []
    for txt in df["방송정보"].astype(str):
        title, comp, kind = split_company_from_broadcast(txt)
        titles.append(title); companies.append(comp); kinds.append(kind)
    df["상품명"] = titles       # <- 이것이 정석님의 "방송정보" 열이 됩니다.
    df["회사명"] = companies
    df["홈쇼핑구분"] = kinds

    # 2) 매출액 환산
    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # 3) 기준가치 매핑 (기준시간 × N일 구조)
    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=[c.strip() for c in ref_values[0]])
        # 열 이름 공백 제거
        ref_df.rename(columns=lambda c: c.strip(), inplace=True)
        
        # '기준시간' 열 이름이 약간 달라도 보정
        if "기준시간" not in ref_df.columns:
            for c in list(ref_df.columns):
                if c.replace(" ", "") == "기준시간":
                    ref_df.rename(columns={c: "기준시간"}, inplace=True)
                    break
        
        ref_df["기준시간"] = ref_df["기준시간"].astype(str).str.strip()

        df["일자"] = pd.to_datetime(df["방송날짜"]).dt.day.astype(str) + "일"
        df["시간대"] = pd.to_datetime(df["방송시작시간"], format="%H:%M", errors="coerce").dt.hour.astype(str)

        def lookup_value(row):
            h = row["시간대"]
            d = row["일자"]
            try:
                # .loc[행조건, 열조건]
                val = ref_df.loc[ref_df["기준시간"] == h, d].values
                if len(val) > 0 and str(val[0]).strip() != "":
                    # 엑셀 값에 쉼표가 있을 수 있으므로 제거
                    return float(str(val[0]).replace(",", ""))
            except Exception:
                pass
            return 0.0

        df["환산가치"] = df.apply(lookup_value, axis=1)
        print("✅ 기준가치 시트 매핑 완료")
    except Exception as e:
        print(f"⚠️ '기준가치' 시트 로드 또는 매핑 오류: {e}")
        df["환산가치"] = 0.0

    # 4) 종료시간 계산 (같은 회사의 다음 방송 시작시각, 없으면 24:30; 최대 2시간 캡)
    def to_dt(hhmm):
        try:
            h, m = map(int, str(hhmm).split(":"))
            return datetime.combine(day, datetime.min.time()) + timedelta(hours=h, minutes=m)
        except Exception:
            return pd.NaT

    df["_start_dt"] = df["방송시작시간"].apply(to_dt)
    # 회사별 정렬 → 다음 방송시각
    df_sorted = df.sort_values(["회사명", "_start_dt"]).reset_index()
    df_sorted["_next_same"] = df_sorted.groupby("회사명")["_start_dt"].shift(-1)
    next_same_map = dict(zip(df_sorted["index"], df_sorted["_next_same"]))
    df["_next_same"] = df.index.map(next_same_map)

    def decide_end(row):
        st = row["_start_dt"]
        et = row["_next_same"]
        if pd.isna(st):
            return pd.NaT
        if pd.isna(et):
            # 마지막 방송: 24:30 (다음날 00:30)
            et = datetime.combine(day, datetime.min.time()) + timedelta(days=1, minutes=30)
        # 최대 2시간 제한
        if et - st > timedelta(hours=2):
            et = st + timedelta(hours=2)
        return et

    df["_end_dt"] = df.apply(decide_end, axis=1)

    def format_end(end_dt):
        if isinstance(end_dt, datetime):
            base0 = datetime.combine(day, datetime.min.time())
            # 24:30분인지 체크
            if (end_dt - base0) >= timedelta(days=1, minutes=30):
                return "24:30"
            return end_dt.strftime("%H:%M")
        return ""
    df["종료시간"] = df["_end_dt"].apply(format_end)

    # 5) 방송시간 절대시 = 종료-시작 (HH:MM), 2시간 초과 시 2시간
    def fmt_duration(st, et):
        if pd.isna(st) or pd.isna(et):
            return "00:00"
        delta = et - st
        if delta < timedelta(0):
            delta = timedelta(0)
        # 2시간 캡은 decide_end에서 이미 적용됨
        total_min = int(delta.total_seconds() // 60)
        hh = total_min // 60
        mm = total_min % 60
        return f"{hh:02d}:{mm:02d}"

    df["방송시간 절대시"] = df.apply(lambda r: fmt_duration(r["_start_dt"], r["_end_dt"]), axis=1)

    # 6) 분리송출 판정 + 환산가치 나누기 (회사명+방송시작시간 기준)
    grp_counts = df.groupby(["회사명", "방송시작시간"])["방송시작시간"].transform("size")
    df["분리송출구분"] = grp_counts.apply(lambda x: "분리송출" if x > 1 else "일반")
    df["분리송출고려환산가치"] = df["환산가치"] / grp_counts.clip(lower=1)

    # 7) 주문효율/h = 매출액 환산수식 ÷ 분리송출고려환산가치
    def safe_eff(sales, adj):
        try:
            adjf = float(adj)
            if adjf != 0.0:
                return float(sales) / adjf
        except:
            pass
        return 0.0
    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)

    # 8) 최종 18개 열 선택 및 순서 지정
    final_cols = [
        # A-I (기존)
        "방송날짜", 
        "방송시작시간",
        "상품명",       # <- C열: 클린된 방송정보
        "분류", 
        "판매량", 
        "매출액", 
        "상품수",
        "회사명", 
        "홈쇼핑구분",
        # J-R (신규)
        "매출액 환산수식",
        "일자",
        "시간대",
        "환산가치",
        "종료시간",
        "방송시간 절대시",
        "분리송출구분",
        "분리송출고려환산가치",
        "주문효율 /h"
    ]
    
    # 혹시 모를 누락 방지
    for c in final_cols:
        if c not in df.columns:
            df[c] = ""
            
    df_final = df[final_cols]
    
    # C열의 "상품명"을 "방송정보"로 변경 (main2.py의 서식/집계 호환성)
    df_final = df_final.rename(columns={"상품명": "방송정보"})

    print("✅ 데이터 전처리 완료 (18개 열 생성)")
    return df_final


# ------------------------------------------------------------
# 구글 시트 서식 지정 (★ 'row_count' 범위 오류 수정 ★)
# ------------------------------------------------------------
def apply_formatting(sh, new_ws, ins_ws, data_row_count):
    # ★★★★★
    # 'new_ws.row_count' 대신, main()에서 계산한
    # 실제 데이터 행 수 'data_row_count'를 받도록 수정
    # ★★★★★
    import traceback
    try:
        reqs = []
        
        # --- A:R (18개 열) 기준으로 설정 ---
        col_count = 18
        # ★★★★★ 'new_ws.row_count' 대신 전달받은 'data_row_count' 사용
        row_count = data_row_count 

        # 1. '어제 날짜' 시트 서식 (A:R 적용)
        
        # (수정됨) 전체 셀에 테두리 (A1:R[row_count])
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": col_count},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"},
            }
        })
        
        # (수정됨) 열 너비 기본 100 (A:R)
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": col_count},
                "properties": {"pixelSize": 100},
                "fields": "pixelSize"
            }
        })
        
        # (기존) C열 600
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
                "properties": {"pixelSize": 600},
                "fields": "pixelSize"
            }
        })
        
        # (기존) H, I열 130
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 9}, # H, I열
                "properties": {"pixelSize": 130},
                "fields": "pixelSize"
            }
        })

        # --- J, Q, R 열 너비 추가 ---
        # J열 (idx 9) 160
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 9, "endIndex": 10},
                "properties": {"pixelSize": 160},
                "fields": "pixelSize"
            }
        })
        # Q열 (idx 16) 150
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 16, "endIndex": 17},
                "properties": {"pixelSize": 150},
                "fields": "pixelSize"
            }
        })
        # R열 (idx 17) 120
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 17, "endIndex": 18},
                "properties": {"pixelSize": 120},
                "fields": "pixelSize"
            }
        })
        
        # --- 정렬 ---
        # (기존) C열 데이터 왼쪽 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 2, "endIndex": 3},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        # (기존) A,B열 가운데 정렬 (헤더+데이터)
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endIndex": 2},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        
        # (수정됨) D~R열 가운데 정렬 (헤더+데이터)
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 3, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        
        # (수정됨) 헤더(A1:R1) 배경색 및 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"
            }
        })

        # --- J열, R열 숫자 서식 (천단위 콤마, 헤더 제외) ---
        number_format_req = {
            "repeatCell": {
                "range": {
                    "sheetId": new_ws.id,
                    "startRowIndex": 1, # 헤더 제외
                    "endRowIndex": row_count # ★★★★★ 범위 수정
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": "#,##0" # 소수점 없음
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        }
        # J열 (idx 9)
        req_j = json.loads(json.dumps(number_format_req)) # 템플릿 복사
        req_j["repeatCell"]["range"]["startColumnIndex"] = 9
        req_j["repeatCell"]["range"]["endColumnIndex"] = 10
        reqs.append(req_j)
        
        # R열 (idx 17)
        req_r = json.loads(json.dumps(number_format_req)) # 템플릿 복사
        req_r["repeatCell"]["range"]["startColumnIndex"] = 17
        req_r["repeatCell"]["range"]["endColumnIndex"] = 18
        reqs.append(req_r)


        # --- 2. 'INS_전일' 시트 서식 (기존 원본과 동일) ---
        # ★★★★★ (참고) 'INS_전일'은 어차피 행이 40개 미만이라 
        # 'ins_ws.row_count'를 써도 아무 문제가 없습니다.
        # ★★★★★
        ins_col_count = 3
        
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": ins_ws.id, "startRowIndex": 0, "endRowIndex": ins_ws.row_count, "startColumnIndex": 0, "endColumnIndex": ins_ws.col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": ins_ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 300},
                "fields": "pixelSize"
            }
        })
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": ins_ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 3},
                "properties": {"pixelSize": 250},
                "fields": "pixelSize"
            }
        })
        
        # A2:C4
        reqs.append({"updateBorders": {"range": {"sheetId": ins_ws.id, "startRowIndex": 1, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"}, "right": {"style": "SOLID"}, "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"}}})
        reqs.append({"repeatCell": {"range": {"sheetId": ins_ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"}})
        
        # A7:C23
        reqs.append({"updateBorders": {"range": {"sheetId": ins_ws.id, "startRowIndex": 6, "endRowIndex": 23, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"}, "right": {"style": "SOLID"}, "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"}}})
        reqs.append({"repeatCell": {"range": {"sheetId": ins_ws.id, "startRowIndex": 6, "endRowIndex": 7, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"}})

        # A26:C36
        reqs.append({"updateBorders": {"range": {"sheetId": ins_ws.id, "startRowIndex": 25, "endRowIndex": 36, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"}, "right": {"style": "SOLID"}, "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"}}})
        
        # --- 'blue: 8.0' 오타 수정된 부분 ---
        reqs.append({"repeatCell": {
            "range": {"sheetId": ins_ws.id, "startRowIndex": 25, "endRowIndex": 26, "startColumnIndex": 0, "endColumnIndex": ins_col_count}, 
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}}, # 8.0 -> 0.8
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"
        }})
        
        sh.batch_update({"requests": reqs})
        print(f"✅ 서식 적용 완료 (적용 행 수: {row_count})")
    except Exception as e:
        print(f"⚠️ 서식 적용 실패: {e}")
        print(traceback.format_exc())

# ------------------------------------------------------------
# 메인 (★ 'row_count' 전달하도록 수정 ★)
# ------------------------------------------------------------
def main():
    # 로컬 테스트용 KEY1 환경 변수 설정
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            json_bytes = f.read()
            os.environ["KEY1"] = base64.b64encode(json_bytes).decode("utf-8")
            print("✅ 로컬 테스트용 KEY1 환경 변수 설정 완료")
    
    driver = make_driver()
    sh = None
    worksheet = None
    new_ws = None
    gc = None # gc를 try 블록 이전에 선언
    
    # ★★★★★ 서식 적용에 필요한 실제 행 수를 저장할 변수
    actual_row_count = 0 
    
    try:
        # 1) 로그인 + 세션 팝업 처리
        login_and_handle_session(driver)

        # 2) 편성표 페이지 크롤링
        df_raw = crawl_schedule(driver)

        # 3) 구글 시트 인증
        gc = gs_client_from_env() # gc에 할당
        sh = gc.open_by_url(SPREADSHEET_URL)
        print("[GS] 스프레드시트 열기 OK")
        
        # 4) ★ 신규 ★ 데이터 전처리 (J-R열 추가)
        print("[STEP] 데이터 전처리 시작 (시간이 걸릴 수 있습니다)...")
        df_processed = preprocess_dataframe(df_raw, sh)
        print("[STEP] 데이터 전처리 완료.")

        # 5) '편성표RAW' 시트 확보(없으면 생성)
        try:
            worksheet = sh.worksheet(WORKSHEET_NAME)
            print("[GS] 기존 워크시트 찾음:", WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df_processed.columns)) # 18개열
            print("[GS] 워크시트 생성:", WORKSHEET_NAME)
        
        # 6) 메인 시트 업로드 (★ 전처리된 18개 열 데이터로 업로드 ★)
        data_to_upload = [df_processed.columns.tolist()] + df_processed.astype(str).values.tolist()
        worksheet.clear()
        worksheet.update(values=data_to_upload, range_name="A1")
        print(f"✅ 구글시트 '편성표RAW' 업로드 완료 (행수: {len(data_to_upload)}, 열수: {len(df_processed.columns)})")

        # 7) 어제 날짜 새 시트 생성 & 값 복사 (★ 전처리된 18개 열 데이터가 복사됨 ★)
        base_title = make_yesterday_title_kst()
        target_title = unique_sheet_title(sh, base_title)
        
        # ★★★★★ 'source_values'를 여기서 읽어와서 실제 행 수를 계산
        source_values = worksheet.get_all_values() or [[""]] # 18개 열 데이터
        
        # ★★★★★ 'actual_row_count' 변수에 실제 데이터 행 수 저장 (헤더 포함)
        actual_row_count = max(2, len(source_values))
        
        cols_cnt = max(2, max(len(r) for r in source_values))
        
        new_ws = sh.add_worksheet(title=target_title, rows=actual_row_count, cols=cols_cnt)
        new_ws.update("A1", source_values)
        print(f"✅ 어제 날짜 시트 생성/복사 완료 → {target_title} (행: {actual_row_count})")
        
        # 8) ★ 삭제 ★ : 방송정보 말미 회사명 제거 로직
        
        # 9) 집계 시트 생성 (main2.py 원본 유지)
        # ★★★★★ 'values'를 'source_values'로 재활용 (이미 읽어왔음)
        values = source_values 
        if not values or len(values) < 2:
            raise Exception("INS_전일 생성 실패: 데이터 행이 없습니다.")
        header = values[0]; body = values[1:]
        df_ins = pd.DataFrame(body, columns=header)
        for col in ["판매량","매출액","홈쇼핑구분","회사명","분류"]:
            if col not in df_ins.columns: df_ins[col] = ""
        df_ins["판매량_int"] = df_ins["판매량"].apply(_to_int_kor)
        df_ins["매출액_int"] = df_ins["매출액"].apply(_to_int_kor)
        
        gubun_tbl = _agg_two(df_ins, ["홈쇼핑구분"])
        plat_tbl  = _agg_two(df_ins, ["회사명"])
        cat_tbl   = _agg_two(df_ins, ["분류"])
        sheet_data = []
        sheet_data.append(["[LIVE/TC 집계]"]); sheet_data += _format_df_table(gubun_tbl); sheet_data.append([""])
        sheet_data.append(["[플랫폼(회사명) 집계]"]); sheet_data += _format_df_table(plat_tbl); sheet_data.append([""])
        sheet_data.append(["[상품분류(분류) 집계]"]); sheet_data += _format_df_table(cat_tbl)
        
        # INS_전일 upsert (main2.py 원본 유지)
        TARGET_TITLE = "INS_전일"
        try:
            ins_ws = sh.worksheet(TARGET_TITLE)
            ins_ws.clear()
            print("[GS] INS_전일 기존 워크시트 찾음 → 초기화")
        except gspread.exceptions.WorksheetNotFound:
            rows_cnt = max(2, len(sheet_data))
            cols_cnt = max(2, max(len(r) for r in sheet_data))
            ins_ws = sh.add_worksheet(title=TARGET_TITLE, rows=rows_cnt, cols=cols_cnt)
            print("[GS] INS_전일 워크시트 생성")
        
        ins_ws.update("A1", sheet_data)
        print("✅ INS_전일 생성/갱신 완료")
        
        # 10) --- 서식 적용 함수 호출 ---
        
        # 토큰 만료 해결을 위한 재-로그인
        print("[STEP] 인증 토큰 갱신 (서식 적용 전)...")
        if gc: # gc가 None이 아닌지 확인
            gc.login()
        print("✅ 인증 토큰 갱신 완료")
        
        # ★★★★★ apply_formatting 호출 시 'actual_row_count' 전달
        apply_formatting(sh, new_ws, ins_ws, actual_row_count)
        # --- 서식 적용 함수 호출 끝 ---

        # 11) 탭 순서 재배치 (main2.py 원본 유지)
        try:
            all_ws_now = sh.worksheets()
            new_order = [ins_ws]
            if new_ws.id != ins_ws.id:
                new_order.append(new_ws)
            for w in all_ws_now:
                if w.id not in (ins_ws.id, new_ws.id):
                    new_order.append(w)
            sh.reorder_worksheets(new_order)
            print("✅ 시트 순서 재배치 완료: INS_전일=1번째, 어제시트=2번째")
        except Exception as e:
            print("⚠️ 시트 순서 재배치 오류:", e)

        print("🎉 전체 파이프라인 완료")
    except Exception as e:
        import traceback
        print("❌ 전체 자동화 과정 중 에러 발생:", e)
        print(traceback.format_exc())
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
