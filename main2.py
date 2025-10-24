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

# ===================== 설정 =====================
WAIT = 5
ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"
SCHEDULE_URL = "https://live.ecomm-data.com/schedule/hs"

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/19pcFwP2XOVEuHPsr9ITudLDSD1Tzg5RwsL3K6maIJ1U/edit?gid=0#gid=0"
WORKSHEET_NAME = "편성표RAW"

# ===================== 유틸 =====================
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

# ===================== 로그인/세션 =====================
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

    time.sleep(2)
    curr = driver.current_url
    email_inputs = driver.find_elements(By.CSS_SELECTOR, "input[name='email']")
    if "/sign_in" in curr and any(e.is_displayed() for e in email_inputs):
        print("❌ 로그인 실패 (폼 그대로 존재함)")
        save_debug(driver, "login_fail")
        raise RuntimeError("로그인 실패")
    print("✅ 로그인 성공 판정! 현재 URL:", curr)
    save_debug(driver, "login_success")

# ===================== 크롤링 =====================
def crawl_schedule(driver):
    # 메뉴 클릭 생략하고 바로 URL 이동
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지로 직접 이동 완료")
    time.sleep(2)

    # 어제 날짜 클릭 (간단 구현: UI text 기준)
    KST = timezone(timedelta(hours=9))
    yesterday = datetime.now(KST).date() - timedelta(days=1)
    date_text = str(yesterday.day)
    print(f"[STEP] 어제 날짜 선택: {yesterday} → '{date_text}'")

    date_button_xpath = f"//div[text()='{date_text}']"
    date_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, date_button_xpath))
    )
    driver.execute_script("arguments[0].click();", date_button)
    print("✅ '하루 전 날짜' 클릭 완료")
    time.sleep(3)

    tables = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )

    all_data = []
    columns = ['방송시간', '방송정보', '분류', '판매량', '매출액', '상품수']

    for table in tables:
        try:
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 7:
                    
                    # 💡 [수정] cols[1] (방송시간) 내부의 span 2개를 찾아 '\n'으로 연결
                    try:
                        spans = cols[1].find_elements(By.TAG_NAME, "span")
                        if len(spans) == 2:
                            broadcast_time = f"{spans[0].text.strip()}\n{spans[1].text.strip()}"
                        else:
                            # <span>이 2개가 아닌 경우 (예상치 못한 구조) 대비
                            broadcast_time = cols[1].text.strip()
                    except Exception:
                        # 예외 발생 시 기존 방식(텍스트 통째로) 사용
                        broadcast_time = cols[1].text.strip()

                    item = {
                        "방송시간": broadcast_time, # 💡 수정된 broadcast_time 사용
                        "방송정보": cols[2].text.strip(),
                        "분류":   cols[3].text.strip(),
                        "판매량":  cols[4].text.strip(),
                        "매출액":  cols[5].text.strip(),
                        "상품수":  cols[6].text.strip()
                    }
                    all_data.append(item)
                else:
                    continue
        except Exception:
            continue

    df = pd.DataFrame(all_data, columns=columns)
    print(f"총 {len(df)}개 편성표 정보 추출 완료")
    return df

# ===================== Google Sheets 인증 =====================
def gs_client_from_env():
    GSVC_JSON_B64 = os.environ.get("KEY1", "")
    if not GSVC_JSON_B64:
        raise RuntimeError("환경변수 KEY1이 비어있습니다(Base64 인코딩된 서비스계정 JSON 필요).")
    svc_info = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = Credentials.from_service_account_info(svc_info, scopes=scope)
    return gspread.authorize(creds)

# ===================== 보조 유틸/매핑 =====================
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
    return text, "", "" # 💡 맵에 없으면 TC가 아닌 빈칸("") 반환 (기존 로직)

def _to_int_kor(s):
    # 안전한 한글 단위 변환 (빈값/하이픈/콤마/공백 대응)
    if s is None:
        return 0
    t = str(s).strip()
    if t == "" or t == "-":
        return 0
    t = t.replace(",", "").replace(" ", "")
    # 순수 숫자 또는 소수 → 정수화
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        return int(float(t))
    unit_map = {"억": 100_000_000, "만": 10_000}
    m = re.fullmatch(r"(-?\d+(?:\.\d+)?)(억|만)", t)
    if m:
        return int(float(m.group(1)) * unit_map[m.group(2)])
    # 혼합형 처리: 1억2만3000, 0.5억 등
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

# ===================== 전처리 =====================
def preprocess_dataframe(df_raw, sh):
    print("🧮 데이터 전처리 시작")
    df = df_raw.copy()

    # 방송날짜/시작시간 분리 (💡 crawl_schedule 수정으로 \n이 보장됨)
    split_result = df["방송시간"].str.split("\n", n=1, expand=True)
    if len(split_result.columns) == 2:
        df["방송날짜"]     = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = split_result[1].str.strip()
    else:
        # 💡 (Fallback) \n이 여전히 없는 경우 (예: crawl_schedule에서 예외 발생)
        df["방송날짜"]     = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = ""
        print("⚠️ 일부 데이터에서 날짜/시간 분리 실패 (\\n 없음)")

    # 어제 날짜(종료시간 계산용)
    try:
        day = pd.to_datetime(df["방송날짜"].iloc[0]).date()
    except Exception:
        KST = timezone(timedelta(hours=9))
        day = datetime.now(KST).date() - timedelta(days=1)

    # 방송정보에서 회사명/구분 분리
    titles, companies, kinds = [], [], []
    for txt in df["방송정보"].astype(str):
        title, comp, kind = split_company_from_broadcast(txt)
        titles.append(title); companies.append(comp); kinds.append(kind)
    df["상품명"] = titles
    df["회사명"] = companies
    df["홈쇼핑구분"] = kinds

    # 매출액 환산
    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # 기준가치 매핑
    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=[c.strip() for c in ref_values[0]])
        ref_df.rename(columns=lambda c: c.strip(), inplace=True)
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
                val = ref_df.loc[ref_df["기준시간"] == h, d].values
                if len(val) > 0 and str(val[0]).strip() != "":
                    return float(str(val[0]).replace(",", ""))
            except Exception:
                pass
            return 0.0

        df["환산가치"] = df.apply(lookup_value, axis=1)
        print("✅ 기준가치 시트 매핑 완료")
    except Exception as e:
        print(f"⚠️ '기준가치' 시트 로드 또는 매핑 오류: {e}")
        df["환산가치"] = 0.0

    # 종료시간 계산
    def to_dt(hhmm):
        try:
            h, m = map(int, str(hhmm).split(":"))
            return datetime.combine(day, datetime.min.time()) + timedelta(hours=h, minutes=m)
        except Exception:
            return pd.NaT

    df["_start_dt"] = df["방송시작시간"].apply(to_dt)
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
            et = datetime.combine(day, datetime.min.time()) + timedelta(days=1, minutes=30)
        if et - st > timedelta(hours=2):
            et = st + timedelta(hours=2)
        return et

    df["_end_dt"] = df.apply(decide_end, axis=1)

    def format_end(end_dt):
        if isinstance(end_dt, datetime):
            base0 = datetime.combine(day, datetime.min.time())
            if (end_dt - base0) >= timedelta(days=1, minutes=30):
                return "24:30"
            return end_dt.strftime("%H:%M")
        return ""
    df["종료시간"] = df["_end_dt"].apply(format_end)

    # 방송시간 절대시
    def fmt_duration(st, et):
        if pd.isna(st) or pd.isna(et):
            return "00:00"
        delta = et - st
        if delta < timedelta(0):
            delta = timedelta(0)
        total_min = int(delta.total_seconds() // 60)
        hh = total_min // 60
        mm = total_min % 60
        return f"{hh:02d}:{mm:02d}"

    df["방송시간 절대시"] = df.apply(lambda r: fmt_duration(r["_start_dt"], r["_end_dt"]), axis=1)

    # 분리송출
    grp_counts = df.groupby(["회사명", "방송시작시간"])["방송시작시간"].transform("size")
    df["분리송출구분"] = grp_counts.apply(lambda x: "분리송출" if x > 1 else "일반")
    df["분리송출고려환산가치"] = df["환산가치"] / grp_counts.clip(lower=1)

    # 주문효율
    def safe_eff(sales, adj):
        try:
            adjf = float(adj)
            if adjf != 0.0:
                return float(sales) / adjf
        except:
            pass
        return 0.0
    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)

    final_cols = [
        "방송날짜","방송시작시간","상품명","분류","판매량","매출액","상품수","회사명","홈쇼핑구분",
        "매출액 환산수식","일자","시간대","환산가치","종료시간","방송시간 절대시","분리송출구분","분리송출고려환산가치","주문효율 /h"
    ]
    for c in final_cols:
        if c not in df.columns:
            df[c] = ""
    df_final = df[final_cols].rename(columns={"상품명": "방송정보"})
    print("✅ 데이터 전처리 완료 (18개 열 생성)")
    return df_final

# ===================== 서식 적용 =====================
def apply_formatting(sh, new_ws, ins_ws, data_row_count):
    import traceback
    try:
        reqs = []
        col_count = 18
        row_count = data_row_count

        # A1:R(row_count) 테두리
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": col_count},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"},
            }
        })
        # 전체 기본 열 너비
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": col_count},
                "properties": {"pixelSize": 100},
                "fields": "pixelSize"
            }
        })
        # C열 600
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
                "properties": {"pixelSize": 600},
                "fields": "pixelSize"
            }
        })
        # H,I열 130
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 9},
                "properties": {"pixelSize": 130},
                "fields": "pixelSize"
            }
        })
        # J, Q, R 열 너비
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 9, "endIndex": 10},
                "properties": {"pixelSize": 160},
                "fields": "pixelSize"
            }
        })
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 16, "endIndex": 17},
                "properties": {"pixelSize": 150},
                "fields": "pixelSize"
            }
        })
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id, "dimension": "COLUMNS", "startIndex": 17, "endIndex": 18},
                "properties": {"pixelSize": 120},
                "fields": "pixelSize"
            }
        })

        # C열 왼쪽 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": 2, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        # A,B 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": 2},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        # D~R 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 3, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        # 헤더 배경/정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endIndex": 1, "startColumnIndex": 0, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"
            }
        })
        
        # 💡 [오전 수정] 숫자 서식: J, R (콤마O, 소수점X 정수)
        def number_format_req(col_idx):
            return {
                "repeatCell": {
                    "range": {"sheetId": new_ws.id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": col_idx, "endColumnIndex": col_idx+1},
                    "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}}, # "1,000" 형태
                    "fields": "userEnteredFormat.numberFormat"
                }
            }
        reqs.append(number_format_req(9))   # J
        reqs.append(number_format_req(17))  # R

        # INS_전일 간단 정렬(기존과 동일)
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": ins_ws.id, "startRowIndex": 0, "endRowIndex": ins_ws.row_count, "startColumnIndex": 0, "endColumnIndex": ins_ws.col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

        # ✅ gspread 표준 방식: dict에 "requests" 키로 전달
        sh.batch_update({"requests": reqs})
        print(f"✅ 서식 적용 완료 (적용 행 수: {row_count})")
    except Exception as e:
        print(f"⚠️ 서식 적용 실패: {e}")
        print(traceback.format_exc())

# ===================== 메인 =====================
def main():
    # 로컬 테스트용 KEY1 자동 주입(있을 때만)
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")
            print("✅ 로컬 테스트용 KEY1 환경 변수 설정 완료")

    driver = None
    try:
        driver = make_driver()

        # 1) 로그인
        login_and_handle_session(driver)

        # 2) 크롤링 (💡 crawl_schedule 수정됨)
        df_raw = crawl_schedule(driver)

        # 3) 구글 시트 인증/오픈
        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)
        print("[GS] 스프레드시트 열기 OK")

        # 4) 전처리 (💡 preprocess_dataframe이 수정된 데이터 처리)
        print("[STEP] 데이터 전처리 시작...")
        df_processed = preprocess_dataframe(df_raw, sh)
        print("[STEP] 데이터 전처리 완료.")

        # 5) RAW 시트 upsert (💡 정렬 안 함, fillna 적용)
        try:
            worksheet = sh.worksheet(WORKSHEET_NAME)
            print("[GS] 기존 워크시트 찾음:", WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df_processed.columns))
            print("[GS] 워크시트 생성:", WORKSHEET_NAME)

        # 💡 [오전 수정] .fillna("") 사용 (숫자 타입 유지)
        df_for_upload = df_processed.fillna("")
        data_to_upload = [df_for_upload.columns.tolist()] + df_for_upload.values.tolist()
        
        worksheet.clear()
        # 💡 [오전 수정] 경고 로그 해결 (명명된 인수 사용)
        worksheet.update(values=data_to_upload, range_name="A1")
        print(f"✅ 구글시트 '편성표RAW' 업로드 완료 (행수: {len(data_to_upload)}, 열수: {len(df_processed.columns)})")


        # 6) 💡 [오후 수정] 어제 날짜 시트 생성 (정렬 추가)
        base_title = make_yesterday_title_kst()
        target_title = unique_sheet_title(sh, base_title)

        print(f"[STEP] 백업 시트 정렬 수행: 회사명(오름차순), 방송시작시간(오름차순)")
        # 💡 정렬 수행
        df_sorted_backup = df_processed.sort_values(
            by=["회사명", "방송시작시간"], 
            ascending=[True, True]
        )

        # 💡 정렬된 데이터프레임을 업로드용 리스트로 변환
        df_backup_upload = df_sorted_backup.fillna("")
        source_values_sorted = [df_backup_upload.columns.tolist()] + df_backup_upload.values.tolist()
        
        actual_row_count = max(2, len(source_values_sorted))
        cols_cnt = max(2, max(len(r) for r in source_values_sorted))

        new_ws = sh.add_worksheet(title=target_title, rows=actual_row_count, cols=cols_cnt)
        
        # 💡 [오전 수정] 경고 로그 해결 + 정렬된(source_values_sorted) 데이터로 업로드
        new_ws.update(values=source_values_sorted, range_name="A1")
        print(f"✅ 어제 날짜 시트 생성/복사/정렬 완료 → {target_title} (행: {actual_row_count})")


        # 7) INS_전일 요약 시트 (💡 정렬되지 않은 원본 RAW 데이터 사용)
        
        # 💡 'INS_전일' 집계는 정렬 전 원본(data_to_upload)을 사용
        values = data_to_upload 
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

        TARGET_TITLE = "INS_전일"
        try:
            ins_ws = sh.worksheet(TARGET_TITLE)
            ins_ws.clear()
            print("[GS] INS_전일 기존 워크시트 찾음 → 초기화")
        except gspread.exceptions.WorksheetNotFound:
            rows_cnt = max(2, len(sheet_data))
            cols_cnt2 = max(2, max(len(r) for r in sheet_data))
            ins_ws = sh.add_worksheet(title=TARGET_TITLE, rows=rows_cnt, cols=cols_cnt2)
            print("[GS] INS_전일 워크시트 생성")
            
        # 💡 [오전 수정] 경고 로그 해결
        ins_ws.update(values=sheet_data, range_name="A1")
        print("✅ INS_전일 생성/갱신 완료")

        # 8) 서식 적용
        time.sleep(1)
        new_ws = sh.worksheet(target_title)
        print(f"[STEP] 서식 적용 시작 (총 {actual_row_count} 행 대상)...")
        apply_formatting(sh, new_ws, ins_ws, actual_row_count)

        # 9) 탭 순서
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
        try:
            if driver is not None:
                driver.quit()
        except:
            pass

if __name__ == "__main__":
    main()
