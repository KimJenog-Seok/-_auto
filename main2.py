#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, time, re, json, base64
from pathlib import Path
from datetime import datetime, timedelta, timezone

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import a1_to_rowcol 

# 🔥 OpenAI (카테고리 분류용)
from openai import OpenAI
from concurrent.futures import ThreadPoolExecutor, as_completed 

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

# 💡 최종 수정: Assistant ID의 'Z'를 소문자 'z'로 변경하여 NotFoundError 해결 시도
ASSISTANT_ID = "asst_Nd5zLY7wqhsQqigS4YIDU5nL" 

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
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지로 직접 이동 완료")
    time.sleep(2)

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

                    try:
                        spans = cols[1].find_elements(By.TAG_NAME, "span")
                        if len(spans) == 2:
                            broadcast_time = f"{spans[0].text.strip()}\n{spans[1].text.strip()}"
                        else:
                            broadcast_time = cols[1].text.strip()
                    except Exception:
                        broadcast_time = cols[1].text.strip()

                    item = {
                        "방송시간": broadcast_time,
                        "방송정보": cols[2].text.strip(),
                        "분류":    cols[3].text.strip(),
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


# ===================== 전처리 =====================
def preprocess_dataframe(df_raw, sh):
    print("🧮 데이터 전처리 시작")
    df = df_raw.copy()

    # 방송날짜/시간 분리
    split_result = df["방송시간"].str.split("\n", n=1, expand=True)
    if len(split_result.columns) == 2:
        df["방송날짜"]      = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = split_result[1].str.strip()
    else:
        df["방송날짜"]      = pd.to_datetime(split_result[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
        df["방송시작시간"] = ""
        print("⚠️ 일부 데이터는 날짜/시간 분리 실패")

    try:
        day = pd.to_datetime(df["방송날짜"].iloc[0]).date()
    except:
        KST = timezone(timedelta(hours=9))
        day = datetime.now(KST).date() - timedelta(days=1)

    titles, companies, kinds = [], [], []
    for txt in df["방송정보"].astype(str):
        title, comp, kind = split_company_from_broadcast(txt)
        titles.append(title); companies.append(comp); kinds.append(kind)
    df["상품명"] = titles
    df["회사명"] = companies
    df["홈쇼핑구분"] = kinds

    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # 기준가치 매핑
    try:
        기준_ws = sh.worksheet("기준가치")
        ref_values = 기준_ws.get_all_values()
        ref_df = pd.DataFrame(ref_values[1:], columns=[c.strip() for c in ref_values[0]])

        if "기준시간" not in ref_df.columns:
            for c in list(ref_df.columns):
                if c.replace(" ", "") == "기준시간":
                    ref_df.rename(columns={c: "기준시간"}, inplace=True)
                    break

        df["일자"] = pd.to_datetime(df["방송날짜"]).dt.day.astype(str) + "일"
        df["시간대"] = pd.to_datetime(df["방송시작시간"], format="%H:%M", errors="coerce").dt.hour.astype(str)

        def lookup_value(row):
            h = row["시간대"]
            d = row["일자"]
            try:
                val = ref_df.loc[ref_df["기준시간"] == h, d].values
                if len(val) > 0 and str(val[0]).strip() != "":
                    return float(str(val[0]).replace(",", ""))
            except:
                pass
            return 0.0

        df["_시간당_환산가치"] = df.apply(lookup_value, axis=1)
        print("✅ 기준가치 매핑 완료")
    except Exception as e:
        print(f"⚠️ 기준가치 시트 오류 (데이터 품질 문제): {e}")
        df["_시간당_환산가치"] = 0.0

    def to_dt(hhmm):
        try:
            h, m = map(int, str(hhmm).split(":"))
            return datetime.combine(day, datetime.min.time()) + timedelta(hours=h, minutes=m)
        except:
            return pd.NaT

    df["_start_dt"] = df["방송시작시간"].apply(to_dt)

    df_sorted = df.sort_values(["회사명", "_start_dt"])
    df_unique_starts = df_sorted.drop_duplicates(subset=["회사명", "_start_dt"])[["회사명", "_start_dt"]].copy()
    df_unique_starts["_next_unique_start"] = df_unique_starts.groupby("회사명")["_start_dt"].shift(-1)
    df = df.merge(df_unique_starts, on=["회사명","_start_dt"], how="left")

    def decide_end(row):
        st = row["_start_dt"]
        et = row["_next_unique_start"]
        if pd.isna(st):
            return pd.NaT
        if pd.isna(et):
            return datetime.combine(day, datetime.min.time()) + timedelta(days=1, minutes=30)
        if et - st > timedelta(hours=2):
            return st + timedelta(hours=2)
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

    def fmt_duration(st, et):
        if pd.isna(st) or pd.isna(et):
            return "00:00"
        delta = et - st
        if delta < timedelta(0):
            delta = timedelta(0)
        total_min = int(delta.total_seconds() // 60)
        return f"{total_min//60:02d}:{total_min%60:02d}"

    df["방송시간 절대시"] = df.apply(lambda r: fmt_duration(r["_start_dt"], r["_end_dt"]), axis=1)
    
    df["_방송시간(분)"] = df["방송시간 절대시"].apply(lambda v: int(v.split(":")[0])*60 + int(v.split(":")[1]) if ":" in v else 0)

    def calculate_actual_value(row):
        per_hour_value = row["_시간당_환산가치"]
        minutes = row["_방송시간(분)"]
        if per_hour_value == 0.0 or minutes == 0:
            return 0.0
        return (per_hour_value / 60.0) * minutes

    if "환산가치" not in df.columns:
        df["환산가치"] = 0.0
    df["환산가치"] = df.apply(calculate_actual_value, axis=1)

    grp_counts = df.groupby(["회사명", "방송시작시간"])["방송시작시간"].transform("size")
    df["분리송출구분"] = grp_counts.apply(lambda x: "분리송출" if x > 1 else "일반")
    df["분리송출고려환산가치"] = df["환산가치"] / grp_counts.clip(lower=1)

    def safe_eff(sales, adj):
        try:
            adjf = float(adj)
            if adjf != 0.0:
                return float(sales) / adjf
        except:
            pass
        return 0.0

    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)

    # 💡 수정 2: AI분류(S열) 포함하여 19개 열 정의
    final_cols = [
        "방송날짜","방송시작시간","상품명","분류","판매량","매출액","상품수","회사명","홈쇼핑구분",
        "매출액 환산수식","일자","시간대","환산가치","종료시간","방송시간 절대시","분리송출구분",
        "분리송출고려환산가치","주문효율 /h","AI분류" 
    ]
    
    for c in final_cols:
        if c not in df.columns:
            df[c] = ""
    
    if "AI분류" not in df.columns:
        df["AI분류"] = ""

    df_final = df[final_cols].rename(columns={"상품명": "방송정보"})
    print("✅ 데이터 전처리 완료 (19개 열 생성)") 
    return df_final

# ===================== 서식 적용 (A~S 전체) =====================
def apply_formatting(sh, new_ws, ins_ws, data_row_count):
    import traceback
    try:
        reqs = []
        col_count = 19  # A~S 열 (19개)
        row_count = data_row_count

        # A1:S(row_count) 테두리
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id,
                          "startRowIndex": 0, "endRowIndex": row_count,
                          "startColumnIndex": 0, "endColumnIndex": col_count},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"},
            }
        })

        # 기본 열 너비
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id,
                          "dimension": "COLUMNS",
                          "startIndex": 0, "endIndex": col_count},
                "properties": {"pixelSize": 100},
                "fields": "pixelSize"
            }
        })

        # C 열 = 600
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id,
                          "dimension": "COLUMNS",
                          "startIndex": 2, "endIndex": 3},
                "properties": {"pixelSize": 600},
                "fields": "pixelSize"
            }
        })

        # H,I 열 = 130
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": new_ws.id,
                          "dimension": "COLUMNS",
                          "startIndex": 7, "endIndex": 9},
                "properties": {"pixelSize": 130},
                "fields": "pixelSize"
            }
        })

        # J, Q, R, S = 160
        for idx in [9, 16, 17, 18]:
            reqs.append({
                "updateDimensionProperties": {
                    "range": {"sheetId": new_ws.id,
                              "dimension": "COLUMNS",
                              "startIndex": idx, "endIndex": idx+1},
                    "properties": {"pixelSize": 160},
                    "fields": "pixelSize"
                }
            })

        # C열 왼쪽 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id,
                          "startRowIndex": 1, "endRowIndex": row_count,
                          "startColumnIndex": 2, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

        # A,B 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id,
                          "startRowIndex": 0, "endRowIndex": row_count,
                          "startColumnIndex": 0, "endColumnIndex": 2},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

        # D~S 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id,
                          "startRowIndex": 0, "endRowIndex": row_count,
                          "startColumnIndex": 3, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

        # 헤더 스타일(A1:S1)
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": new_ws.id,
                          "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": col_count},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8},
                    "horizontalAlignment": "CENTER",
                }},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"
            }
        })

        # 숫자 서식 적용(J, R)
        def number_format(col_idx):
            return {
                "repeatCell": {
                    "range": {"sheetId": new_ws.id,
                              "startRowIndex": 1, "endRowIndex": row_count,
                              "startColumnIndex": col_idx, "endColumnIndex": col_idx+1},
                    "cell": {"userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
                    }},
                    "fields": "userEnteredFormat.numberFormat"
                }
            }
        reqs.append(number_format(9))    # J
        reqs.append(number_format(17))  # R

        # INS_전일 가운데 정렬
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": ins_ws.id,
                          "startRowIndex": 0, "endRowIndex": ins_ws.row_count,
                          "startColumnIndex": 0, "endColumnIndex": ins_ws.col_count},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

        sh.batch_update({"requests": reqs})
        print(f"✅ 서식 적용 완료 (A~S, {row_count}행)")
    except Exception as e:
        print(f"⚠️ 서식 적용 실패: {e}")
        print(traceback.format_exc())


# ===================== 병렬 카테고리 분류 (100행 제한 제거) =====================
def classify_one_row(client, assistant_id, title, base):
    """
    단일 행 카테고리 분류 함수 (스레드에서 실행)
    """
    try:
        thread = client.beta.threads.create()
        client.beta.threads.messages.create(
            thread_id=thread.id,
            role="user",
            content=f"{title} — {base}"
        )
        run = client.beta.threads.runs.create_and_poll(
            thread_id=thread.id,
            assistant_id=assistant_id
        )
        msgs = client.beta.threads.messages.list(thread_id=thread.id)
        result = msgs.data[0].content[0].text.value.strip()

        # 정제
        result = re.sub(r"[`´]+", "", result)
        result = result.strip()
        result = re.split(r"[—\-–]", result)[-1].strip()
        result = result.splitlines()[0].strip()

        return result

    except Exception as e:
        # e.message가 아닌 type(e).__name__을 반환하여 NotFoundError를 명확히 함
        return f"분류 오류: {type(e).__name__}"


def run_category_classification(sh, target_title):
    """
    병렬(5개)로 전체 행 분류
    """
    print(f"[CAT] 카테고리 분류 대상 시트: {target_title}")

    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
    ASSISTANT_ID_TO_USE = ASSISTANT_ID 

    if not OPENAI_API_KEY:
        raise RuntimeError("❌ OPENAI_API_KEY 환경변수가 없습니다.")

    client = OpenAI(api_key=OPENAI_API_KEY)
    ws = sh.worksheet(target_title)

    rows = ws.get_all_values()
    if not rows or len(rows) < 2:
        print("[CAT] 데이터 없음 → 분류 생략")
        return

    header = rows[0]
    data   = rows[1:]

    total = len(data)
    # 💡 수정: 100행 제한을 제거하고 전체 행을 limit으로 설정
    limit = total 
    print(f"[CAT] 총 {total}개 중 **전체 {limit}개** 병렬 분류 시작")

    results = [""] * total # 전체 행 개수만큼 리스트 초기화
    tasks = []

    with ThreadPoolExecutor(max_workers=5) as executor:
        for idx in range(limit):
            row = data[idx]
            # 인덱스 범위 체크 (C열=2, D열=3)
            title = row[2] if len(row) > 2 else "" 
            base  = row[3] if len(row) > 3 else ""

            print(f"[CAT] 제출 → 행 {idx+2}: {title[:25]}...")

            tasks.append((
                idx,
                executor.submit(classify_one_row, client, ASSISTANT_ID_TO_USE, title, base)
            ))

        for idx, future in tasks:
            results[idx] = future.result()
            print(f"[CAT] 완료 ← 행 {idx+2}") 

    # S열 전체 업데이트 (S2:S끝)
    update_range = f"S2:S{total+1}"
    update_values = [[r] for r in results[0:total]] 

    # 💡 수정 4: gspread 감가상각 경고 해결
    ws.update(range_name=update_range, values=update_values)
    print("🎯 S열 카테고리 병렬 분류 완료 (전체 행)")

# ===================== 메인 파이프라인 =====================
def main():
    # 로컬 테스트용 KEY1 자동 주입 (GitHub에서는 무시됨)
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")
            print("✅ 로컬 KEY1 환경변수 세팅 완료")

    driver = None
    try:
        driver = make_driver()

        # 1) 로그인
        login_and_handle_session(driver)

        # 2) 크롤링
        df_raw = crawl_schedule(driver)

        # 3) Google Sheets 연결
        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)
        print("[GS] 스프레드시트 연결 OK")

        # 4) 전처리
        print("[STEP] 전처리 시작…")
        df_processed = preprocess_dataframe(df_raw, sh)
        print("[STEP] 전처리 완료")

        # 5) RAW 시트 업데이트
        try:
            ws_raw = sh.worksheet(WORKSHEET_NAME)
            print("[GS] RAW 시트 발견")
        except:
            ws_raw = sh.add_worksheet(title=WORKSHEET_NAME,
                                      rows=2,
                                      cols=len(df_processed.columns))
            print("[GS] RAW 시트 생성")

        df_u = df_processed.fillna("")
        payload = [df_u.columns.tolist()] + df_u.values.tolist()

        ws_raw.clear()
        # 💡 수정 5: gspread 감가상각 경고 해결
        ws_raw.update(range_name="A1", values=payload)
        print(f"✅ RAW 업데이트 완료 ({len(payload)}행)")

        # 6) 백업 시트 생성(어제 날짜)
        base_title = make_yesterday_title_kst()
        backup_title = unique_sheet_title(sh, base_title)

        print("[STEP] 백업 시트용 정렬 실행")
        df_sorted = df_processed.sort_values(
            by=["회사명", "방송시작시간"],
            ascending=[True, True]
        )

        df_bu = df_sorted.fillna("")
        bu_values = [df_bu.columns.tolist()] + df_bu.values.tolist()

        rows_cnt = max(2, len(bu_values))
        cols_cnt = max(len(r) for r in bu_values)
        
        # 💡 수정 6: S열(19번째 열)까지 쓰기 위해 최소 19개 열을 확보
        cols_cnt = max(19, cols_cnt) 

        ws_bu = sh.add_worksheet(title=backup_title,
                                 rows=rows_cnt,
                                 cols=cols_cnt)
        # 💡 수정 7: gspread 감가상각 경고 해결
        ws_bu.update(range_name="A1", values=bu_values)
        print(f"✅ 백업 시트 생성 완료 → {backup_title}")

        # 7) INS_전일 생성
        header = payload[0]
        body   = payload[1:]
        df_ins = pd.DataFrame(body, columns=header)

        for c in ["판매량", "매출액", "홈쇼핑구분", "회사명", "분류"]:
            if c not in df_ins.columns:
                df_ins[c] = ""

        df_ins["판매량_int"] = df_ins["판매량"].apply(_to_int_kor)
        df_ins["매출액_int"] = df_ins["매출액"].apply(_to_int_kor)

        tbl1 = _agg_two(df_ins, ["홈쇼핑구분"])
        tbl2 = _agg_two(df_ins, ["회사명"])
        tbl3 = _agg_two(df_ins, ["분류"])

        ins_data = []
        ins_data.append(["[LIVE/TC 집계]"])
        ins_data += _format_df_table(tbl1)
        ins_data.append([""])

        ins_data.append(["[플랫폼(회사명) 집계]"])
        ins_data += _format_df_table(tbl2)
        ins_data.append([""])

        ins_data.append(["[상품분류(분류) 집계]"])
        ins_data += _format_df_table(tbl3)

        max_ins_cols = max(len(r) for r in ins_data)

        try:
            ws_ins = sh.worksheet("INS_전일")
            ws_ins.clear()
            # 💡 INS 시트 크기 재조정 
            if ws_ins.row_count < len(ins_data) or ws_ins.col_count < max_ins_cols:
                 ws_ins.resize(rows=max(2, len(ins_data)), cols=max_ins_cols)
            print("[GS] 기존 INS_전일 초기화")
        except:
            ws_ins = sh.add_worksheet(title="INS_전일",
                                      rows=max(2, len(ins_data)),
                                      cols=max_ins_cols)
            print("[GS] INS_전일 새로 생성")

        # 💡 수정 8: gspread 감가상각 경고 해결
        ws_ins.update(range_name="A1", values=ins_data)
        print("✅ INS_전일 생성/반영 완료")

        # 8) 병렬 카테고리 분류 실행(전체 행)
        print("[STEP] 병렬 카테고리 분류 시작…")
        run_category_classification(sh, backup_title)
        print("🎯 카테고리 분류 완료")

        # 9) 서식 적용(A~S 열 전체)
        print("[STEP] 서식 적용 시작…")
        rows_cnt_bu = ws_bu.row_count
        apply_formatting(sh, ws_bu, ws_ins, rows_cnt_bu)
        print("🎉 서식 적용 완료")

        # 🔟 시트 순서 재배치
        try:
            all_ws = sh.worksheets()
            new_order = [ws_ins, ws_bu]
            for w in all_ws:
                if w.id not in (ws_ins.id, ws_bu.id):
                    new_order.append(w)
            sh.reorder_worksheets(new_order)
            print("📌 시트 순서 재정렬 완료")
        except Exception as e:
            print("⚠️ 시트 순서 재배치 오류:", e)

        print("🎉 전체 파이프라인 완료!")
    except Exception as e:
        import traceback
        print("❌ 전체 파이프라인 오류:", e)
        print(traceback.format_exc())
    finally:
        try:
            if driver:
                driver.quit()
        except:
            pass


if __name__ == "__main__":
    main()
