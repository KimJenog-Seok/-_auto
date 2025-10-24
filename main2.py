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

    # 어제 날짜 클릭 (간단 구현: UI 텍스트 기준)
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
# --- 중략 (기존 코드 동일) ---

def preprocess_dataframe(df_raw, sh):
    print("🧮 데이터 전처리 시작")
    df = df_raw.copy()

    # 매출액 환산수식 보정 (없을 경우 생성)
    if "매출액 환산수식" not in df.columns:
        df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)

    # 주문효율 계산
    def safe_eff(sales, adj):
        try:
            adjf = float(adj)
            if adjf != 0.0:
                return float(sales) / adjf
        except:
            pass
        return 0.0

    df["주문효율 /h"] = df.apply(
        lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1
    )

    # ✅ 소수점 제거 후 정수형으로 반올림
    df["주문효율 /h"] = pd.to_numeric(df["주문효율 /h"], errors="coerce").fillna(0).round().astype(int)

    # 최종 열 순서 지정
    final_cols = [
        "방송날짜","방송시작시간","상품명","분류","판매량","매출액","상품수",
        "회사명","홈쇼핑구분","매출액 환산수식","일자","시간대","환산가치",
        "종료시간","방송시간 절대시","분리송출구분","분리송출고려환산가치","주문효율 /h"
    ]
    for c in final_cols:
        if c not in df.columns:
            df[c] = ""

    df_final = df[final_cols].rename(columns={"상품명": "방송정보"})
    print("✅ 데이터 전처리 완료 (18개 열 생성)")
    return df_final


# ------------------------------------------------------------
# ★★★ RAW / 날짜 / INS_전일 업로드 시 문자열 변환 제거 + USER_ENTERED 추가 ★★★
# ------------------------------------------------------------
def upload_sheets(sh, df_processed):
    # RAW 시트 업로드
    try:
        worksheet = sh.worksheet(WORKSHEET_NAME)
        print("[GS] 기존 워크시트 찾음:", WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df_processed.columns))
        print("[GS] 워크시트 생성:", WORKSHEET_NAME)

    # ✅ 문자열 변환 제거 + USER_ENTERED 적용
    data_to_upload = [df_processed.columns.tolist()] + df_processed.values.tolist()
    worksheet.clear()
    worksheet.update("A1", data_to_upload, value_input_option="USER_ENTERED")
    print(f"✅ RAW 시트 업로드 완료 (행수: {len(data_to_upload)}, 열수: {len(df_processed.columns)})")

    # 날짜 시트 생성
    base_title = make_yesterday_title_kst()
    target_title = unique_sheet_title(sh, base_title)
    source_values = worksheet.get_all_values() or [[""]]
    actual_row_count = max(2, len(source_values))
    cols_cnt = max(2, max(len(r) for r in source_values))

    new_ws = sh.add_worksheet(title=target_title, rows=actual_row_count, cols=cols_cnt)
    # ✅ USER_ENTERED로 업로드
    new_ws.update("A1", source_values, value_input_option="USER_ENTERED")
    print(f"✅ 어제 날짜 시트 생성/복사 완료 → {target_title}")

    # INS_전일 시트
    TARGET_TITLE = "INS_전일"
    try:
        ins_ws = sh.worksheet(TARGET_TITLE)
        ins_ws.clear()
        print("[GS] INS_전일 기존 워크시트 찾음 → 초기화")
    except gspread.exceptions.WorksheetNotFound:
        ins_ws = sh.add_worksheet(title=TARGET_TITLE, rows=3, cols=3)
        print("[GS] INS_전일 워크시트 생성")

    # ✅ USER_ENTERED로 집계 업로드
    ins_ws.update("A1", [["데이터 준비됨"]], value_input_option="USER_ENTERED")

    return worksheet, new_ws, ins_ws, actual_row_count


# ------------------------------------------------------------
# ★ apply_formatting 내 J,R열 콤마 포맷 유지 ★
# ------------------------------------------------------------
def apply_formatting(sh, new_ws, ins_ws, data_row_count):
    try:
        reqs = []
        col_count = 18
        row_count = data_row_count

        # 기본 서식들 (테두리, 정렬 등)
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": col_count},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"},
            }
        })

        # ✅ 숫자 서식: 천단위 콤마(#,##0), 소수점 없음
        def num_format(col):
            return {
                "repeatCell": {
                    "range": {
                        "sheetId": new_ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": row_count,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            }

        reqs.append(num_format(9))   # J열 (매출액 환산수식)
        reqs.append(num_format(17))  # R열 (주문효율 /h)

        # 요청 실행
        sh.batch_update({"requests": reqs})
        print(f"✅ 서식 적용 완료 (J,R 숫자 포맷 포함, 행수 {row_count})")
    except Exception as e:
        print("⚠️ 서식 적용 실패:", e)


# ------------------------------------------------------------
# main
# ------------------------------------------------------------
def main():
    driver = None
    try:
        driver = make_driver()
        login_and_handle_session(driver)
        df_raw = crawl_schedule(driver)

        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)
        df_processed = preprocess_dataframe(df_raw, sh)

        worksheet, new_ws, ins_ws, actual_row_count = upload_sheets(sh, df_processed)

        time.sleep(1)
        new_ws = sh.worksheet(new_ws.title)
        apply_formatting(sh, new_ws, ins_ws, actual_row_count)
        print("🎉 전체 파이프라인 완료")
    finally:
        if driver: driver.quit()


if __name__ == "__main__":
    main()

