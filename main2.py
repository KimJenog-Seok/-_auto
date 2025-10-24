#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, time, re, json, base64
from pathlib import Path
from datetime import datetime, timedelta, timezone
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===============================================================
# 환경 설정
# ===============================================================
WAIT = 5
ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"
SCHEDULE_URL = "https://live.ecomm-data.com/schedule/hs"

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/19pcFwP2XOVEuHPsr9ITudLDSD1Tzg5RwsL3K6maIJ1U/edit?gid=0#gid=0"
WORKSHEET_NAME = "편성표RAW"

# ===============================================================
# 유틸
# ===============================================================
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
    driver.save_screenshot(str(png))
    with open(html, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"[DEBUG] {png.name}, {html.name} 저장 완료")

# ===============================================================
# 로그인
# ===============================================================
def login_and_handle_session(driver):
    driver.get("https://live.ecomm-data.com")
    print("[STEP] 메인 페이지 진입 완료")
    login_link = WebDriverWait(driver, WAIT).until(EC.element_to_be_clickable((By.LINK_TEXT, "로그인")))
    driver.execute_script("arguments[0].click();", login_link)
    print("[STEP] 로그인 링크 클릭 완료")

    WebDriverWait(driver, 10).until(EC.url_contains("/user/sign_in"))
    print("✅ 로그인 페이지 진입 완료:", driver.current_url)

    email_input = driver.find_element(By.NAME, "email")
    pw_input = driver.find_element(By.NAME, "password")
    email_input.send_keys(ECOMM_ID)
    pw_input.send_keys(ECOMM_PW)

    btn = driver.find_element(By.XPATH, "//button[contains(text(), '로그인')]")
    driver.execute_script("arguments[0].click();", btn)
    print("✅ 로그인 시도!")

    time.sleep(2)
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "ul > li")
        if items:
            items[-1].click()
            time.sleep(1)
            close_btn = driver.find_element(By.XPATH, "//button[text()='종료 후 접속']")
            driver.execute_script("arguments[0].click();", close_btn)
            print("✅ 세션 초과 팝업 처리 완료")
            time.sleep(2)
    except:
        pass

    if "/sign_in" in driver.current_url:
        save_debug(driver, "login_fail")
        raise RuntimeError("로그인 실패")
    print("✅ 로그인 성공!")

# ===============================================================
# 크롤링
# ===============================================================
def crawl_schedule(driver):
    driver.get(SCHEDULE_URL)
    print("✅ 편성표 홈쇼핑 페이지로 이동 완료")

    KST = timezone(timedelta(hours=9))
    yesterday = datetime.now(KST).date() - timedelta(days=1)
    date_text = str(yesterday.day)
    print(f"[STEP] 어제 날짜 선택: {yesterday} → '{date_text}'")

    date_button_xpath = f"//div[text()='{date_text}']"
    date_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, date_button_xpath)))
    driver.execute_script("arguments[0].click();", date_button)
    print("✅ 날짜 클릭 완료")
    time.sleep(3)

    tables = driver.find_elements(By.TAG_NAME, "table")
    all_data = []
    for table in tables:
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
    df = pd.DataFrame(all_data)
    print(f"총 {len(df)}개 편성표 추출 완료")
    return df

# ===============================================================
# 구글 시트 인증
# ===============================================================
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

# ===============================================================
# 한글 숫자 변환
# ===============================================================
def _to_int_kor(s):
    if s is None: return 0
    t = str(s).replace(",", "").replace(" ", "")
    if t == "" or t == "-": return 0
    if re.fullmatch(r"\d+(\.\d+)?", t):
        return int(float(t))
    unit_map = {"억": 100_000_000, "만": 10_000}
    total = 0
    for k, v in unit_map.items():
        if k in t:
            n = t.split(k)[0]
            try: total += float(n) * v
            except: pass
            t = t.split(k)[1]
    try: total += float(t)
    except: pass
    return int(total)

# ===============================================================
# 전처리
# ===============================================================
def preprocess_dataframe(df_raw, sh):
    print("🧮 데이터 전처리 시작")
    df = df_raw.copy()

    # 방송날짜 / 시간 분리
    split = df["방송시간"].str.split("\n", n=1, expand=True)
    df["방송날짜"] = pd.to_datetime(split[0].str.strip(), format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")
    df["방송시작시간"] = split[1].str.strip() if len(split.columns) == 2 else ""

    # 회사명/홈쇼핑구분
    df["회사명"] = df["방송정보"].apply(lambda x: re.sub(r".*(CJ|GS|현대|롯데|NS|공영|홈앤|쇼핑엔티|신세계|SK|KT알파).*", r"\1", str(x)))
    df["홈쇼핑구분"] = df["회사명"].apply(lambda x: "Live" if "홈쇼핑" in x else "TC")

    # 매출액 환산
    df["매출액 환산수식"] = df["매출액"].apply(_to_int_kor)
    df["환산가치"] = 0.0

    # ✅ 분리송출고려환산가치 보정
    if "분리송출고려환산가치" not in df.columns:
        grp_counts = df.groupby(["회사명", "방송시작시간"])["방송시작시간"].transform("size")
        df["분리송출고려환산가치"] = df["환산가치"] / grp_counts.clip(lower=1)

    # ✅ 매출액 환산수식 보정
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

    df["주문효율 /h"] = df.apply(lambda r: safe_eff(r["매출액 환산수식"], r["분리송출고려환산가치"]), axis=1)
    df["주문효율 /h"] = pd.to_numeric(df["주문효율 /h"], errors="coerce").fillna(0).round().astype(int)
    print("✅ 데이터 전처리 완료")
    return df

# ===============================================================
# 시트 서식
# ===============================================================
def apply_formatting(sh, new_ws, ins_ws, row_count):
    try:
        reqs = []

        def num_format(col):
            return {
                "repeatCell": {
                    "range": {"sheetId": new_ws.id, "startRowIndex": 1, "endRowIndex": row_count, "startColumnIndex": col, "endColumnIndex": col + 1},
                    "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
                    "fields": "userEnteredFormat.numberFormat"
                }
            }

        # 숫자 포맷 J,R열
        reqs.append(num_format(9))
        reqs.append(num_format(17))

        # 기본 테두리
        reqs.append({
            "updateBorders": {
                "range": {"sheetId": new_ws.id, "startRowIndex": 0, "endRowIndex": row_count, "startColumnIndex": 0, "endColumnIndex": 18},
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"},
            }
        })

        sh.batch_update({"requests": reqs})
        print("✅ 서식 적용 완료 (J,R열 #,##0 포함)")
    except Exception as e:
        print("⚠️ 서식 적용 실패:", e)

# ===============================================================
# 메인
# ===============================================================
def main():
    key_path = Path("C:/key/composed-apogee-442305-k5-b134efa6db1c.json")
    if key_path.exists() and not os.environ.get("KEY1"):
        with open(key_path, "rb") as f:
            os.environ["KEY1"] = base64.b64encode(f.read()).decode("utf-8")

    driver = None
    try:
        driver = make_driver()
        login_and_handle_session(driver)
        df_raw = crawl_schedule(driver)

        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)

        df_processed = preprocess_dataframe(df_raw, sh)
        data_to_upload = [df_processed.columns.tolist()] + df_processed.values.tolist()

        # RAW 시트 업로드
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=len(df_processed.columns))
        ws.clear()
        ws.update("A1", data_to_upload, value_input_option="USER_ENTERED")
        print("✅ RAW 시트 업로드 완료")

        # 날짜 시트 생성
        title = datetime.now().strftime("%m/%d")
        new_ws = sh.add_worksheet(title=title, rows=len(data_to_upload)+1, cols=len(df_processed.columns))
        new_ws.update("A1", data_to_upload, value_input_option="USER_ENTERED")
        print(f"✅ 날짜 시트 '{title}' 생성 완료")

        # INS_전일
        try:
            ins_ws = sh.worksheet("INS_전일")
        except gspread.exceptions.WorksheetNotFound:
            ins_ws = sh.add_worksheet(title="INS_전일", rows=3, cols=3)

        time.sleep(1)
        new_ws = sh.worksheet(title)
        apply_formatting(sh, new_ws, ins_ws, len(data_to_upload)+1)
        print("🎉 전체 완료!")

    except Exception as e:
        import traceback
        print("❌ 오류:", e)
        print(traceback.format_exc())
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
