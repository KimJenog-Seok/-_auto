#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, json, base64
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials


# ===================== Google Sheets 인증 =====================
def gs_client_from_env():
    GSVC_JSON_B64 = os.environ.get("KEY1", "")
    if not GSVC_JSON_B64:
        raise RuntimeError("환경변수 KEY1(Base64 인코딩된 서비스계정 JSON) 없음")

    svc_info = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(svc_info, scopes=scope)
    return gspread.authorize(creds)


# ===================== 어제 날짜 기반 시트 이름 찾기 =====================
def make_yesterday_title_kst():
    KST = timezone(timedelta(hours=9))
    today = datetime.now(KST).date()
    yday = today - timedelta(days=1)
    # 💡 수정: 메인 코드와 형식을 맞춤 (25/12/18)
    return yday.strftime("%y/%m/%d")


def find_latest_backup_sheet(sh, base_title):
    """
    기본 시트(11/18) 또는 11/18-1, 11/18-2 중 가장 마지막 번호를 반환
    """

    candidates = []
    for ws in sh.worksheets():
        title = ws.title
        if title == base_title:
            candidates.append((0, title))
        else:
            # 11/18-2 같은 형식 검사
            if title.startswith(base_title + "-"):
                try:
                    num = int(title.split("-")[-1])
                    candidates.append((num, title))
                except:
                    pass 

    if not candidates:
        raise RuntimeError(f"백업 시트를 찾을 수 없음: {base_title}")

    # 번호가 가장 큰 것이 최신
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]  # (번호, 제목) → 제목만


# ===================== 서식 적용 =====================
def apply_decimal_formatting(sh, ws):
    """
    대상 워크시트의 M열(12번 index=12), Q열(16번 index=16)에
    소수점 둘째 자리 숫자 서식 적용
    """

    sheet_id = ws.id

    requests = []

    # M열 = 12번째 (A=0 기준 → index=12)
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": 12,
                "endColumnIndex": 13
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    # Q열 = 16번째 (index=16)
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": 16,
                "endColumnIndex": 17
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    sh.batch_update({"requests": requests})
    print(f"✨ 서식 적용 완료: 시트 '{ws.title}' (M,Q 열 → #,##0.00)")


# ===================== 메인 =====================
def main():
    SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/19pcFwP2XOVEuHPsr9ITudLDSD1Tzg5RwsL3K6maIJ1U/edit"

    # 1) 접속
    gc = gs_client_from_env()
    sh = gc.open_by_url(SPREADSHEET_URL)
    print("🔗 구글시트 연결 완료")

    # 2) 어제 날짜 제목 구하기
    base_title = make_yesterday_title_kst()
    print("📌 어제 날짜 시트 기본 이름:", base_title)

    # 3) 최신 시트 탐색
    latest_title = find_latest_backup_sheet(sh, base_title)
    print("📌 대상 백업 시트:", latest_title)

    ws = sh.worksheet(latest_title)

    # 4) 서식 적용
    apply_decimal_formatting(sh, ws)

    print("🎉 format_only.py 전체 완료!")


if __name__ == "__main__":
    main()
