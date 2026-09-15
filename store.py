"""
포장 사양 대시보드 - 기록 저장/조회 (History Store)
=====================================================

고객사 · 제품명 · 재원 · 포장 결과를 누적 저장하고,
비슷한 사이즈의 과거 기록을 조회할 수 있게 합니다.

저장소 우선순위
-----------------------------------------------------
  1) Google Sheets — Streamlit secrets에 아래 두 값이 설정되어 있으면 사용:
       st.secrets["gcp_service_account"]  (서비스 계정 JSON 키, dict)
       st.secrets["gsheet_id"]            (스프레드시트 ID)
     Streamlit Cloud가 재배포/재시작되어도 데이터가 사라지지 않습니다.
  2) 로컬 CSV(records.csv / box_prices.csv) — 위 secrets가 없을 때 폴백.
     Streamlit Cloud는 파일 시스템이 임시(ephemeral)이므로 재배포 시
     초기화됩니다. 로컬 개발/오프라인 실행용으로만 사용하세요.

Google Sheets 연동 실패(권한 오류, 네트워크 오류 등) 시에도 앱이 죽지
않도록 로컬 CSV로 안전하게 폴백합니다.
"""

import os
import pandas as pd

RECORDS_FILE = "records.csv"
PRICES_FILE = "box_prices.csv"
RECORDS_SHEET = "records"
PRICES_SHEET = "box_prices"

COLUMNS = ["저장일시", "고객사", "제품명", "L", "W", "H", "무게(g)",
           "추천 포장재", "추천 박스", "추천 수량",
           "실제 포장재", "실제 박스", "실제 수량", "일치",
           "트레이재질", "트레이두께", "제작업체", "색상", "비고"]

PRICE_COLS = ["박스명", "확정단가", "갱신일시"]


# ---------------------------------------------------------------------------
# Google Sheets 연결
# ---------------------------------------------------------------------------
def _gsheet():
    """설정된 Google Sheets 스프레드시트 객체. 미설정/실패 시 None."""
    try:
        import streamlit as st
        if "gcp_service_account" not in st.secrets or "gsheet_id" not in st.secrets:
            return None
        return _open_gsheet(dict(st.secrets["gcp_service_account"]), st.secrets["gsheet_id"])
    except Exception:
        return None


def _open_gsheet(sa_info, sheet_id):
    import streamlit as st

    @st.cache_resource(show_spinner=False)
    def _client(sheet_id):
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        gc = gspread.authorize(creds)
        return gc.open_by_key(sheet_id)

    return _client(sheet_id)


def _get_or_create_ws(sh, title, columns):
    try:
        ws = sh.worksheet(title)
    except Exception:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(len(columns), 10))
        ws.update([columns])
        return ws
    header = ws.row_values(1)
    if header != columns:
        ws.update([columns])
    return ws


def _ws_to_df(ws, columns):
    rows = ws.get_all_records()
    if not rows:
        return pd.DataFrame(columns=columns)
    df = pd.DataFrame(rows)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    return df[columns].fillna("")


def _df_to_ws(ws, df, columns):
    df = df.reindex(columns=columns).fillna("")
    values = [columns] + df.astype(str).values.tolist()
    ws.clear()
    ws.update(values)


# ---------------------------------------------------------------------------
# 견적 이력 (records)
# ---------------------------------------------------------------------------
def load_df():
    """저장된 기록을 DataFrame으로 반환 (없으면 빈 표)."""
    sh = _gsheet()
    if sh is not None:
        try:
            ws = _get_or_create_ws(sh, RECORDS_SHEET, COLUMNS)
            return _ws_to_df(ws, COLUMNS)
        except Exception:
            pass
    if os.path.exists(RECORDS_FILE):
        try:
            df = pd.read_csv(RECORDS_FILE)
            for c in COLUMNS:                 # 누락 컬럼 보정
                if c not in df.columns:
                    df[c] = ""
            return df[COLUMNS].fillna("")     # 빈칸이 None/NaN으로 보이지 않게
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNS)


def save_df(df):
    """DataFrame을 저장소(Google Sheets 우선, 없으면 CSV)에 저장 (best-effort)."""
    sh = _gsheet()
    if sh is not None:
        try:
            ws = _get_or_create_ws(sh, RECORDS_SHEET, COLUMNS)
            _df_to_ws(ws, df, COLUMNS)
            return True
        except Exception:
            pass
    try:
        df[COLUMNS].to_csv(RECORDS_FILE, index=False)
        return True
    except Exception:
        return False


def append_record(record: dict):
    """기록 1건 추가 후 저장. 반환: 갱신된 DataFrame."""
    df = load_df()
    row = {c: record.get(c, "") for c in COLUMNS}
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    save_df(df)
    return df


# ---------------------------------------------------------------------------
# 구매 확정 단가 (prices)
# ---------------------------------------------------------------------------
def load_prices():
    """박스명 → 확정단가(원) dict. 없으면 빈 dict."""
    df = load_prices_df()
    return {str(r["박스명"]): float(r["확정단가"])
            for _, r in df.iterrows() if str(r["박스명"])}


def load_prices_df():
    sh = _gsheet()
    if sh is not None:
        try:
            ws = _get_or_create_ws(sh, PRICES_SHEET, PRICE_COLS)
            return _ws_to_df(ws, PRICE_COLS)
        except Exception:
            pass
    if os.path.exists(PRICES_FILE):
        try:
            return pd.read_csv(PRICES_FILE)[PRICE_COLS]
        except Exception:
            pass
    return pd.DataFrame(columns=PRICE_COLS)


def save_price(box_name, price, when=""):
    """박스 확정단가 1건 갱신(있으면 덮어씀) 후 저장."""
    df = load_prices_df()
    df = df[df["박스명"].astype(str) != str(box_name)]      # 기존 항목 제거
    row = {"박스명": box_name, "확정단가": price, "갱신일시": when}
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

    sh = _gsheet()
    if sh is not None:
        try:
            ws = _get_or_create_ws(sh, PRICES_SHEET, PRICE_COLS)
            _df_to_ws(ws, df, PRICE_COLS)
            return True
        except Exception:
            pass
    try:
        df[PRICE_COLS].to_csv(PRICES_FILE, index=False)
        return True
    except Exception:
        return False


def similar_records(df, dims, tol_ratio=0.15, tol_min=5.0):
    """
    현재 제품 재원(dims=(L,W,H))과 비슷한 사이즈의 과거 기록만 필터링.
    세 변을 정렬해 각 변 차이가 max(변×tol_ratio, tol_min) 이내면 '비슷함'.
    """
    if df is None or df.empty:
        return df
    a = sorted(float(x) for x in dims)

    def is_similar(r):
        try:
            b = sorted([float(r["L"]), float(r["W"]), float(r["H"])])
        except (TypeError, ValueError):
            return False
        return all(abs(b[i] - a[i]) <= max(a[i] * tol_ratio, tol_min)
                   for i in range(3))

    return df[df.apply(is_similar, axis=1)]
