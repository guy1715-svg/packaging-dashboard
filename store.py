"""
포장 사양 대시보드 - 기록 저장/조회 (History Store)
=====================================================

고객사 · 제품명 · 재원 · 포장 결과를 CSV 파일에 누적 저장하고,
비슷한 사이즈의 과거 기록을 조회할 수 있게 합니다.

⚠️ Streamlit Cloud는 파일 시스템이 임시(ephemeral)입니다.
   재배포/재시작 시 records.csv가 초기화되므로, 영구 보관은
   대시보드의 '전체 내려받기(CSV)'로 파일을 보관하거나
   추후 구글 시트 연동으로 대체하세요.
"""

import os
import pandas as pd

RECORDS_FILE = "records.csv"

COLUMNS = ["저장일시", "고객사", "제품명", "L", "W", "H", "무게(g)",
           "추천 포장재", "추천 박스", "추천 수량",
           "실제 포장재", "실제 박스", "실제 수량", "일치",
           "트레이재질", "트레이두께", "제작업체", "색상", "비고"]


def load_df():
    """저장된 기록을 DataFrame으로 반환 (없으면 빈 표)."""
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
    """DataFrame을 CSV로 저장 (best-effort)."""
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


PRICES_FILE = "box_prices.csv"
PRICE_COLS = ["박스명", "확정단가", "갱신일시"]


def load_prices():
    """박스명 → 확정단가(원) dict. 파일 없으면 빈 dict."""
    if os.path.exists(PRICES_FILE):
        try:
            df = pd.read_csv(PRICES_FILE)
            return {str(r["박스명"]): float(r["확정단가"])
                    for _, r in df.iterrows() if str(r["박스명"])}
        except Exception:
            pass
    return {}


def load_prices_df():
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
