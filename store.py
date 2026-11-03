"""
포장 사양 대시보드 - 기록 저장/조회 (History Store)
=====================================================

고객사 · 제품명 · 재원 · 포장 결과를 로컬 CSV(records.csv / box_prices.csv)에
누적 저장하고, 비슷한 사이즈의 과거 기록을 조회할 수 있게 합니다.

⚠️ Streamlit Cloud는 파일 시스템이 임시(ephemeral)입니다. 앱을 재배포하거나
재시작하면 이 CSV 파일들이 초기화됩니다. 영구 보관이 필요하면 대시보드의
'전체 내려받기(CSV)' 버튼으로 주기적으로 백업하고, 필요 시 '업로드로 복원'
기능으로 되살리세요.

동시 사용자 접근
-----------------------------------------------------
여러 사람이 동시에 기록을 저장하거나 단가를 입력할 수 있으므로, 파일 락
(fcntl.flock)으로 읽기→수정→쓰기 구간을 직렬화해 동시 저장 시 한쪽 기록이
유실되지 않게 합니다.

캐싱
-----------------------------------------------------
Streamlit은 위젯 조작마다 스크립트 전체를 재실행하므로, 캐싱이 없으면
load_df()/load_prices_df()가 상호작용마다 디스크에서 CSV를 다시 읽습니다.
`st.cache_data(ttl=30)`로 캐싱하고, 이 모듈로 데이터를 저장할 때마다 캐시를
즉시 비워(clear) 같은 세션 안에서는 방금 저장한 내용이 바로 반영되도록
합니다. streamlit이 없는 환경(예: pytest)에서는 캐싱 없이 그대로 동작합니다.
"""

import contextlib
import logging
import os
import pandas as pd

logger = logging.getLogger(__name__)

try:
    import fcntl
except ImportError:                                   # Windows 등 미지원 환경
    fcntl = None

try:
    import streamlit as st
except ImportError:                                   # pytest 등 streamlit 미설치 환경
    st = None


def _cache_data(**kwargs):
    """st.cache_data 래퍼. streamlit 미설치 시 캐싱 없이 원본 함수 그대로 사용."""
    if st is None:
        def _identity(fn):
            fn.clear = lambda: None
            return fn
        return _identity
    return st.cache_data(**kwargs)


def _clear_cache(fn):
    try:
        fn.clear()
    except Exception:
        pass


@contextlib.contextmanager
def _file_lock(path):
    """path에 대한 읽기-수정-쓰기 구간을 프로세스 간에 직렬화."""
    if fcntl is None:
        yield
        return
    lock_path = path + ".lock"
    with open(lock_path, "a") as f:
        fcntl.flock(f.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(f.fileno(), fcntl.LOCK_UN)


RECORDS_FILE = "records.csv"
PRICES_FILE = "box_prices.csv"

COLUMNS = ["저장일시", "고객사", "제품명", "L", "W", "H", "무게(g)",
           "추천 포장재", "추천 박스", "추천 수량",
           "실제 포장재", "실제 박스", "실제 수량", "일치",
           "트레이재질", "트레이두께", "제작업체", "색상", "비고"]

PRICE_COLS = ["박스명", "확정단가", "갱신일시"]


# ---------------------------------------------------------------------------
# 견적 이력 (records)
# ---------------------------------------------------------------------------
@_cache_data(ttl=30, show_spinner=False)
def load_df():
    """저장된 기록을 DataFrame으로 반환 (없으면 빈 표).

    30초간 캐시됩니다 — 이 모듈로 저장(save_df/append_record)하면
    즉시 캐시를 비우므로 같은 세션에서는 바로 최신값이 보입니다.
    """
    if os.path.exists(RECORDS_FILE):
        try:
            df = pd.read_csv(RECORDS_FILE)
            for c in COLUMNS:                 # 누락 컬럼 보정
                if c not in df.columns:
                    df[c] = ""
            return df[COLUMNS].fillna("")     # 빈칸이 None/NaN으로 보이지 않게
        except Exception:
            logger.exception("%s 파일을 읽지 못했습니다 — 빈 기록으로 표시합니다.", RECORDS_FILE)
    return pd.DataFrame(columns=COLUMNS)


def save_df(df):
    """DataFrame을 CSV로 저장 (best-effort)."""
    with _file_lock(RECORDS_FILE):
        try:
            df[COLUMNS].to_csv(RECORDS_FILE, index=False)
            _clear_cache(load_df)
            return True
        except Exception:
            logger.exception("%s 파일 저장 실패.", RECORDS_FILE)
            return False


def append_record(record: dict):
    """기록 1건 추가 후 저장. 반환: 갱신된 DataFrame."""
    row = {c: record.get(c, "") for c in COLUMNS}
    with _file_lock(RECORDS_FILE):
        df = load_df()
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        try:
            df[COLUMNS].to_csv(RECORDS_FILE, index=False)
        except Exception:
            logger.exception("%s 파일에 기록 추가 실패.", RECORDS_FILE)
    _clear_cache(load_df)
    return df


# ---------------------------------------------------------------------------
# 구매 확정 단가 (prices)
# ---------------------------------------------------------------------------
def load_prices():
    """박스명 → 확정단가(원) dict. 없으면 빈 dict."""
    df = load_prices_df()
    return {str(r["박스명"]): float(r["확정단가"])
            for _, r in df.iterrows() if str(r["박스명"])}


@_cache_data(ttl=30, show_spinner=False)
def load_prices_df():
    """30초간 캐시됩니다 — save_price로 저장하면 즉시 캐시를 비웁니다."""
    if os.path.exists(PRICES_FILE):
        try:
            return pd.read_csv(PRICES_FILE)[PRICE_COLS]
        except Exception:
            logger.exception("%s 파일을 읽지 못했습니다 — 빈 단가 목록으로 표시합니다.", PRICES_FILE)
    return pd.DataFrame(columns=PRICE_COLS)


def save_price(box_name, price, when=""):
    """박스 확정단가 1건 갱신(있으면 덮어씀) 후 저장."""
    with _file_lock(PRICES_FILE):
        df = load_prices_df()
        df = df[df["박스명"].astype(str) != str(box_name)]      # 기존 항목 제거
        row = {"박스명": box_name, "확정단가": price, "갱신일시": when}
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        try:
            df[PRICE_COLS].to_csv(PRICES_FILE, index=False)
            _clear_cache(load_prices_df)
            return True
        except Exception:
            logger.exception("%s 파일 저장 실패.", PRICES_FILE)
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
