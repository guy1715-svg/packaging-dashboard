"""
store.py 회귀 테스트 (로컬 CSV 저장 경로)
============================================
로컬 CSV 저장/조회 로직이 올바르게 동작하는지 검증한다.

각 테스트는 임시 디렉터리로 이동해 실행해, 실제 프로젝트의
records.csv/box_prices.csv를 건드리지 않는다.

실행:  pytest -q          (프로젝트 루트에서)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import pytest

import store


@pytest.fixture(autouse=True)
def _isolated_cwd(tmp_path, monkeypatch):
    """각 테스트를 빈 임시 디렉터리에서 실행 (CSV 파일 격리)."""
    monkeypatch.chdir(tmp_path)
    yield


def test_load_df_empty_when_no_file():
    df = store.load_df()
    assert df.empty
    assert list(df.columns) == store.COLUMNS


def test_append_record_then_load():
    store.append_record({"고객사": "A사", "제품명": "부품1", "L": 10, "W": 20, "H": 30})
    df = store.load_df()
    assert len(df) == 1
    assert df.iloc[0]["고객사"] == "A사"
    assert df.iloc[0]["제품명"] == "부품1"


def test_append_record_missing_fields_default_to_blank():
    store.append_record({"고객사": "B사"})
    df = store.load_df()
    assert df.iloc[0]["고객사"] == "B사"
    assert df.iloc[0]["제품명"] == ""    # 지정하지 않은 컬럼은 빈 문자열


def test_multiple_appends_accumulate():
    for i in range(3):
        store.append_record({"고객사": f"고객{i}"})
    df = store.load_df()
    assert len(df) == 3
    assert list(df["고객사"]) == ["고객0", "고객1", "고객2"]


def test_save_df_overwrites():
    store.append_record({"고객사": "A사"})
    store.save_df(pd.DataFrame(columns=store.COLUMNS))    # 전체 삭제와 동일한 패턴
    assert store.load_df().empty


def test_load_prices_empty_when_no_file():
    assert store.load_prices() == {}
    assert store.load_prices_df().empty


def test_save_price_then_load():
    store.save_price("55-2", 1200.0, "2026-01-01 10:00")
    prices = store.load_prices()
    assert prices["55-2"] == 1200.0


def test_save_price_overwrites_existing_box():
    store.save_price("55-2", 1000.0, "2026-01-01")
    store.save_price("55-2", 1500.0, "2026-01-02")
    df = store.load_prices_df()
    assert len(df[df["박스명"] == "55-2"]) == 1            # 중복 행 없이 갱신됨
    assert store.load_prices()["55-2"] == 1500.0


def test_save_price_does_not_clobber_other_boxes():
    store.save_price("55-2", 1000.0, "2026-01-01")
    store.save_price("57-2", 2000.0, "2026-01-01")
    prices = store.load_prices()
    assert prices["55-2"] == 1000.0
    assert prices["57-2"] == 2000.0


def test_similar_records_filters_by_tolerance():
    store.append_record({"고객사": "A", "L": 100, "W": 100, "H": 100})
    store.append_record({"고객사": "B", "L": 500, "W": 500, "H": 500})
    hist = store.load_df()
    sim = store.similar_records(hist, (100, 100, 100))
    assert len(sim) == 1
    assert sim.iloc[0]["고객사"] == "A"


def test_similar_records_empty_history():
    assert store.similar_records(pd.DataFrame(columns=store.COLUMNS), (1, 1, 1)).empty
