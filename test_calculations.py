"""
calculations.py 회귀 테스트
============================
적재 공식·무게 제한·트레이·지퍼백·파렛트/컨테이너 계산을
알려진 값으로 고정한다. 계산 로직을 수정하면 이 테스트가 먼저 깨져
잘못된 변경을 즉시 알려준다.

실행:  pytest -q          (프로젝트 루트에서)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from calculations import (                     # noqa: E402
    loading_qty_axis_aligned, loading_qty_best_orientation,
    tray_cell_count, trays_per_box, fit_zipper_bag, bag_layer_capacity,
    boxes_per_pallet, boxes_per_container, cbm, build_packaging_rows,
)

BOX = {"inner_l": 500, "inner_w": 300, "inner_h": 370}
BAGS = [
    {"박스명": "s", "inner_l": 60, "inner_w": 90},
    {"박스명": "m", "inner_l": 120, "inner_w": 170},
    {"박스명": "l", "inner_l": 240, "inner_w": 340},
]
PALLET = {"l": 1100, "w": 1100, "pallet_h": 150, "max_stack": 1500}
CONTAINER = {"l": 5898, "w": 2352, "h": 2393}


# --- 축정렬 적재 --------------------------------------------------------------
def test_axis_aligned_basic():
    qty, grid = loading_qty_axis_aligned((50, 30, 15), BOX)
    assert qty == 2400
    assert grid == (10, 10, 24)


def test_axis_aligned_zero_and_too_big():
    assert loading_qty_axis_aligned((0, 30, 15), BOX) == (0, (0, 0, 0))
    assert loading_qty_axis_aligned((600, 30, 15), BOX)[0] == 0   # 한 변이 박스보다 큼


# --- 최적 방향 적재 (회전) ----------------------------------------------------
def test_best_orientation_beats_axis():
    axis, _ = loading_qty_axis_aligned((40, 40, 100), BOX)
    best, grid, orient = loading_qty_best_orientation((40, 40, 100), BOX)
    assert axis == 252
    assert best == 324                     # 눕히면 더 많이 들어감
    assert best > axis


# --- 부피(CBM) ----------------------------------------------------------------
def test_cbm():
    assert cbm(1000, 1000, 1000) == 1.0
    assert abs(cbm(500, 300, 370) - 0.0555) < 1e-9


# --- 지퍼백 -------------------------------------------------------------------
def test_bag_layer_capacity():
    assert bag_layer_capacity((50, 30, 15), {"inner_l": 240, "inner_w": 340}) == 176
    assert bag_layer_capacity((50, 30, 15), {"inner_l": 60, "inner_w": 90}) == 12


def test_fit_zipper_bag():
    assert fit_zipper_bag((50, 30, 15), BAGS)["박스명"] == "s"    # 가장 작은 것
    assert fit_zipper_bag((400, 400, 5), BAGS) is None            # 들어가는 봉투 없음


# --- 트레이 -------------------------------------------------------------------
def test_trays_per_box():
    total, base, layers = trays_per_box(200, 140, 30, BOX)
    assert (total, base, layers) == (48, 4, 12)


def test_tray_cell_count_with_pitch():
    cnt, grid = tray_cell_count((10, 10, 5), {"inner_l": 315, "inner_w": 410},
                                pitch_x=50, pitch_y=50)
    assert cnt == 48
    assert grid == (6, 8)


# --- 물류 ---------------------------------------------------------------------
def test_boxes_per_pallet():
    total, base, layers = boxes_per_pallet(400, 300, 200, PALLET)
    assert (total, base, layers) == (36, 6, 6)


def test_boxes_per_container_at_least_axis():
    axis = 14 * 7 * 11                       # 5898/400 · 2352/300 · 2393/200
    qty, _ = boxes_per_container(400, 300, 200, CONTAINER)
    assert qty >= axis


# --- 통합: build_packaging_rows ----------------------------------------------
BOXES = [{"박스명": "B", "size": "500*300*370",
          "inner_l": 500, "inner_w": 300, "inner_h": 370, "비고": ""}]


def test_rows_bulk_no_weight():
    rows = build_packaging_rows((50, 30, 15), BOXES, inner_mode="없음(벌크)")
    assert rows[0]["박스당 총 제품"] == 2400
    assert rows[0]["제한 요인"] == "-"


def test_rows_weight_cap():
    rows = build_packaging_rows((50, 30, 15), BOXES, inner_mode="없음(벌크)",
                                unit_weight_g=60, weight_limit_kg=10)
    assert rows[0]["박스당 총 제품"] == 166        # floor(10000/60)
    assert "무게 제한" in rows[0]["제한 요인"]


def test_rows_safety_pct():
    rows = build_packaging_rows((50, 30, 15), BOXES, inner_mode="없음(벌크)",
                                safety_pct=10)
    assert rows[0]["박스당 총 제품"] == 2160        # int(2400 * 0.9)


def test_rows_tray():
    rows = build_packaging_rows((50, 30, 15), BOXES, inner_mode="범용트레이",
                                tray_cells=48, tray_l=200, tray_w=140,
                                tray_thickness=30)
    assert rows[0]["박스당 총 제품"] == 48 * 48     # 칸수 × 박스당 트레이수


def test_rows_bag_self_consistent():
    m = loading_qty_best_orientation((60, 90, 45), BOX)[0]
    rows = build_packaging_rows((50, 30, 15), BOXES, inner_mode="지퍼백",
                                bag_count=12, bag_l=60, bag_w=90, bag_h=45)
    assert rows[0]["박스당 총 제품"] == 12 * m
