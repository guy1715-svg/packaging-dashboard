"""
포장 사양 대시보드 - 계산 로직(Calculation Engine)
====================================================

핵심 공식은 "(박스 내경 / 제품 외경)" 기반의 적재 수량 계산입니다.
구매팀/개발팀이 공식을 그대로 검증할 수 있도록 단순·투명하게 구현했습니다.
"""

import math
from itertools import permutations


def loading_qty_axis_aligned(product, box):
    """
    축 정렬(회전 없음) 기준 박스당 적재 수량.

    공식:
        n_L = floor(박스내경_L / 제품_L)
        n_W = floor(박스내경_W / 제품_W)
        n_H = floor(박스내경_H / 제품_H)
        적재수량 = n_L × n_W × n_H

    product : (L, W, H)  제품 외경 (mm)
    box     : {"inner_l", "inner_w", "inner_h", ...}
    """
    pl, pw, ph = product
    if pl <= 0 or pw <= 0 or ph <= 0:
        return 0, (0, 0, 0)

    n_l = math.floor(box["inner_l"] / pl)
    n_w = math.floor(box["inner_w"] / pw)
    n_h = math.floor(box["inner_h"] / ph)
    n_l, n_w, n_h = max(n_l, 0), max(n_w, 0), max(n_h, 0)
    return n_l * n_w * n_h, (n_l, n_w, n_h)


def loading_qty_best_orientation(product, box):
    """
    제품을 6가지 방향(회전)으로 놓아보고 최대 적재 수량을 반환.
    현장에서는 제품을 눕히거나 세워 적재하므로 실제 최적값에 가깝습니다.

    반환: (최대 적재수량, 최적 방향의 (nL, nW, nH), 최적 제품방향 (L,W,H))
    """
    best_qty = 0
    best_grid = (0, 0, 0)
    best_orient = product
    for orient in set(permutations(product)):
        qty, grid = loading_qty_axis_aligned(orient, box)
        if qty > best_qty:
            best_qty, best_grid, best_orient = qty, grid, orient
    return best_qty, best_grid, best_orient


def volume_efficiency(product, box, qty):
    """적재 부피 효율(%) = (제품 총부피 / 박스 내경 부피) × 100"""
    pl, pw, ph = product
    prod_vol = pl * pw * ph * qty
    box_vol = box["inner_l"] * box["inner_w"] * box["inner_h"]
    if box_vol <= 0:
        return 0.0
    return round(prod_vol / box_vol * 100, 1)


def cost_per_box(box, entity, qty):
    """
    박스당 총원가 계산 (법인별 인건비 가중치 반영).

        포장인건비(가중) = 박스당 포장 인건비 × 인건비 가중치
        박스당 총원가    = 박스 단가 + 포장인건비(가중)

    반환 dict:
        box_cost           : 박스 자재 단가 (법인 통화)
        weighted_labor     : 가중 인건비 (법인 통화)
        total_local        : 박스당 총원가 (법인 통화)
        total_quote        : 견적 통화 환산 총원가
        unit_cost_quote    : 제품 1개당 원가 (견적 통화, 참고용)
    """
    box_cost = box["box_cost"]
    weighted_labor = entity["packing_labor_per_box"] * entity["labor_weight"]
    total_local = box_cost + weighted_labor
    total_quote = total_local * entity["fx_rate"]
    unit_cost_quote = (total_quote / qty) if qty > 0 else 0.0
    return {
        "box_cost": round(box_cost, 2),
        "weighted_labor": round(weighted_labor, 2),
        "total_local": round(total_local, 2),
        "total_quote": round(total_quote, 2),
        "unit_cost_quote": round(unit_cost_quote, 4),
    }


def weight_cap_qty(box, unit_weight_g):
    """
    무게 제한 기준 박스당 최대 적재 수량.

        무게 한도 개수 = ⌊박스 허용중량(kg) × 1000 ÷ 제품 1개 무게(g)⌋

    unit_weight_g <= 0 이면 무게 제한 없음(None) 으로 처리합니다.
    """
    if not unit_weight_g or unit_weight_g <= 0:
        return None
    return max(math.floor(box["max_weight_kg"] * 1000 / unit_weight_g), 0)


def build_quote_rows(product, boxes, entity, use_best_orientation=True,
                     unit_weight_g=0.0, part_name=""):
    """
    선택된 포장 방식/법인의 모든 박스에 대해 견적 행(row) 리스트를 생성.
    UI 표시 및 Excel/PDF 내보내기에 공통으로 사용됩니다.

    unit_weight_g > 0 이면 무게 제한을 반영:
        최종 적재수량 = MIN(부피 기준 개수, 무게 기준 개수)
    part_name 은 각 행의 '품명' 열로 표기됩니다.
    """
    rows = []
    for box in boxes:
        if use_best_orientation:
            vol_qty, grid, orient = loading_qty_best_orientation(product, box)
        else:
            vol_qty, grid = loading_qty_axis_aligned(product, box)
            orient = product

        # 무게 제한 적용 → 최종 적재수량 & 제한 요인 판정
        w_cap = weight_cap_qty(box, unit_weight_g)
        if w_cap is not None and w_cap < vol_qty:
            qty = w_cap
            limit_factor = "무게 제한"
        else:
            qty = vol_qty
            limit_factor = "부피 제한" if unit_weight_g else "-"

        total_weight_kg = round(qty * unit_weight_g / 1000, 2) if unit_weight_g else 0.0
        eff = volume_efficiency(orient, box, qty)
        cost = cost_per_box(box, entity, qty)

        rows.append({
            "품명": part_name if part_name else "-",
            "박스명": box["model"],
            "박스 내경(L×W×H)": f'{box["inner_l"]}×{box["inner_w"]}×{box["inner_h"]}',
            "허용중량(kg)": box["max_weight_kg"],
            "적재 배열(nL×nW×nH)": f"{grid[0]}×{grid[1]}×{grid[2]}",
            "박스당 적재수량": qty,
            "제한 요인": limit_factor,
            "박스 총중량(kg)": total_weight_kg,
            "부피효율(%)": eff,
            "박스 단가": cost["box_cost"],
            "포장 인건비(가중)": cost["weighted_labor"],
            "박스당 총원가": cost["total_local"],
            "제품 1개당 원가(견적통화)": cost["unit_cost_quote"],
            "구매 확정 단가": None,   # ← 구매팀 회신용 (기본 공란)
        })
    return rows
