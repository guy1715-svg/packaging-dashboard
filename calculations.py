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


def product_fits_2d(product, box):
    """제품의 가장 작은 두 변이 포장재 바닥(inner_l × inner_w) 안에 들어가는지."""
    dims = sorted(product)          # 제품을 눕혀 가장 납작하게
    base = sorted([box["inner_l"], box["inner_w"]])
    return dims[0] <= base[0] and dims[1] <= base[1]


def tray_cell_count(product, tray, gap=0.0):
    """
    트레이(가로 inner_l × 세로 inner_w)에 제품이 몇 칸(개) 배치되는지 계산.

        칸수 = ⌊(트레이가로+간격) ÷ (제품+간격)⌋ × ⌊(트레이세로+간격) ÷ (제품+간격)⌋

    - 제품은 가장 납작하게 눕힌 두 변(제일 작은 두 치수)을 바닥면으로 사용
    - gap: 칸(제품) 사이 여유 간격(mm)
    반환: (칸수, (가로칸, 세로칸))
    """
    a, b = sorted(product)[:2]      # 제품 바닥면 두 변
    L, W = tray["inner_l"], tray["inner_w"]

    def grid(pa, pb):
        if pa <= 0 or pb <= 0:
            return 0, 0
        na = max(int((L + gap) // (pa + gap)), 0)
        nb = max(int((W + gap) // (pb + gap)), 0)
        return na, nb

    # 제품을 트레이 위에서 두 방향으로 돌려보고 더 많이 들어가는 쪽 채택
    g1 = grid(a, b)
    g2 = grid(b, a)
    if g1[0] * g1[1] >= g2[0] * g2[1]:
        na, nb = g1
    else:
        na, nb = g2
    return na * nb, (na, nb)


def trays_per_box(tray_l, tray_w, tray_thickness, box, wall_margin=0.0):
    """
    박스 1개에 트레이가 몇 장 적재되는지 계산.

        바닥면 트레이 수 = ⌊박스가로 ÷ 트레이가로⌋ × ⌊박스세로 ÷ 트레이세로⌋ (양방향 중 최대)
        적층 단수        = ⌊박스높이 ÷ 트레이두께⌋
        박스당 트레이     = 바닥면 트레이 수 × 적층 단수

    반환: (총 트레이수, 바닥면 트레이수, 적층 단수)
    """
    L = max(box["inner_l"] - wall_margin, 0)
    W = max(box["inner_w"] - wall_margin, 0)
    H = max(box["inner_h"] - wall_margin, 0)

    def per_layer(a, b):
        if a <= 0 or b <= 0:
            return 0
        return int(L // a) * int(W // b)

    base = max(per_layer(tray_l, tray_w), per_layer(tray_w, tray_l))
    layers = int(H // tray_thickness) if tray_thickness > 0 else 0
    return base * layers, base, layers


def build_quote_rows(product, boxes, entity, use_best_orientation=True,
                     unit_weight_g=0.0, part_name="", wall_margin=0.0, tray_gap=0.0):
    """
    선택된 법인/분류의 모든 포장재에 대해 견적 행(row) 리스트를 생성.
    포장재 유형(pack_type)에 따라 적재수량 계산 방식이 다릅니다.

      box  : 3D 적재. 최종수량 = MIN(부피 기준, 무게 기준). wall_margin(벽두께 여유)
             만큼 내경을 줄여 계산.
      tray : 칸수(cells)를 적재수량으로 사용.
      bag  : 1개입(제품이 봉투 규격에 들어가면 1, 아니면 0).

    part_name 은 각 행의 '품명' 열로 표기됩니다.
    """
    rows = []
    for box in boxes:
        pt = box.get("pack_type", "box")
        name = box.get("박스명", box.get("model", ""))

        eff = 0.0
        if pt == "tray":
            cells, tgrid = tray_cell_count(product, box, gap=tray_gap)
            base_qty = cells
            arrange = f"{tgrid[0]}×{tgrid[1]} = {cells}칸"
            unit_type = "트레이"
            base_limit = "트레이 칸수" if cells else "제품이 트레이보다 큼"
        elif pt == "bag":
            base_qty = 1 if product_fits_2d(product, box) else 0
            arrange = "1개입" if base_qty else "사이즈 초과"
            unit_type = "지퍼백"
            base_limit = "1개입" if base_qty else "사이즈 초과"
        else:  # box (3D 적재)
            eff_box = dict(box)
            if wall_margin:
                eff_box["inner_l"] = max(box["inner_l"] - wall_margin, 0)
                eff_box["inner_w"] = max(box["inner_w"] - wall_margin, 0)
                eff_box["inner_h"] = max(box["inner_h"] - wall_margin, 0)
            if use_best_orientation:
                base_qty, grid, orient = loading_qty_best_orientation(product, eff_box)
            else:
                base_qty, grid = loading_qty_axis_aligned(product, eff_box)
                orient = product
            arrange = f"{grid[0]}×{grid[1]}×{grid[2]}"
            unit_type = "박스"
            base_limit = "부피 제한" if unit_weight_g else "-"
            eff = volume_efficiency(orient, eff_box, base_qty)

        # 무게 제한 적용 (모든 유형 공통) → 최종 적재수량 & 제한 요인
        w_cap = weight_cap_qty(box, unit_weight_g)
        if w_cap is not None and w_cap < base_qty:
            qty = w_cap
            limit_factor = "무게 제한"
            if pt == "box":
                eff = volume_efficiency(orient, eff_box, qty)
        else:
            qty = base_qty
            limit_factor = base_limit

        total_weight_kg = round(qty * unit_weight_g / 1000, 2) if unit_weight_g else 0.0
        cost = cost_per_box(box, entity, qty)

        rows.append({
            "품명": part_name if part_name else "-",
            "박스명": name,
            "규격(Size)": box.get("size", f'{box["inner_l"]}*{box["inner_w"]}*{box["inner_h"]}'),
            "유형": unit_type,
            "재질/겹": box.get("재질", ""),
            "적재 배열": arrange,
            "박스당 적재수량": qty,
            "제한 요인": limit_factor,
            "박스 총중량(kg)": total_weight_kg,
            "허용중량(kg)": box.get("max_weight_kg", 0),
            "부피효율(%)": eff,
            "박스 단가": cost["box_cost"],
            "포장 인건비(가중)": cost["weighted_labor"],
            "박스당 총원가": cost["total_local"],
            "제품 1개당 원가(견적통화)": cost["unit_cost_quote"],
            "비고": box.get("비고", ""),
            "구매 확정 단가": None,   # ← 구매팀 회신용 (기본 공란)
        })
    return rows
