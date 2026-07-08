"""
포장 사양 대시보드 - 계산 로직(Calculation Engine)
====================================================

핵심 흐름: 제품 → (포장재: 없음/지퍼백/트레이) → 박스
공식을 그대로 검증할 수 있도록 단순·투명하게 구현했습니다.
"""

import math
from itertools import permutations


def loading_qty_axis_aligned(product, box):
    """축 정렬(회전 없음) 기준 박스당 적재 수량 = ⌊내경/외경⌋ 을 L·W·H 곱."""
    pl, pw, ph = product
    if pl <= 0 or pw <= 0 or ph <= 0:
        return 0, (0, 0, 0)
    n_l = max(math.floor(box["inner_l"] / pl), 0)
    n_w = max(math.floor(box["inner_w"] / pw), 0)
    n_h = max(math.floor(box["inner_h"] / ph), 0)
    return n_l * n_w * n_h, (n_l, n_w, n_h)


def loading_qty_best_orientation(product, box):
    """제품을 6방향으로 돌려 최대 적재 수량. 반환: (수량, (nL,nW,nH), 제품방향)."""
    best_qty, best_grid, best_orient = 0, (0, 0, 0), product
    for orient in set(permutations(product)):
        qty, grid = loading_qty_axis_aligned(orient, box)
        if qty > best_qty:
            best_qty, best_grid, best_orient = qty, grid, orient
    return best_qty, best_grid, best_orient


def tray_cell_count(product, tray, gap=0.0, pitch_x=0.0, pitch_y=0.0):
    """
    트레이(가로 inner_l × 세로 inner_w)에 제품이 몇 칸(개) 배치되는지 계산.
    피치가 주어지면 도면 피치로, 없으면 제품 크기+여유(gap)로 자동 계산.
    반환: (칸수, (가로칸, 세로칸))
    """
    L, W = tray["inner_l"], tray["inner_w"]
    if pitch_x > 0 and pitch_y > 0:
        def pgrid(px, py):
            return max(int(L // px), 0), max(int(W // py), 0)
        g1, g2 = pgrid(pitch_x, pitch_y), pgrid(pitch_y, pitch_x)
        na, nb = g1 if g1[0] * g1[1] >= g2[0] * g2[1] else g2
        return na * nb, (na, nb)

    a, b = sorted(product)[:2]

    def grid(pa, pb):
        if pa <= 0 or pb <= 0:
            return 0, 0
        return (max(int((L + gap) // (pa + gap)), 0),
                max(int((W + gap) // (pb + gap)), 0))

    g1, g2 = grid(a, b), grid(b, a)
    na, nb = g1 if g1[0] * g1[1] >= g2[0] * g2[1] else g2
    return na * nb, (na, nb)


def trays_per_box(tray_l, tray_w, tray_thickness, box, wall_margin=0.0):
    """
    박스 1개에 트레이가 몇 장 적재되는지.
        바닥면 트레이 = ⌊박스가로/트레이가로⌋ × ⌊박스세로/트레이세로⌋ (양방향 중 최대)
        적층 단수     = ⌊박스높이/트레이두께⌋
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


def _apply_margin(box, margin):
    if not margin:
        return box
    b = dict(box)
    b["inner_l"] = max(box["inner_l"] - margin, 0)
    b["inner_w"] = max(box["inner_w"] - margin, 0)
    b["inner_h"] = max(box["inner_h"] - margin, 0)
    return b


def fit_zipper_bag(product, bags):
    """제품이 들어가는 가장 작은(면적 기준) 지퍼백을 반환. 없으면 None."""
    a, b = sorted(product)[:2]
    cand = []
    for bg in bags:
        bl, bw = sorted([bg["inner_l"], bg["inner_w"]])
        if a <= bl and b <= bw:
            cand.append((bl * bw, bg))
    return min(cand, key=lambda x: x[0])[1] if cand else None


def bag_layer_capacity(product, bag):
    """지퍼백 1봉지에 제품이 '한 겹'으로 몇 개 깔리는지 (양방향 중 최대)."""
    a, b = sorted(product)[:2]
    L, W = bag["inner_l"], bag["inner_w"]

    def f(x, y):
        if x <= 0 or y <= 0:
            return 0
        return int(L // x) * int(W // y)

    return max(f(a, b), f(b, a))


def build_packaging_rows(product, outer_boxes, *, inner_mode, outer_group="",
                         unit_weight_g=0.0, part_name="", wall_margin=0.0,
                         use_best=True, tray_cells=0, tray_grid=(0, 0),
                         tray_l=0.0, tray_w=0.0, tray_thickness=0.0,
                         bag_name="", bag_count=1, bag_l=0.0, bag_w=0.0, bag_h=0.0,
                         weight_limit_kg=10.0, safety_pct=0.0):
    """
    제품 → (포장재) → 박스 흐름으로 '박스 1개당 총 제품 수'를 박스별로 계산.

      트레이 : 총제품 = 트레이 칸수 × (박스당 트레이 장수)
      지퍼백 : 총제품 = 입수(bag_count) × (박스당 지퍼백 봉수)
      없음   : 총제품 = 제품을 박스에 직접 3D 적재

    safety_pct : 적재 여유율(%). 이론 수량에서 그만큼 감산.
    weight_limit_kg : 박스당 총중량 한도(기본 10kg). 초과 시 수량 제한.
    각 행에 그리기용 필드(_cols/_rows/_layers/_unit)를 포함합니다.
    """
    is_tray = "트레이" in inner_mode
    is_bag = "지퍼백" in inner_mode
    rows = []
    for box in outer_boxes:
        b = _apply_margin(box, wall_margin)
        if is_tray:
            m_per_box, base_cnt, layers = trays_per_box(
                tray_l, tray_w, tray_thickness, box, wall_margin=wall_margin)
            base_total = tray_cells * m_per_box
            method = f"트레이 {tray_cells}칸 × {m_per_box}장 (바닥{base_cnt}×{layers}단)"
            dcols, drows, dlayers, dunit = tray_grid[0], tray_grid[1], m_per_box, "칸"
        elif is_bag:
            m_per_box, grid, _ = loading_qty_best_orientation((bag_l, bag_w, bag_h), b)
            base_total = bag_count * m_per_box
            method = f"{bag_count}개입 지퍼백 × {m_per_box}봉 (3D {grid[0]}×{grid[1]}×{grid[2]})"
            dcols, drows, dlayers, dunit = grid[0], grid[1], grid[2], "봉"
        else:
            if use_best:
                qty, grid, _ = loading_qty_best_orientation(product, b)
            else:
                qty, grid = loading_qty_axis_aligned(product, b)
            base_total = qty
            method = f"벌크 3D {grid[0]}×{grid[1]}×{grid[2]}"
            dcols, drows, dlayers, dunit = grid[0], grid[1], grid[2], "개"

        # 적재 여유율(안전계수) 적용
        if safety_pct > 0:
            base_total = int(base_total * (1 - safety_pct / 100.0))

        # 무게 제한: 박스당 총중량이 weight_limit_kg 을 넘지 않도록
        if unit_weight_g and unit_weight_g > 0:
            w_cap = int(weight_limit_kg * 1000 // unit_weight_g)
            if w_cap < base_total:
                total, limit = w_cap, f"무게 제한 ({weight_limit_kg:g}kg)"
            else:
                total, limit = base_total, "무게 OK"
        else:
            total, limit = base_total, ("여유율 반영" if safety_pct > 0 else "-")

        total_w = round(total * unit_weight_g / 1000, 2) if unit_weight_g else 0.0
        inner_label = inner_mode + (f" · {bag_name}" if (is_bag and bag_name) else "")

        rows.append({
            "품명": part_name if part_name else "-",
            "박스종류": outer_group,
            "박스명": box.get("박스명", ""),
            "규격(Size)": box.get("size", ""),
            "포장재": inner_label,
            "적재 방식": method,
            "박스당 총 제품": total,
            "제한 요인": limit,
            "박스 총중량(kg)": total_w,
            "비고": box.get("비고", ""),
            "구매 확정 단가": None,
            "_cols": dcols, "_rows": drows, "_layers": dlayers, "_unit": dunit,
        })
    return rows
