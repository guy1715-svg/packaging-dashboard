"""
포장 사양 대시보드 - 기준 데이터(Master Data)
=================================================

이 파일은 성우전자(본사)와 성우비나(베트남)의 포장 기준 데이터를 담고 있습니다.
표준 박스 리스트, 법인별 환율/인건비 가중치, 단가 등을 이곳에서 한곳에서 관리합니다.

⚠️ 실제 운영 시 이 딕셔너리 값만 수정하면 대시보드 전체에 반영됩니다.
   (박스 추가/삭제, 단가 변경, 환율 변경 등은 여기서만 하면 됩니다.)
"""

# ---------------------------------------------------------------------------
# 1) 법인(Entity) 기준 정보
#    - 환율(currency, rate): 견적서 통화 환산용
#    - labor_weight: 법인별 인건비 가중치 (본사 1.0 기준, 상대 비율)
#    - 이 값들은 UI 사이드바에서 오버라이드(수정) 가능합니다.
# ---------------------------------------------------------------------------
ENTITIES = {
    "성우전자 (본사)": {
        "code": "SW-KR",
        "country": "대한민국",
        "base_currency": "KRW",
        "quote_currency": "KRW",
        "fx_rate": 1.0,           # KRW -> KRW
        "labor_weight": 1.00,     # 본사 인건비 기준 = 1.0
        "packing_labor_per_box": 1500.0,  # 박스당 포장 인건비 (KRW)
    },
    "성우비나 (베트남)": {
        "code": "SW-VN",
        "country": "베트남",
        "base_currency": "VND",
        "quote_currency": "KRW",
        "fx_rate": 0.055,         # 1 VND ≈ 0.055 KRW (예시, 사이드바에서 수정)
        "labor_weight": 0.35,     # 베트남 인건비 (본사 대비 약 35%, 예시)
        "packing_labor_per_box": 8000.0,  # 박스당 포장 인건비 (VND)
    },
}

# ---------------------------------------------------------------------------
# 2) 포장 방식(Packaging Method)별 표준 박스 카탈로그
#    - inner_* : 박스 "내경" (mm) → 적재 계산의 기준
#    - box_cost: 박스 1개 단가 (해당 법인 통화 기준)
#    - max_weight_kg: 박스 허용 최대 중량
#
#    구조: BOX_CATALOG[포장방식][법인코드] = [박스1, 박스2, ...]
#    새 박스를 추가하려면 아래 리스트에 dict 하나만 추가하면 됩니다.
# ---------------------------------------------------------------------------
BOX_CATALOG = {
    "REEL": {
        "SW-KR": [
            {"model": "REEL-S", "inner_l": 190, "inner_w": 190, "inner_h": 200, "box_cost": 800,  "max_weight_kg": 8},
            {"model": "REEL-M", "inner_l": 340, "inner_w": 340, "inner_h": 360, "box_cost": 1400, "max_weight_kg": 15},
            {"model": "REEL-L", "inner_l": 400, "inner_w": 400, "inner_h": 420, "box_cost": 2000, "max_weight_kg": 20},
        ],
        "SW-VN": [
            {"model": "REEL-S-VN", "inner_l": 190, "inner_w": 190, "inner_h": 200, "box_cost": 12000, "max_weight_kg": 8},
            {"model": "REEL-M-VN", "inner_l": 340, "inner_w": 340, "inner_h": 360, "box_cost": 21000, "max_weight_kg": 15},
            {"model": "REEL-L-VN", "inner_l": 400, "inner_w": 400, "inner_h": 420, "box_cost": 30000, "max_weight_kg": 20},
        ],
    },
    "Bulk": {
        "SW-KR": [
            {"model": "BULK-A", "inner_l": 300, "inner_w": 200, "inner_h": 150, "box_cost": 600,  "max_weight_kg": 10},
            {"model": "BULK-B", "inner_l": 400, "inner_w": 300, "inner_h": 250, "box_cost": 1100, "max_weight_kg": 18},
            {"model": "BULK-C", "inner_l": 550, "inner_w": 380, "inner_h": 320, "box_cost": 1800, "max_weight_kg": 25},
        ],
        "SW-VN": [
            {"model": "BULK-A-VN", "inner_l": 300, "inner_w": 200, "inner_h": 150, "box_cost": 9000,  "max_weight_kg": 10},
            {"model": "BULK-B-VN", "inner_l": 400, "inner_w": 300, "inner_h": 250, "box_cost": 16500, "max_weight_kg": 18},
            {"model": "BULK-C-VN", "inner_l": 550, "inner_w": 380, "inner_h": 320, "box_cost": 27000, "max_weight_kg": 25},
        ],
    },
    "Tray": {
        "SW-KR": [
            {"model": "TRAY-JEDEC-1", "inner_l": 322, "inner_w": 136, "inner_h": 100, "box_cost": 1200, "max_weight_kg": 6},
            {"model": "TRAY-JEDEC-2", "inner_l": 322, "inner_w": 136, "inner_h": 200, "box_cost": 1600, "max_weight_kg": 10},
            {"model": "TRAY-BIG",     "inner_l": 355, "inner_w": 355, "inner_h": 250, "box_cost": 2200, "max_weight_kg": 14},
        ],
        "SW-VN": [
            {"model": "TRAY-JEDEC-1-VN", "inner_l": 322, "inner_w": 136, "inner_h": 100, "box_cost": 18000, "max_weight_kg": 6},
            {"model": "TRAY-JEDEC-2-VN", "inner_l": 322, "inner_w": 136, "inner_h": 200, "box_cost": 24000, "max_weight_kg": 10},
            {"model": "TRAY-BIG-VN",     "inner_l": 355, "inner_w": 355, "inner_h": 250, "box_cost": 33000, "max_weight_kg": 14},
        ],
    },
}

PACKAGING_METHODS = list(BOX_CATALOG.keys())  # ["REEL", "Bulk", "Tray"]

# ---------------------------------------------------------------------------
# 3) 견적 요청서 기본 항목 (구매팀 전달용 표준 양식 헤더)
# ---------------------------------------------------------------------------
QUOTE_FORM_META = {
    "title": "포장자재 견적 요청서 (Packaging Quotation Request)",
    "prepared_by": "개발팀 (R&D / Packaging)",
    "columns": [
        "박스 모델", "박스 내경(L×W×H)", "박스당 적재수량",
        "박스 단가", "포장 인건비(가중)", "박스당 총원가",
        "구매 확정 단가",  # ← 구매팀이 회신하는 필드
    ],
}
