"""
성우전자 / 성우비나 포장 사양 & 견적 요청 대시보드
=====================================================

실행:
    pip install -r requirements.txt
    streamlit run app.py

구조:
    사이드바 → 입력(제품 사이즈, 포장 방식, 법인) + 법인별 환율/인건비 오버라이드
    메인     → 적재 효율 계산 결과 + 구매팀 전달용 표준 양식 + 다운로드
"""

import copy
import pandas as pd
import streamlit as st

from data import ENTITIES, BOX_CATALOG, categories_for
from calculations import build_quote_rows
from exporters import to_excel_bytes, to_pdf_bytes, default_header_info

st.set_page_config(
    page_title="포장 사양 · 견적 대시보드",
    page_icon="📦",
    layout="wide",
)

# ---------------------------------------------------------------------------
# 헤더
# ---------------------------------------------------------------------------
st.title("📦 성우전자 · 성우비나 포장 사양 견적 대시보드")
st.caption("포장 사양 입력 → 적재 효율 자동 계산 → 구매팀 전달용 견적 요청서 생성")

# ---------------------------------------------------------------------------
# 사이드바 - 입력
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("① 기본 입력")

    entity_name = st.selectbox("대상 법인", list(ENTITIES.keys()))
    entity_code = ENTITIES[entity_name]["code"]
    category = st.radio("포장재 분류", categories_for(entity_code))

    catalog_boxes = BOX_CATALOG[entity_code][category]
    is_tray = bool(catalog_boxes) and catalog_boxes[0].get("pack_type") == "tray"

    part_name = st.text_input("품명 (제품명/품번)", value="",
                              placeholder="예: SW-CONN-0250")

    st.markdown("**제품 외경 (mm)**")
    c1, c2, c3 = st.columns(3)
    with c1:
        pl = st.number_input("L", min_value=0.1, value=50.0, step=1.0)
    with c2:
        pw = st.number_input("W", min_value=0.1, value=30.0, step=1.0)
    with c3:
        ph = st.number_input("H", min_value=0.1, value=15.0, step=1.0)
    product = (pl, pw, ph)

    unit_weight_g = st.number_input(
        "제품 1개 무게 (g)", min_value=0.0, value=0.0, step=1.0,
        help="0으로 두면 무게 제한 없이 부피만으로 계산합니다. "
             "값을 넣으면 박스 허용중량을 초과하지 않도록 적재수량을 제한합니다.")

    # --- 트레이 전용: 트레이 사이즈 입력 → 칸수 자동 계산 ---
    tray_gap = 0.0
    custom_tray = None
    if is_tray:
        st.markdown("**트레이 사이즈 (제품 삽입 칸수 계산용)**")
        t1, t2 = st.columns(2)
        with t1:
            tray_l = st.number_input("트레이 가로 (mm)", min_value=0.0,
                                     value=315.0, step=5.0)
        with t2:
            tray_w = st.number_input("트레이 세로 (mm)", min_value=0.0,
                                     value=410.0, step=5.0)
        tray_gap = st.number_input(
            "칸 사이 여유 간격 (mm)", min_value=0.0, value=0.0, step=0.5,
            help="제품과 제품(칸) 사이에 두는 간격. 0이면 딱 붙여 계산합니다.")
        add_custom = st.checkbox("입력한 트레이 사이즈로 계산 추가", value=True,
                                 help="체크하면 위에서 입력한 트레이 규격이 표 맨 위에 "
                                      "'사용자 지정 트레이'로 추가됩니다.")
        if add_custom and tray_l > 0 and tray_w > 0:
            custom_tray = {
                "박스명": "사용자 지정 트레이", "size": f"{tray_l:g}*{tray_w:g}",
                "inner_l": tray_l, "inner_w": tray_w, "inner_h": 0,
                "pack_type": "tray", "재질": "", "비고": "입력값 기준",
                "box_cost": 0, "max_weight_kg": 15,
            }

    use_best = st.toggle("최적 방향(회전) 적재 사용", value=True,
                         help="제품을 6방향으로 놓아보고 최대 적재수량을 계산합니다. "
                              "끄면 회전 없는 축정렬 기준(내경/외경)만 사용합니다.")

    wall_margin = st.number_input(
        "박스 벽두께 여유 (mm)", min_value=0.0, value=0.0, step=1.0,
        help="박스 규격은 보통 외경입니다. 골판지 벽두께만큼 내경이 줄어드니, "
             "여유값을 넣으면 각 변에서 그만큼 빼고 계산합니다. (0=엑셀 규격 그대로)")

    st.divider()
    st.header("② 법인별 가중치 (수정 가능)")
    # data.py 값을 기본값으로 하되, 사이드바에서 오버라이드
    base = ENTITIES[entity_name]
    entity = copy.deepcopy(base)
    entity["fx_rate"] = st.number_input(
        f'환율 (1 {base["base_currency"]} → {base["quote_currency"]})',
        min_value=0.0, value=float(base["fx_rate"]), step=0.001, format="%.4f")
    entity["labor_weight"] = st.number_input(
        "인건비 가중치 (본사 1.0 기준)",
        min_value=0.0, value=float(base["labor_weight"]), step=0.05)
    entity["packing_labor_per_box"] = st.number_input(
        f'박스당 포장 인건비 ({base["base_currency"]})',
        min_value=0.0, value=float(base["packing_labor_per_box"]), step=100.0)

# ---------------------------------------------------------------------------
# 메인 - 계산
# ---------------------------------------------------------------------------
boxes = list(catalog_boxes)
if custom_tray is not None:
    boxes = [custom_tray] + boxes      # 사용자 지정 트레이를 맨 위에
rows = build_quote_rows(product, boxes, entity, use_best_orientation=use_best,
                        unit_weight_g=unit_weight_g, part_name=part_name,
                        wall_margin=wall_margin, tray_gap=tray_gap)

# 요약 지표
top = st.columns(4)
top[0].metric("대상 법인", entity_name.split(" ")[0])
top[1].metric("포장재 분류", category)
top[2].metric("제품 사이즈", f"{pl:g}×{pw:g}×{ph:g}")
best_row = max(rows, key=lambda r: r["박스당 적재수량"]) if rows else None
top[3].metric("최대 적재수량 / 박스",
              f'{best_row["박스당 적재수량"]:,} ({best_row["박스명"]})' if best_row else "0")

st.divider()

# ---------------------------------------------------------------------------
# 탭 구성
# ---------------------------------------------------------------------------
tab_calc, tab_form, tab_data = st.tabs(
    ["📊 적재 효율 계산", "📄 구매팀 견적 양식", "🗄️ 기준 데이터(박스 리스트)"])

# --- 탭1: 적재 효율 ---
with tab_calc:
    st.subheader(f"{category} · {entity_name} 적재 효율")
    if not rows:
        st.error("적재 가능한 박스가 없습니다. 제품 사이즈가 박스 내경보다 큰지 확인하세요.")
    else:
        df = pd.DataFrame(rows).drop(columns=["구매 확정 단가"])
        st.dataframe(
            df,
            use_container_width=True, hide_index=True,
            column_config={
                "박스당 적재수량": st.column_config.NumberColumn(format="%d 개"),
                "부피효율(%)": st.column_config.ProgressColumn(
                    format="%.1f%%", min_value=0, max_value=100),
            },
        )
        st.bar_chart(df.set_index("박스명")["박스당 적재수량"])
        st.info("적재수량 계산:  ▸박스 = ⌊박스내경 ÷ 제품외경⌋ 을 L·W·H에 적용 후 곱셈  "
                "(최적 방향 ON 시 6방향 중 최대).  ▸트레이 = 칸수.  ▸지퍼백 = 1개입.")
        if unit_weight_g:
            st.info("무게 반영:  최종 적재수량 = MIN(부피 기준 개수, 무게 기준 개수).  "
                    "'제한 요인' 열에서 무게·부피 중 무엇이 한계였는지 확인할 수 있습니다.")
        else:
            st.caption("💡 사이드바에 '제품 1개 무게(g)'를 입력하면 박스 허용중량까지 반영해 더 정확히 계산합니다.")

# --- 탭2: 견적 양식 + 다운로드 ---
with tab_form:
    st.subheader("구매팀 전달용 표준 견적 요청서")
    header_info = default_header_info(entity_name, entity, category, product,
                                      part_name=part_name, unit_weight_g=unit_weight_g)

    meta_cols = st.columns(3)
    items = list(header_info.items())
    for i, (k, v) in enumerate(items):
        meta_cols[i % 3].write(f"**{k}**: {v}")

    st.markdown("##### 견적 항목 (구매팀이 '구매 확정 단가'를 입력해 회신)")
    if not rows:
        st.warning("표시할 항목이 없습니다.")
    else:
        editable = pd.DataFrame(rows)
        edited = st.data_editor(
            editable,
            use_container_width=True, hide_index=True, num_rows="fixed",
            column_config={
                "구매 확정 단가": st.column_config.NumberColumn(
                    "구매 확정 단가 ✍️",
                    help="구매팀이 최종 단가를 입력하는 필드입니다. (데이터 축적용)",
                    min_value=0.0),
            },
            disabled=[c for c in editable.columns if c != "구매 확정 단가"],
        )
        export_rows = edited.to_dict(orient="records")

        st.markdown("##### 다운로드")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Excel 견적 요청서 (.xlsx)",
                data=to_excel_bytes(header_info, export_rows),
                file_name=f'견적요청서_{entity_code}_{category}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ PDF 견적 요청서 (.pdf)",
                data=to_pdf_bytes(header_info, export_rows),
                file_name=f'견적요청서_{entity_code}_{category}.pdf',
                mime="application/pdf",
                use_container_width=True,
            )

# --- 탭3: 기준 데이터 열람 ---
with tab_data:
    st.subheader("표준 박스 카탈로그 (기준 데이터)")
    st.caption("실제 값 수정은 `data.py`의 BOX_CATALOG / ENTITIES 딕셔너리에서 관리합니다.")
    st.markdown(f"**{category} · {entity_name}**")
    st.dataframe(pd.DataFrame(boxes), use_container_width=True, hide_index=True)

    st.markdown("**법인 기준 정보 (현재 적용값 · 사이드바 오버라이드 반영)**")
    st.json({
        "법인": entity_name,
        "코드": entity["code"],
        "환율": entity["fx_rate"],
        "인건비 가중치": entity["labor_weight"],
        "박스당 포장 인건비": entity["packing_labor_per_box"],
        "견적 통화": entity["quote_currency"],
    })

st.divider()
st.caption("© 성우전자/성우비나 개발팀 · 포장 사양 견적 대시보드")
