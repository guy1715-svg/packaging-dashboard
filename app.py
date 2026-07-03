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
import altair as alt
import pandas as pd
import streamlit as st

from data import ENTITIES, BOX_CATALOG, PACKAGING_CONFIG
from calculations import build_packaging_rows, tray_cell_count, fit_zipper_bag
from exporters import to_excel_bytes, to_pdf_bytes, default_header_info

st.set_page_config(
    page_title="포장 사양 · 견적 대시보드",
    page_icon="📦",
    layout="wide",
)

# ---------------------------------------------------------------------------
# 디자인 시스템 (검증된 다크 팔레트)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
:root{
  --bg:#0d1117; --surface:#161b22; --surface2:#1c2432; --border:#2a3140;
  --text:#e6edf3; --muted:#8b98a5; --accent:#3987e5; --accent2:#199e70; --warn:#c98500;
}
.block-container{padding-top:1.6rem;padding-bottom:2rem;max-width:1320px;}
h1,h2,h3{letter-spacing:-.01em;}
h1{font-size:1.85rem !important;}
/* 상단 컨텍스트 칩 바 */
.ctxbar{display:flex;gap:8px;flex-wrap:wrap;margin:.2rem 0 1rem;}
.chip{background:var(--surface);border:1px solid var(--border);border-radius:999px;
  padding:5px 13px;font-size:.82rem;color:var(--text);white-space:nowrap;}
.chip b{color:var(--muted);font-weight:600;margin-right:6px;font-size:.76rem;}
/* 섹션 헤더 */
.sec{display:flex;align-items:center;gap:9px;margin:.4rem 0 .5rem;font-weight:700;
  font-size:1.06rem;color:var(--text);}
.sec .dot{width:9px;height:9px;border-radius:50%;background:var(--accent);
  box-shadow:0 0 0 4px rgba(57,135,229,.16);}
/* KPI 카드 행 */
.kpi-row{display:flex;gap:14px;align-items:stretch;flex-wrap:wrap;margin:.3rem 0 .2rem;}
.kpi-card{flex:1;min-width:150px;background:#161b22;border:1px solid #2a3140;
  border-radius:14px;padding:15px 18px;position:relative;overflow:hidden;
  animation:kpiIn .45s cubic-bezier(.2,.7,.3,1) both;
  transition:transform .16s ease,border-color .16s ease,box-shadow .16s ease;}
.kpi-card:hover{transform:translateY(-3px);border-color:var(--accent);
  box-shadow:0 6px 22px rgba(0,0,0,.35);}
.kpi-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent);}
.kpi-card.total{background:linear-gradient(180deg,rgba(25,158,112,.16),var(--surface));
  border-color:var(--accent2);}
.kpi-card.total::before{background:var(--accent2);}
.kpi-label{font-size:.78rem;color:var(--muted);margin-bottom:7px;}
.kpi-value{font-size:2.05rem;font-weight:750;color:var(--text);line-height:1.05;
  font-variant-numeric:tabular-nums;}
.kpi-card.total .kpi-value{color:#40d6a0;}
.kpi-unit{font-size:.92rem;font-weight:600;color:var(--muted);margin-left:5px;}
.kpi-sub{font-size:.73rem;color:var(--muted);margin-top:7px;}
.kpi-op{display:flex;align-items:center;font-size:1.5rem;color:var(--muted);font-weight:700;}
@keyframes kpiIn{from{opacity:0;transform:translateY(9px);}to{opacity:1;transform:none;}}
/* 사이드바 스텝 배지 */
.step{display:flex;align-items:center;gap:9px;margin:.1rem 0 .3rem;font-weight:700;
  font-size:.98rem;color:var(--text);}
.step .num{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;
  border-radius:50%;background:var(--accent);color:#06121f;font-size:.8rem;font-weight:800;}
section[data-testid="stSidebar"]{border-right:1px solid var(--border);}
/* 탭 */
.stTabs [data-baseweb="tab-list"]{gap:4px;}
.stTabs [data-baseweb="tab"]{padding:8px 16px;border-radius:9px 9px 0 0;}
</style>
""", unsafe_allow_html=True)


def kpi_row(cards):
    """KPI 카드 한 줄 렌더링. cards: [{label,value,unit,sub,variant,op}]"""
    html = ['<div class="kpi-row">']
    for c in cards:
        if c.get("op"):
            html.append(f'<div class="kpi-op">{c["op"]}</div>')
        html.append(
            f'<div class="kpi-card {c.get("variant","")}">'
            f'<div class="kpi-label">{c["label"]}</div>'
            f'<div class="kpi-value">{c["value"]}'
            f'<span class="kpi-unit">{c.get("unit","")}</span></div>'
            f'<div class="kpi-sub">{c.get("sub","")}</div></div>'
        )
    html.append('</div>')
    st.markdown("".join(html), unsafe_allow_html=True)


def section(title):
    st.markdown(f'<div class="sec"><span class="dot"></span>{title}</div>',
                unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 헤더
# ---------------------------------------------------------------------------
st.title("📦 성우전자 · 성우비나 포장 견적 대시보드")
st.caption("포장 사양 입력 → 적재 효율 자동 계산 → 구매팀 전달용 견적 요청서 생성")

# ---------------------------------------------------------------------------
# 사이드바 - 입력
# ---------------------------------------------------------------------------
def sidebar_step(n, title):
    st.markdown(f'<div class="step"><span class="num">{n}</span>{title}</div>',
                unsafe_allow_html=True)


with st.sidebar:
    sidebar_step(1, "기본 입력")

    entity_name = st.selectbox("대상 법인", list(ENTITIES.keys()))
    entity_code = ENTITIES[entity_name]["code"]
    cfg = PACKAGING_CONFIG[entity_code]

    part_name = st.text_input("품명 (제품명/품번)", value="",
                              placeholder="예: SW-CONN-0250")

    st.markdown("**제품 외경 (mm)**")
    c1, c2, c3 = st.columns(3)
    pl = c1.number_input("L", min_value=0.1, value=50.0, step=1.0)
    pw = c2.number_input("W", min_value=0.1, value=30.0, step=1.0)
    ph = c3.number_input("H", min_value=0.1, value=15.0, step=1.0)
    product = (pl, pw, ph)

    unit_weight_g = st.number_input(
        "제품 1개 무게 (g) · 0=무게 무시", min_value=0.0, value=0.0, step=1.0,
        help="값을 넣으면 박스 허용중량을 넘지 않도록 적재수량을 제한합니다.")

    st.markdown("")
    sidebar_step(2, "포장 방식")

    inner_mode = st.selectbox("포장재 선택 (안쪽)", cfg["inner_options"],
                              help="제품을 담는 1차 포장재. 없음=벌크로 박스에 직접 적재")
    outer_group = st.selectbox("박스 종류 선택 (바깥)", cfg["outer_groups"],
                               help="제품/트레이를 담는 최종 박스 종류")

    is_tray = "트레이" in inner_mode
    is_bag = "지퍼백" in inner_mode

    # --- 트레이 설정 (범용트레이 선택 시) ---
    tray_gap = tray_pitch_x = tray_pitch_y = 0.0
    tray_l = tray_w = tray_thickness = 0.0
    if is_tray:
        with st.expander("🔧 트레이 설정", expanded=True):
            t1, t2 = st.columns(2)
            tray_l = t1.number_input("트레이 가로", min_value=0.0, value=315.0, step=5.0)
            tray_w = t2.number_input("트레이 세로", min_value=0.0, value=410.0, step=5.0)
            tray_thickness = st.number_input(
                "트레이 두께/높이 (mm)", min_value=0.0, value=15.0, step=1.0,
                help="트레이 1장 높이 → 박스에 몇 단 쌓이는지 계산에 사용")
            p1, p2 = st.columns(2)
            tray_pitch_x = p1.number_input("제품 피치 X", min_value=0.0, value=0.0,
                                           step=0.5, help="칸 중심간 거리(도면값). 0=자동")
            tray_pitch_y = p2.number_input("제품 피치 Y", min_value=0.0, value=0.0,
                                           step=0.5, help="0=자동(제품크기+여유로 계산)")
            tray_gap = st.number_input(
                "칸 사이 여유 간격 (mm)", min_value=0.0, value=0.0, step=0.5,
                help="피치를 비웠을 때만 사용. 제품 사이에 두는 간격.")

    # --- 고급 옵션 (기본 접힘) ---
    with st.expander("⚙️ 고급 옵션"):
        use_best = st.toggle("최적 방향(회전) 적재", value=True,
                             help="제품을 6방향으로 돌려 최대 적재수량을 계산합니다.")
        wall_margin = st.number_input(
            "박스 벽두께 여유 (mm)", min_value=0.0, value=0.0, step=1.0,
            help="박스 규격이 외경일 때, 벽두께만큼 빼고 계산합니다. (0=규격 그대로)")

    st.markdown("")
    sidebar_step(3, "단가 · 환율")

    base = ENTITIES[entity_name]
    entity = copy.deepcopy(base)
    with st.expander("💱 단가 · 환율 (법인별)"):
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
# 메인 - 계산 (제품 → 포장재 → 박스)
# ---------------------------------------------------------------------------
outer_boxes = BOX_CATALOG[entity_code][outer_group]

# 트레이 칸수 (범용트레이 선택 시)
tray_cells, tray_grid = 0, (0, 0)
if is_tray and tray_l > 0 and tray_w > 0:
    tray_cells, tray_grid = tray_cell_count(
        product, {"inner_l": tray_l, "inner_w": tray_w},
        gap=tray_gap, pitch_x=tray_pitch_x, pitch_y=tray_pitch_y)

# 적합 지퍼백 (지퍼백 선택 시)
bag_name = ""
if is_bag and cfg["bag_group"]:
    _bag = fit_zipper_bag(product, BOX_CATALOG[entity_code][cfg["bag_group"]])
    bag_name = _bag["박스명"] if _bag else "적합 규격 없음"

rows = build_packaging_rows(
    product, outer_boxes, entity, inner_mode=inner_mode, outer_group=outer_group,
    unit_weight_g=unit_weight_g, part_name=part_name, wall_margin=wall_margin,
    use_best=use_best, tray_cells=tray_cells, tray_l=tray_l, tray_w=tray_w,
    tray_thickness=tray_thickness, bag_name=bag_name)

best_row = max(rows, key=lambda r: r["박스당 총 제품"]) if rows else None

# 컨텍스트 칩 바
_wchip = f'<span class="chip"><b>무게</b> {unit_weight_g:g} g</span>' if unit_weight_g else ""
_bchip = f'<span class="chip"><b>지퍼백</b> {bag_name}</span>' if (is_bag and bag_name) else ""
st.markdown(
    '<div class="ctxbar">'
    f'<span class="chip"><b>법인</b> {entity_name}</span>'
    f'<span class="chip"><b>포장재</b> {inner_mode}</span>'
    f'<span class="chip"><b>박스</b> {outer_group}</span>'
    f'<span class="chip"><b>제품(mm)</b> {pl:g}×{pw:g}×{ph:g}</span>'
    f'{_wchip}{_bchip}'
    '</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# 탭 구성
# ---------------------------------------------------------------------------
tab_calc, tab_form, tab_data = st.tabs(
    ["📊 적재 효율 계산", "📄 구매팀 견적 양식", "🗄️ 기준 데이터(박스 리스트)"])

# --- 탭1: 적재 효율 ---
with tab_calc:
    section(f"제품 → {inner_mode} → {outer_group}")

    if not rows or best_row is None:
        st.error("표시할 박스가 없습니다.")
    else:
        # 히어로 KPI 카드
        cards = []
        if is_tray:
            cards.append({"label": "트레이당 제품", "value": f"{tray_cells:,}",
                          "unit": "개", "sub": f"칸 배열 {tray_grid[0]}×{tray_grid[1]}"})
        cards.append({"label": f"박스당 총 제품 · {best_row['박스명']}",
                      "value": f"{best_row['박스당 총 제품']:,}", "unit": "개",
                      "sub": f"규격 {best_row['규격(Size)']}",
                      "op": "→" if is_tray else "", "variant": "total"})
        if is_bag:
            cards.append({"label": "적합 지퍼백", "value": bag_name or "-", "unit": "",
                          "sub": "제품이 들어가는 최소 규격"})
        else:
            cards.append({"label": "검토 박스", "value": f"{len(rows)}", "unit": "종",
                          "sub": outer_group})
        kpi_row(cards)
        if best_row["박스당 총 제품"] == 0:
            st.warning("이 조합으로는 적재되지 않습니다. 제품 사이즈·트레이 설정·박스 종류를 확인하세요.")
        st.markdown("")

        df = pd.DataFrame(rows)
        section(f"{outer_group} · 박스별 총 제품 수")
        show_cols = ["박스명", "규격(Size)", "포장재", "적재 방식",
                     "박스당 총 제품", "제한 요인", "비고"]
        if unit_weight_g:
            show_cols.insert(5, "박스 총중량(kg)")
        st.dataframe(
            df[show_cols], use_container_width=True, hide_index=True,
            column_config={
                "박스당 총 제품": st.column_config.NumberColumn(format="%d 개"),
            },
        )

        # 막대 차트: 최다 적재 박스 그린 강조
        mx = df["박스당 총 제품"].max()
        chart = (
            alt.Chart(df)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("박스명:N", sort=None, axis=alt.Axis(labelAngle=0, title=None)),
                y=alt.Y("박스당 총 제품:Q", title="박스당 총 제품",
                        axis=alt.Axis(grid=True)),
                color=alt.condition(alt.datum["박스당 총 제품"] == mx,
                                    alt.value("#199e70"), alt.value("#3987e5")),
                tooltip=["박스명", "규격(Size)", "적재 방식", "박스당 총 제품"],
            )
            .properties(height=300)
            .configure_view(strokeWidth=0)
            .configure_axis(labelColor="#8b98a5", titleColor="#8b98a5",
                            gridColor="#2a3140", domainColor="#2a3140")
        )
        st.altair_chart(chart, use_container_width=True)

        cta, info_ = st.columns([1, 1])
        cta.info("📄 견적서로 만들려면 상단 **'구매팀 견적 양식'** 탭으로 이동하세요.")
        with info_.expander("ℹ️ 계산 방식 보기"):
            st.markdown(
                "- **없음(벌크)**: 제품을 박스에 직접 3D 적재 "
                "⌊박스 ÷ 제품⌋ (최적 방향 6방향 중 최대)\n"
                "- **지퍼백**: 제품을 봉투에 넣어 박스에 3D 적재 (적합 봉투 자동 선정)\n"
                "- **범용트레이**: 트레이 칸수 × (박스당 트레이 장수) = 박스당 총 제품\n"
                "- **무게 입력 시**: 박스 허용중량을 넘지 않도록 제한 → '제한 요인'에 표시")

# --- 탭2: 견적 양식 + 다운로드 ---
with tab_form:
    section("구매팀 전달용 표준 견적 요청서")
    header_info = default_header_info(entity_name, entity, outer_group, product,
                                      part_name=part_name, unit_weight_g=unit_weight_g,
                                      inner_mode=inner_mode, bag_name=bag_name)

    # 메타 정보를 칩으로
    chips = "".join(f'<span class="chip"><b>{k}</b> {v}</span>'
                    for k, v in header_info.items())
    st.markdown(f'<div class="ctxbar">{chips}</div>', unsafe_allow_html=True)

    st.markdown("##### 견적 항목  ·  구매팀이 **‘구매 확정 단가’** 열에 입력해 회신")
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

        section("다운로드하여 구매팀에 전달")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Excel 견적 요청서 (.xlsx)",
                data=to_excel_bytes(header_info, export_rows),
                file_name=f'견적요청서_{entity_code}_{outer_group}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ PDF 견적 요청서 (.pdf)",
                data=to_pdf_bytes(header_info, export_rows),
                file_name=f'견적요청서_{entity_code}_{outer_group}.pdf',
                mime="application/pdf",
                use_container_width=True,
            )

# --- 탭3: 기준 데이터 열람 ---
with tab_data:
    section("표준 박스 카탈로그 (기준 데이터)")
    st.caption("실제 값 수정은 `data.py`의 BOX_CATALOG / ENTITIES 딕셔너리에서 관리합니다.")
    st.markdown(f"**{outer_group} · {entity_name}**")
    st.dataframe(pd.DataFrame(outer_boxes), use_container_width=True, hide_index=True)

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
