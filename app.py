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
import streamlit.components.v1 as components

from data import (DEFAULT_ENTITY, BOX_CATALOG, INNER_OPTIONS, OUTER_GROUPS,
                  BAG_GROUP, TRAY_GROUP)
from calculations import (build_packaging_rows, tray_cell_count,
                          fit_zipper_bag, bag_layer_capacity)
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


def packing_image(cols, rows_, layers, unit, box_name, total):
    """박스 1개의 윗면 적재 배치를 그린 SVG 이미지 (한 층 기준 + 단수 표기)."""
    cols, rows_, layers = int(cols), int(rows_), max(int(layers), 1)
    if cols <= 0 or rows_ <= 0:
        st.info("이 조합은 적재되지 않아 이미지를 표시할 수 없습니다.")
        return
    cap = 22
    dc, dr = min(cols, cap), min(rows_, cap)
    trunc = cols > cap or rows_ > cap
    cell, gap, pad, top = 20, 4, 22, 54
    w = pad * 2 + dc * (cell + gap) - gap
    h = top + pad + dr * (cell + gap) - gap + pad
    w = max(w, 320)
    rects = []
    for r in range(dr):
        for c in range(dc):
            x = pad + c * (cell + gap)
            y = top + r * (cell + gap)
            rects.append(
                f'<rect x="{x}" y="{y}" width="{cell}" height="{cell}" rx="3" '
                f'fill="#3987e5" fill-opacity="0.9" stroke="#7ec8f3" stroke-width="1"/>')
    caption = f"윗면 {cols}×{rows_} = {cols*rows_} {unit}/층  ·  {layers}층"
    if trunc:
        caption += "  (그림은 일부만 표시)"
    svg = f'''
<div style="background:#161b22;border:1px solid #2a3140;border-radius:14px;
     padding:14px 16px;overflow-x:auto;font-family:system-ui,sans-serif;">
  <div style="color:#e6edf3;font-weight:700;font-size:15px;margin-bottom:2px;">
     📦 {box_name} · 적재 배치도</div>
  <div style="color:#8b98a5;font-size:12.5px;margin-bottom:10px;">
     {caption}  →  <b style="color:#40d6a0;">박스당 {total:,} 개</b></div>
  <svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">
    <rect x="6" y="{top-8}" width="{w-12}" height="{h-top-4}" rx="10"
          fill="none" stroke="#2a3140" stroke-width="2"/>
    {''.join(rects)}
  </svg>
</div>'''
    components.html(svg, height=h + 70, scrolling=True)


# ---------------------------------------------------------------------------
# 헤더
# ---------------------------------------------------------------------------
st.title("📦 성우 포장 견적 대시보드")
st.caption("포장 사양 입력 → 적재 효율 자동 계산 → 구매팀 전달용 견적 요청서 생성")

# ---------------------------------------------------------------------------
# 사이드바 - 입력
# ---------------------------------------------------------------------------
def sidebar_step(n, title):
    st.markdown(f'<div class="step"><span class="num">{n}</span>{title}</div>',
                unsafe_allow_html=True)


with st.sidebar:
    sidebar_step(1, "기본 입력")

    part_name = st.text_input("품명 (제품명/품번)", value="",
                              placeholder="예: SW-CONN-0250")

    st.markdown("**제품 외경 (mm)**")
    c1, c2, c3 = st.columns(3)
    pl = c1.number_input("L", min_value=0.1, value=50.0, step=1.0)
    pw = c2.number_input("W", min_value=0.1, value=30.0, step=1.0)
    ph = c3.number_input("H", min_value=0.1, value=15.0, step=1.0)
    product = (pl, pw, ph)

    w1, w2 = st.columns(2)
    unit_weight_g = w1.number_input(
        "제품 1개 무게 (g)", min_value=0.0, value=0.0, step=1.0,
        help="0=무게 무시. 값을 넣으면 박스 총중량 한도를 넘지 않도록 제한합니다.")
    weight_limit_kg = w2.number_input(
        "박스 총중량 한도 (kg)", min_value=0.1, value=10.0, step=1.0,
        help="박스 1개당 허용 총중량. 기본 10kg.")

    st.markdown("")
    sidebar_step(2, "포장 방식")

    inner_mode = st.selectbox("포장재 선택 (안쪽)", INNER_OPTIONS,
                              help="제품을 담는 1차 포장재. 없음=벌크로 박스에 직접 적재")
    outer_group = st.selectbox("박스 종류 선택 (바깥)", OUTER_GROUPS,
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

    # --- 지퍼백 설정 (지퍼백 선택 시) : 규격·입수 모두 자동 추천 ---
    bag_count = 1
    bag_l = bag_w = bag_h = 0.0
    bag_name = ""
    if is_bag:
        _bags = BOX_CATALOG[BAG_GROUP]
        with st.expander("👜 지퍼백 설정", expanded=True):
            # 규격 추천: 제품이 들어가는 가장 작은 봉투
            _fit = fit_zipper_bag(product, _bags) or (_bags[-1] if _bags else None)
            _names = [f'{bg["박스명"]} · {bg["size"]}' for bg in _bags]
            _idx = _bags.index(_fit) if _fit in _bags else 0
            _sel = st.selectbox("지퍼백 규격 (자동 추천 · 변경 가능)", _names, index=_idx,
                                help="제품이 들어가는 최소 규격을 자동 추천합니다.")
            _chosen = _bags[_names.index(_sel)]
            _per_layer = bag_layer_capacity(product, _chosen) or 1
            st.caption(f"✅ 추천 입수: **한 봉지 {_per_layer}개** "
                       f"(봉투 {_chosen['size']}에 한 겹 가득)")
            _override = st.checkbox("입수 직접 지정", value=False,
                                    help="추천값 대신 원하는 입수를 직접 넣습니다.")
            bag_count = st.number_input(
                "지퍼백 1개당 제품 수 (입수)", min_value=1, value=int(_per_layer),
                step=1, disabled=not _override) if _override else int(_per_layer)
            _layers_in_bag = -(-bag_count // _per_layer)      # 올림
            bag_l, bag_w = _chosen["inner_l"], _chosen["inner_w"]
            bag_h = _layers_in_bag * max(product)
            bag_name = _chosen["박스명"]

    # --- 고급 옵션 (기본 접힘) ---
    with st.expander("⚙️ 고급 옵션"):
        use_best = st.toggle("최적 방향(회전) 적재", value=True,
                             help="제품을 6방향으로 돌려 최대 적재수량을 계산합니다.")
        wall_margin = st.number_input(
            "박스 벽두께 여유 (mm)", min_value=0.0, value=0.0, step=1.0,
            help="박스 규격이 외경일 때, 벽두께만큼 빼고 계산합니다. (0=규격 그대로)")

    st.markdown("")
    entity = copy.deepcopy(DEFAULT_ENTITY)
    with st.expander("💰 포장 인건비 (선택)"):
        entity["packing_labor_per_box"] = st.number_input(
            "박스당 포장 인건비 (KRW)", min_value=0.0,
            value=float(DEFAULT_ENTITY["packing_labor_per_box"]), step=100.0,
            help="박스당 총원가 = 박스 단가 + 포장 인건비")

# ---------------------------------------------------------------------------
# 메인 - 계산 (제품 → 포장재 → 박스)
# ---------------------------------------------------------------------------
outer_boxes = BOX_CATALOG[outer_group]

# 트레이 칸수 (범용트레이 선택 시)
tray_cells, tray_grid = 0, (0, 0)
if is_tray and tray_l > 0 and tray_w > 0:
    tray_cells, tray_grid = tray_cell_count(
        product, {"inner_l": tray_l, "inner_w": tray_w},
        gap=tray_gap, pitch_x=tray_pitch_x, pitch_y=tray_pitch_y)

rows = build_packaging_rows(
    product, outer_boxes, entity, inner_mode=inner_mode, outer_group=outer_group,
    unit_weight_g=unit_weight_g, part_name=part_name, wall_margin=wall_margin,
    use_best=use_best, tray_cells=tray_cells, tray_grid=tray_grid,
    tray_l=tray_l, tray_w=tray_w, tray_thickness=tray_thickness,
    bag_name=bag_name, bag_count=bag_count, bag_l=bag_l, bag_w=bag_w, bag_h=bag_h,
    weight_limit_kg=weight_limit_kg)

best_row = max(rows, key=lambda r: r["박스당 총 제품"]) if rows else None

# 컨텍스트 칩 바
_wchip = f'<span class="chip"><b>무게</b> {unit_weight_g:g} g</span>' if unit_weight_g else ""
_bchip = f'<span class="chip"><b>지퍼백</b> {bag_name}</span>' if (is_bag and bag_name) else ""
_pchip = f'<span class="chip"><b>품명</b> {part_name}</span>' if part_name else ""
st.markdown(
    '<div class="ctxbar">'
    f'{_pchip}'
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
        elif is_bag:
            cards.append({"label": "지퍼백 1개당 (입수)", "value": f"{bag_count:,}",
                          "unit": "개",
                          "sub": f"봉투 {bag_l:g}×{bag_w:g}×{bag_h:g} · 적합 {bag_name}"})
        cards.append({"label": f"🏆 추천 박스 · {best_row['박스명']}",
                      "value": f"{best_row['박스당 총 제품']:,}", "unit": "개",
                      "sub": f"규격 {best_row['규격(Size)']} · 최다 적재",
                      "op": "→" if (is_tray or is_bag) else "", "variant": "total"})
        cards.append({"label": "검토 박스", "value": f"{len(rows)}", "unit": "종",
                      "sub": outer_group})
        kpi_row(cards)
        if best_row["박스당 총 제품"] == 0:
            st.warning("이 조합으로는 적재되지 않습니다. 제품 사이즈·트레이 설정·박스 종류를 확인하세요.")
        st.markdown("")

        # 적재 배치 이미지 (최다 적재 박스 기준)
        section("적재 배치도")
        packing_image(best_row["_cols"], best_row["_rows"], best_row["_layers"],
                      best_row["_unit"], best_row["박스명"], best_row["박스당 총 제품"])
        st.markdown("")

        df = pd.DataFrame(rows).drop(columns=["_cols", "_rows", "_layers", "_unit"])
        _mx = df["박스당 총 제품"].max()
        df.insert(0, "추천", df["박스당 총 제품"].apply(
            lambda v: "🏆" if v == _mx and v > 0 else ""))
        section(f"{outer_group} · 박스별 총 제품 수")
        show_cols = ["추천", "박스명", "규격(Size)", "포장재", "적재 방식",
                     "박스당 총 제품", "제한 요인", "비고"]
        if unit_weight_g:
            show_cols.insert(6, "박스 총중량(kg)")
        st.dataframe(
            df[show_cols], use_container_width=True, hide_index=True,
            column_config={
                "박스당 총 제품": st.column_config.NumberColumn(format="%d 개"),
            },
        )

        cta, info_ = st.columns([1, 1])
        cta.info("📄 견적서로 만들려면 상단 **'구매팀 견적 양식'** 탭으로 이동하세요.")
        with info_.expander("ℹ️ 계산 방식 보기"):
            st.markdown(
                "- **없음(벌크)**: 제품을 박스에 직접 3D 적재 "
                "⌊박스 ÷ 제품⌋ (최적 방향 6방향 중 최대)\n"
                "- **지퍼백**: 지퍼백 입수(1봉지당 제품 수) × (박스당 봉지 수) = 박스당 총 제품\n"
                "- **범용트레이**: 트레이 칸수 × (박스당 트레이 장수) = 박스당 총 제품\n"
                "- **무게 입력 시**: 박스 허용중량을 넘지 않도록 제한 → '제한 요인'에 표시")

# --- 탭2: 견적 양식 + 다운로드 ---
with tab_form:
    section("구매팀 전달용 표준 견적 요청서")
    header_info = default_header_info("성우", entity, outer_group, product,
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
        editable = pd.DataFrame(rows).drop(
            columns=["_cols", "_rows", "_layers", "_unit"], errors="ignore")
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
                file_name=f'견적요청서_{outer_group}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ PDF 견적 요청서 (.pdf)",
                data=to_pdf_bytes(header_info, export_rows),
                file_name=f'견적요청서_{outer_group}.pdf',
                mime="application/pdf",
                use_container_width=True,
            )

# --- 탭3: 기준 데이터 열람 ---
with tab_data:
    section("표준 포장재 카탈로그 (기준 데이터)")
    st.caption("실제 값 수정은 `data.py`의 리스트(CARTONS·DANPLA·PLASTIC·ZIPPERS·TRAYS)에서 관리합니다.")
    view = st.selectbox("분류 선택", list(BOX_CATALOG.keys()))
    st.dataframe(pd.DataFrame(BOX_CATALOG[view]), use_container_width=True, hide_index=True)

st.divider()
st.caption("© 성우 개발팀 · 포장 사양 견적 대시보드")
