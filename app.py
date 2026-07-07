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
import plotly.graph_objects as go

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
.block-container{padding-top:1.4rem;padding-bottom:2.6rem;max-width:1340px;}
h1{font-size:1.9rem !important;font-weight:800;}
h1,h2,h3{letter-spacing:-.015em;}

/* ---------- 사이드바 ---------- */
section[data-testid="stSidebar"]{border-right:1px solid var(--border);
  background:linear-gradient(180deg,#12161d,#0f1319);}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.55rem;}
section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"]{gap:.5rem;}
section[data-testid="stSidebar"] label p,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{
  font-size:.78rem !important;color:var(--muted) !important;font-weight:600;margin-bottom:1px;}
[data-baseweb="input"],[data-baseweb="select"]>div{
  background:var(--surface2) !important;border-radius:9px !important;
  border:1px solid var(--border) !important;}
[data-baseweb="input"]:focus-within{border-color:var(--accent) !important;
  box-shadow:0 0 0 3px rgba(57,135,229,.18) !important;}
.stNumberInput input,.stTextInput input{font-variant-numeric:tabular-nums;}
[data-testid="stExpander"]{border:1px solid var(--border);border-radius:10px;
  background:var(--surface);}

/* 스텝 배지 */
.step{display:flex;align-items:center;gap:9px;margin:.25rem 0 .4rem;font-weight:800;
  font-size:1rem;color:var(--text);}
.step .num{display:inline-flex;align-items:center;justify-content:center;width:23px;height:23px;
  border-radius:7px;background:linear-gradient(135deg,#3987e5,#2a6fc4);color:#fff;
  font-size:.8rem;font-weight:800;box-shadow:0 2px 8px rgba(57,135,229,.4);}

/* ---------- 컨텍스트 칩 ---------- */
.ctxbar{display:flex;gap:8px;flex-wrap:wrap;margin:.1rem 0 1.1rem;}
.chip{background:var(--surface);border:1px solid var(--border);border-radius:999px;
  padding:5px 13px;font-size:.82rem;color:var(--text);white-space:nowrap;}
.chip b{color:var(--muted);font-weight:600;margin-right:6px;font-size:.76rem;}

/* ---------- 섹션 헤더 ---------- */
.sec{display:flex;align-items:center;gap:9px;margin:1.15rem 0 .6rem;font-weight:750;
  font-size:1.08rem;color:var(--text);}
.sec .dot{width:9px;height:9px;border-radius:50%;background:var(--accent);
  box-shadow:0 0 0 4px rgba(57,135,229,.16);}

/* ---------- KPI 카드 ---------- */
.kpi-row{display:flex;gap:16px;align-items:stretch;flex-wrap:wrap;margin:.2rem 0 .3rem;}
.kpi-card{flex:1;min-width:160px;
  background:linear-gradient(160deg,#1a2029,#141922);
  border:1px solid #2a3140;border-radius:16px;padding:18px 20px;
  position:relative;overflow:hidden;
  box-shadow:0 1px 0 rgba(255,255,255,.03) inset,0 8px 24px -14px rgba(0,0,0,.6);
  animation:kpiIn .5s cubic-bezier(.2,.7,.3,1) both;
  transition:transform .18s ease,border-color .18s ease,box-shadow .18s ease;}
.kpi-card:hover{transform:translateY(-4px);border-color:#3d4757;
  box-shadow:0 16px 36px -16px rgba(0,0,0,.75);}
.kpi-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;
  background:linear-gradient(180deg,var(--accent),#2a6fc4);}
.kpi-card::after{content:"";position:absolute;right:-42px;top:-42px;width:120px;height:120px;
  border-radius:50%;background:radial-gradient(circle,rgba(57,135,229,.14),transparent 70%);}
.kpi-card.total{flex:1.35;
  background:linear-gradient(160deg,rgba(25,158,112,.22),#141c1c 62%);
  border-color:rgba(64,214,160,.45);
  box-shadow:0 0 0 1px rgba(64,214,160,.12) inset,0 14px 40px -16px rgba(20,120,84,.6);}
.kpi-card.total::before{background:linear-gradient(180deg,#40d6a0,#199e70);}
.kpi-card.total::after{background:radial-gradient(circle,rgba(64,214,160,.22),transparent 70%);}
.kpi-label{font-size:.8rem;color:var(--muted);margin-bottom:8px;font-weight:600;
  position:relative;z-index:1;}
.kpi-value{font-size:2.15rem;font-weight:800;color:var(--text);line-height:1.02;
  font-variant-numeric:tabular-nums;position:relative;z-index:1;}
.kpi-card.total .kpi-value{font-size:2.7rem;color:#48e0aa;
  text-shadow:0 0 26px rgba(64,214,160,.35);}
.kpi-unit{font-size:.95rem;font-weight:600;color:var(--muted);margin-left:5px;}
.kpi-sub{font-size:.74rem;color:var(--muted);margin-top:9px;position:relative;z-index:1;}
.kpi-op{display:flex;align-items:center;font-size:1.6rem;color:#5b6b7d;font-weight:800;}
@keyframes kpiIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:none;}}

/* ---------- 배치도 (Plotly 컨테이너 + 수치 카드) ---------- */
[data-testid="stPlotlyChart"]{border:1px solid var(--border);border-radius:16px;
  overflow:hidden;background:linear-gradient(160deg,#161b22,#12161d);padding:6px;
  box-shadow:0 12px 32px -18px rgba(0,0,0,.75);}
.stat{background:linear-gradient(160deg,#1a2029,#141922);border:1px solid #2a3140;
  border-radius:13px;padding:13px 16px;margin-bottom:12px;position:relative;overflow:hidden;
  box-shadow:0 6px 18px -12px rgba(0,0,0,.6);}
.stat::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent);}
.stat .l{font-size:.74rem;color:var(--muted);margin-bottom:5px;font-weight:600;}
.stat .v{font-size:1.5rem;font-weight:800;color:var(--text);
  font-variant-numeric:tabular-nums;line-height:1.05;}
.stat .u{font-size:.8rem;color:var(--muted);font-weight:600;margin-left:4px;}
.stat.hi{background:linear-gradient(160deg,rgba(25,158,112,.22),#141c1c 62%);
  border-color:rgba(64,214,160,.45);}
.stat.hi::before{background:linear-gradient(180deg,#40d6a0,#199e70);}
.stat.hi .v{color:#48e0aa;font-size:1.95rem;text-shadow:0 0 20px rgba(64,214,160,.3);}

/* ---------- 탭(라디오 기반) ---------- */
div[role="radiogroup"]{gap:4px;border-bottom:1px solid var(--border);margin-bottom:.4rem;}
div[role="radiogroup"] > label{background:transparent;border:1px solid transparent;
  border-bottom:none;border-radius:10px 10px 0 0;padding:9px 18px;margin:0;
  cursor:pointer;transition:background .15s,color .15s;font-weight:600;}
div[role="radiogroup"] > label:hover{background:rgba(57,135,229,.08);}
div[role="radiogroup"] > label > div:first-child{display:none;}   /* 라디오 동그라미 숨김 */
div[role="radiogroup"] > label:has(input:checked){background:rgba(57,135,229,.12);
  border-color:var(--border);color:#7ec8f3;
  box-shadow:inset 0 -2px 0 var(--accent);}

/* ---------- 견적 유도 CTA ---------- */
.cta{display:flex;align-items:center;justify-content:space-between;gap:14px;
  margin:.2rem 0 .3rem;padding:16px 20px;border-radius:14px;
  background:linear-gradient(100deg,rgba(57,135,229,.16),rgba(25,158,112,.10));
  border:1px solid #2f4560;transition:border-color .16s;}
.cta:hover{border-color:#3d6ea0;}
.cta .t{color:var(--text);font-weight:700;font-size:.98rem;}
.cta .d{color:var(--muted);font-size:.8rem;margin-top:3px;}
.cta .go{white-space:nowrap;background:linear-gradient(135deg,#3987e5,#2a6fc4);
  color:#fff;font-weight:700;font-size:.9rem;padding:10px 18px;border-radius:10px;
  box-shadow:0 6px 18px -6px rgba(57,135,229,.7);}
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


_CUBE_FACES = [(0, 1, 2), (0, 2, 3), (4, 5, 6), (4, 6, 7), (0, 1, 5), (0, 5, 4),
               (1, 2, 6), (1, 6, 5), (2, 3, 7), (2, 7, 6), (3, 0, 4), (3, 4, 7)]
_CUBE_VERTS = [(0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0),
               (0, 0, 1), (1, 0, 1), (1, 1, 1), (0, 1, 1)]


def _cuboids_mesh(cells, color, opacity):
    """여러 개의 3D 직육면체(제품 블록)를 하나의 Mesh3d로 묶어 반환 (성능)."""
    s = 0.86
    X, Y, Z, I, J, K = [], [], [], [], [], []
    for ox, oy, oz in cells:
        base = len(X)
        for vx, vy, vz in _CUBE_VERTS:
            X.append(ox + vx * s); Y.append(oy + vy * s); Z.append(oz + vz * s)
        for a, b, c in _CUBE_FACES:
            I.append(base + a); J.append(base + b); K.append(base + c)
    return go.Mesh3d(x=X, y=Y, z=Z, i=I, j=J, k=K, color=color, opacity=opacity,
                     flatshading=True, hoverinfo="skip", lighting=dict(ambient=.55,
                     diffuse=.8, specular=.2))


def _box_edges(nx, ny, nz):
    """박스(Carton) 투명 가이드라인 테두리 (12모서리)."""
    p = [(0, 0, 0), (nx, 0, 0), (nx, ny, 0), (0, ny, 0),
         (0, 0, nz), (nx, 0, nz), (nx, ny, nz), (0, ny, nz)]
    edges = [(0, 1), (1, 2), (2, 3), (3, 0), (4, 5), (5, 6), (6, 7), (7, 4),
             (0, 4), (1, 5), (2, 6), (3, 7)]
    X, Y, Z = [], [], []
    for a, b in edges:
        X += [p[a][0], p[b][0], None]
        Y += [p[a][1], p[b][1], None]
        Z += [p[a][2], p[b][2], None]
    return go.Scatter3d(x=X, y=Y, z=Z, mode="lines",
                        line=dict(color="#5b6b7d", width=4), hoverinfo="skip")


def packing_fig_3d(nx, ny, nz, highlight=None):
    """제품 블록이 박스 안에 3D로 적층된 모습 (드래그 회전/확대). highlight=강조할 층(1~)."""
    cap = 8
    dx = max(min(int(nx), cap), 1)
    dy = max(min(int(ny), cap), 1)
    dz = max(min(int(nz), cap), 1)
    dim, hi = [], []
    for k in range(dz):
        for j in range(dy):
            for i in range(dx):
                (hi if (highlight and k + 1 == highlight) else dim).append((i, j, k))
    data = [_box_edges(dx, dy, dz)]
    if dim:
        data.append(_cuboids_mesh(dim, "#3987e5", 0.42 if highlight else 0.9))
    if hi:
        data.append(_cuboids_mesh(hi, "#40d6a0", 0.97))
    fig = go.Figure(data)
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=0, b=0),
        height=460, showlegend=False,
        scene=dict(
            xaxis=dict(visible=False), yaxis=dict(visible=False), zaxis=dict(visible=False),
            bgcolor="rgba(0,0,0,0)", aspectmode="data",
            camera=dict(eye=dict(x=1.6, y=1.5, z=1.15)),
        ),
    )
    return fig


def stat_card(label, value, unit="", hi=False):
    return (f'<div class="stat {"hi" if hi else ""}">'
            f'<div class="l">{label}</div>'
            f'<div class="v">{value}<span class="u">{unit}</span></div></div>')


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

            # 봉투별 입수 가이드라인 (이 제품 기준, 한 겹)
            _caps = [(bg["박스명"], bg["size"], bag_layer_capacity(product, bg))
                     for bg in _bags]
            _pos = [c for _, _, c in _caps if c > 0]
            if _pos:
                st.caption(f"📋 봉투별 입수 가이드: **{min(_pos):,} ~ {max(_pos):,}개** "
                           f"(제품 {pl:g}×{pw:g}×{ph:g} 기준 · 한 겹)")
                if st.checkbox("봉투별 입수 상세 보기", value=False):
                    st.dataframe(
                        pd.DataFrame([{"지퍼백": n, "규격": s,
                                       "입수(개)": c if c > 0 else "안 들어감"}
                                      for n, s, c in _caps]),
                        hide_index=True, use_container_width=True)
            else:
                st.warning("이 제품이 들어가는 지퍼백이 없습니다. 규격을 확인하세요.")

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
VIEWS = ["📊 적재 효율 계산", "📄 구매팀 견적 양식", "🗄️ 기준 데이터(박스 리스트)"]
if "active_view" not in st.session_state:
    st.session_state.active_view = VIEWS[0]

view = st.radio("보기", VIEWS, key="active_view",
                horizontal=True, label_visibility="collapsed")

# --- 탭1: 적재 효율 ---
if view == VIEWS[0]:
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

        # 적재 배치도 (좌: Plotly 시각화 / 우: 수치 카드)
        section(f"적재 배치도 · {best_row['박스명']}")
        _c, _r = int(best_row["_cols"]), int(best_row["_rows"])
        _lay, _u = int(best_row["_layers"]), best_row["_unit"]
        _total = best_row["박스당 총 제품"]
        _per_layer = _c * _r
        _geom = _per_layer * _lay                       # 부피 기준 최대
        _capped = _per_layer > 0 and _total < _geom     # 무게 한도 등으로 줄었는지
        _eff_layers = -(-_total // _per_layer) if _per_layer > 0 else 0
        viz, stt = st.columns([2, 1], gap="large")
        with viz:
            if _c > 0 and _r > 0:
                _shown_layers = min(_lay, 8)
                hl = None
                if _shown_layers > 1:
                    hl = st.slider("적층 확인 (층)", 1, _shown_layers, 1,
                                   help="선택한 층이 에메랄드색으로 강조됩니다. "
                                        "차트를 드래그하면 회전, 스크롤하면 확대됩니다.")
                st.plotly_chart(packing_fig_3d(_c, _r, _lay, highlight=hl),
                                use_container_width=True,
                                config={"displayModeBar": False})
                trunc = _c > 8 or _r > 8 or _lay > 8
                cap_note = (f"  ·  ⚠️ 부피상 {_geom:,}개 가능하나 "
                            f"'{best_row['제한 요인']}'으로 {_total:,}개만 적재") if _capped else ""
                st.caption(f"🧊 {_c}×{_r}×{_lay} 적층 · 드래그로 회전 / 스크롤로 확대"
                           + ("  (그림은 8×8×8까지 대표 표시)" if trunc else "")
                           + cap_note)
            else:
                st.info("이 조합은 적재되지 않습니다.")
        with stt:
            st.markdown(
                stat_card("박스당 총 제품", f"{_total:,}", "개", hi=True)
                + stat_card(f"층당 개수 ({_c}×{_r})", f"{_per_layer:,}", _u)
                + stat_card("실제 적층 단수", f"{_eff_layers:,}", "층")
                + stat_card("적재 방식", best_row["적재 방식"], ""),
                unsafe_allow_html=True)
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

        st.markdown("")
        st.markdown(
            '<div class="cta"><div>'
            '<div class="t">📄 이 구성으로 견적 요청서를 만들 준비가 됐어요</div>'
            '<div class="d">아래 버튼을 누르면 <b>구매팀 견적 양식</b> 화면으로 이동해 '
            "Excel · PDF로 내보내고 '구매 확정 단가'까지 받을 수 있어요.</div>"
            '</div></div>', unsafe_allow_html=True)
        st.button("📄 구매팀 견적 양식으로 이동  →", type="primary",
                  use_container_width=True,
                  on_click=lambda: st.session_state.update(active_view=VIEWS[1]))
        with st.expander("ℹ️ 계산 방식 보기"):
            st.markdown(
                "- **없음(벌크)**: 제품을 박스에 직접 3D 적재 "
                "⌊박스 ÷ 제품⌋ (최적 방향 6방향 중 최대)\n"
                "- **지퍼백**: 지퍼백 입수(1봉지당 제품 수) × (박스당 봉지 수) = 박스당 총 제품\n"
                "- **범용트레이**: 트레이 칸수 × (박스당 트레이 장수) = 박스당 총 제품\n"
                "- **무게 입력 시**: 박스 허용중량을 넘지 않도록 제한 → '제한 요인'에 표시")

# --- 탭2: 견적 양식 + 다운로드 ---
elif view == VIEWS[1]:
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
else:
    section("표준 포장재 카탈로그 (기준 데이터)")
    st.caption("실제 값 수정은 `data.py`의 리스트(CARTONS·DANPLA·PLASTIC·ZIPPERS·TRAYS)에서 관리합니다.")
    view = st.selectbox("분류 선택", list(BOX_CATALOG.keys()))
    st.dataframe(pd.DataFrame(BOX_CATALOG[view]), use_container_width=True, hide_index=True)

st.divider()
st.caption("© 성우 개발팀 · 포장 사양 견적 대시보드")
