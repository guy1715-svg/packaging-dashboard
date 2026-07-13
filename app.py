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

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from data import (BOX_CATALOG, INNER_OPTIONS, OUTER_GROUPS,
                  BAG_GROUP, TRAY_GROUP, PALLETS, CONTAINERS)
from calculations import (build_packaging_rows, tray_cell_count,
                          fit_zipper_bag, bag_layer_capacity,
                          boxes_per_pallet, boxes_per_container, cbm)
import store
import mesh_loader
from datetime import datetime
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
/* 숫자 입력창 +/- 스텝퍼 제거 (깔끔하게) */
[data-testid="stNumberInput"] button{display:none !important;}
[data-testid="stNumberInput"] input{text-align:left;padding-right:10px;}
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
.kpi-card{flex:1;min-width:170px;min-height:142px;
  display:flex;flex-direction:column;justify-content:center;
  background:linear-gradient(160deg,#1a2029,#141922);
  border:1px solid #2a3140;border-radius:16px;padding:22px 24px;
  position:relative;overflow:hidden;
  box-shadow:0 1px 0 rgba(255,255,255,.03) inset,0 8px 24px -14px rgba(0,0,0,.6);
  animation:kpiIn .5s cubic-bezier(.2,.7,.3,1) both;
  transition:transform .18s ease,border-color .18s ease,box-shadow .18s ease;}
.kpi-card:hover{transform:translateY(-4px);border-color:#3d4757;
  box-shadow:0 16px 36px -16px rgba(0,0,0,.75);}
.kpi-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;
  background:linear-gradient(180deg,var(--accent),#2a6fc4);opacity:.7;}
.kpi-card.total{flex:1.35;
  background:linear-gradient(160deg,rgba(25,158,112,.15),#141c1c 66%);
  border-color:rgba(64,214,160,.30);
  box-shadow:0 0 0 1px rgba(64,214,160,.07) inset,0 14px 46px -22px rgba(64,214,160,.42);}
.kpi-card.total:hover{border-color:rgba(64,214,160,.45);
  box-shadow:0 0 0 1px rgba(64,214,160,.10) inset,0 20px 52px -22px rgba(64,214,160,.5);}
.kpi-card.total::before{background:linear-gradient(180deg,#40d6a0,#199e70);opacity:1;}
/* 빈 자리표시용 유령 카드 — 행 간 그리드 정렬을 위해 공간만 차지, 보이지 않음 */
.kpi-card.ghost{background:transparent;border-color:transparent;box-shadow:none;
  animation:none;}
.kpi-card.ghost::before{display:none;}
.kpi-card.ghost:hover{transform:none;box-shadow:none;border-color:transparent;}
.kpi-label{font-size:.8rem;color:var(--muted);margin-bottom:8px;font-weight:600;
  position:relative;z-index:1;}
.kpi-value{font-size:2.15rem;font-weight:800;color:var(--text);line-height:1.02;
  font-variant-numeric:tabular-nums;position:relative;z-index:1;}
.kpi-card.total .kpi-value{font-size:2.7rem;color:#48e0aa;
  text-shadow:0 0 20px rgba(64,214,160,.24);}
.kpi-unit{font-size:.95rem;font-weight:600;color:var(--muted);margin-left:5px;}
.kpi-sub{font-size:.74rem;color:var(--muted);margin-top:9px;position:relative;z-index:1;}
.kpi-op{display:flex;align-items:center;font-size:1.6rem;color:#5b6b7d;font-weight:800;}
@keyframes kpiIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:none;}}

/* ---------- 치수 뱃지 ---------- */
.badges{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:-2px 0 12px;}
.badge{background:#1c2432;border:1px solid #2a3140;border-radius:9px;
  padding:5px 12px;font-size:.84rem;color:#9fb0c0;letter-spacing:.01em;}
.badge b{color:#eaf2fb;font-weight:700;font-variant-numeric:tabular-nums;}
.badge.g{border-color:rgba(64,214,160,.35);}
.badge.g b{color:#48e0aa;}
.badge.unit{background:transparent;border:none;color:#6b7684;padding-left:2px;}

/* ---------- 슬라이더 에메랄드 통일 ---------- */
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"]{
  background-color:#34d399 !important;
  box-shadow:0 0 0 4px rgba(52,211,153,.22) !important;}
/* 채워진 트랙(파랑)을 hue-rotate로 에메랄드화 (채움 위치 유지) */
[data-testid="stSlider"] [data-baseweb="slider"] > div:nth-child(1) > div:nth-child(1) > div:nth-child(2){
  filter:hue-rotate(-58deg) saturate(1.15) !important;}
[data-testid="stSlider"] [data-testid="stThumbValue"]{color:#48e0aa !important;font-weight:700;}

/* ---------- 배치도 (Plotly 컨테이너 + 수치 카드) ---------- */
[data-testid="stPlotlyChart"]{border:1px solid var(--border);border-radius:16px;
  overflow:hidden;background:linear-gradient(160deg,#161b22,#12161d);padding:2px;
  box-shadow:0 12px 32px -18px rgba(0,0,0,.75);}
.stat{background:linear-gradient(160deg,#1a2029,#141922);border:1px solid #2a3140;
  border-radius:13px;padding:13px 16px;margin-bottom:12px;position:relative;overflow:hidden;
  box-shadow:0 6px 18px -12px rgba(0,0,0,.6);}
.stat::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent);}
.stat .l{font-size:.74rem;color:var(--muted);margin-bottom:5px;font-weight:600;}
.stat .v{font-size:1.5rem;font-weight:800;color:var(--text);
  font-variant-numeric:tabular-nums;line-height:1.05;}
.stat .u{font-size:.8rem;color:var(--muted);font-weight:600;margin-left:4px;}
.stat.hi{background:linear-gradient(160deg,rgba(25,158,112,.10),#141922 66%);
  border:1px solid rgba(64,214,160,.22);
  box-shadow:0 6px 18px -12px rgba(0,0,0,.6);}
.stat.hi::before{background:linear-gradient(180deg,#40d6a0,#199e70);}
.stat.hi .v{color:#5fd6a6;font-size:1.95rem;}

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

/* ---------- 기록 저장 카드(폼) ---------- */
[data-testid="stForm"]{border:1px solid var(--border) !important;border-radius:16px;
  background:linear-gradient(180deg,#161b23,#12161d);
  padding:22px 24px 18px !important;box-shadow:0 10px 30px -18px rgba(0,0,0,.7);}
[data-testid="stForm"] [data-testid="stHorizontalBlock"]{gap:1.4rem;}
[data-testid="stForm"] [data-testid="stVerticalBlock"]{gap:.9rem;}
[data-testid="stForm"] label p{font-size:.82rem !important;color:var(--muted) !important;
  font-weight:600;}
[data-testid="stForm"] textarea{background:var(--surface2) !important;
  border-radius:10px !important;}
</style>
""", unsafe_allow_html=True)


def kpi_row(cards):
    """KPI 카드 한 줄 렌더링. cards: [{label,value,unit,sub,variant,op}]"""
    html = ['<div class="kpi-row">']
    for c in cards:
        if c.get("op"):
            html.append(f'<div class="kpi-op">{c["op"]}</div>')
        _style = f' style="flex:{c["w"]}"' if c.get("w") else ""
        html.append(
            f'<div class="kpi-card {c.get("variant","")}"{_style}>'
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
                     flatshading=True, hoverinfo="skip",
                     lighting=dict(ambient=.62, diffuse=.85, specular=.12,
                                   roughness=.55, fresnel=.1),
                     lightposition=dict(x=1200, y=800, z=2000))


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


def _dim_labels(dx, dy, dz, box_l, box_w, box_h):
    """박스 가로/세로/높이 치수 텍스트 (박스 바깥, 잘 보이게)."""
    off = 1.6
    xs, ys, zs, txt, col = [], [], [], [], []
    if box_l:
        xs.append(dx / 2); ys.append(-off); zs.append(0)
        txt.append(f"↔ 가로 {box_l:g}mm"); col.append("#8fd0ff")
    if box_w:
        xs.append(-off); ys.append(dy / 2); zs.append(0)
        txt.append(f"↔ 세로 {box_w:g}mm"); col.append("#8fd0ff")
    if box_h:
        xs.append(-off); ys.append(-off); zs.append(dz / 2)
        txt.append(f"↕ 높이 {box_h:g}mm"); col.append("#6ff0c0")
    return go.Scatter3d(x=xs, y=ys, z=zs, mode="text", text=txt,
                        textfont=dict(color=col, size=16, family="system-ui"),
                        hoverinfo="skip")


def packing_fig_3d(nx, ny, nz, active_layers=None, top_count=None,
                   box_l=0, box_w=0, box_h=0):
    """
    제품 블록이 박스 안에 3D로 적층된 모습 (드래그 회전/확대).
    active_layers=N 이면 1~N층을 채워 표시(맨 위 층 초록 강조), N+1층부터는 반투명 잔여 공간.
    top_count 을 주면 맨 위(N층)에는 그 개수만 채워 그림(무게 제한 자투리 층 표현).
    """
    XY_CAP, Z_CAP = 8, 40      # 바닥 격자는 8×8까지 축약, 층수는 40층까지 실제 표시
    dx = max(min(int(nx), XY_CAP), 1)
    dy = max(min(int(ny), XY_CAP), 1)
    dz = max(min(int(nz), Z_CAP), 1)
    N = active_layers if active_layers else dz
    layer_area = dx * dy
    # 맨 위 층에 실제로 놓일 블록 수(자투리). None/과대 → 가득 채움.
    tc = layer_area if top_count is None else max(0, min(int(top_count), layer_area))
    solid, top, ghost = [], [], []
    for k in range(dz):
        placed_in_top = 0
        for j in range(dy):
            for i in range(dx):
                if k + 1 < N:
                    solid.append((i, j, k))
                elif k + 1 == N:
                    if placed_in_top < tc:
                        top.append((i, j, k))
                    else:
                        ghost.append((i, j, k))
                    placed_in_top += 1
                else:
                    ghost.append((i, j, k))
    # 치수 텍스트는 위쪽 배지·'박스 규격' 카드에 이미 표시되므로 3D에는 생략(잘림·중복 방지)
    data = [_box_edges(dx, dy, dz)]
    if ghost:
        data.append(_cuboids_mesh(ghost, "#3a5a86", 0.09))   # 잔여공간(옅은 코발트)
    if solid:
        data.append(_cuboids_mesh(solid, "#4479c4", 0.84))   # 코발트 블루
    if top:
        data.append(_cuboids_mesh(top, "#2fbf8f", 0.9))      # 에메랄드
    # 실제 박스 비율(L:W:H)로 스케일 → 직사각형/정사각형 박스가 눈으로 구분됨
    _m = max(box_l, box_w, box_h) or 1
    aspect = dict(x=(box_l / _m) or 1, y=(box_w / _m) or 1, z=(box_h / _m) or 1) \
        if (box_l and box_w and box_h) else dict(x=1, y=1, z=1)
    # 축·격자·배경을 모두 끈 '허공' 씬 (박스만 깔끔히 떠 있게)
    _clean_axis = dict(visible=False, showgrid=False, showline=False,
                       zeroline=False, showbackground=False, showticklabels=False,
                       showspikes=False, title="")
    fig = go.Figure(data)
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=0, b=0),
        height=430, showlegend=False,
        scene=dict(
            xaxis=_clean_axis, yaxis=_clean_axis, zaxis=_clean_axis,
            bgcolor="rgba(0,0,0,0)", aspectmode="manual", aspectratio=aspect,
            domain=dict(x=[0, 1], y=[0, 1]),
            camera=dict(eye=dict(x=1.55, y=1.45, z=1.08)),
        ),
    )
    return fig


def mesh_preview_fig(geom):
    """업로드한 실제 제품 메쉬를 3D로 렌더 (형상 확인용)."""
    verts, faces = geom["verts"], geom["faces"]
    xs = [v[0] for v in verts]; ys = [v[1] for v in verts]; zs = [v[2] for v in verts]
    i = [f[0] for f in faces]; j = [f[1] for f in faces]; k = [f[2] for f in faces]
    L, W, H = geom["dims"]
    fig = go.Figure(go.Mesh3d(
        x=xs, y=ys, z=zs, i=i, j=j, k=k,
        color="#4479c4", opacity=1.0, flatshading=True,
        lighting=dict(ambient=0.55, diffuse=0.8, specular=0.2, roughness=0.6),
        lightposition=dict(x=100, y=200, z=300), hoverinfo="skip"))
    span = max(max(xs) - min(xs), max(ys) - min(ys), max(zs) - min(zs), 1)
    fig.update_layout(
        height=340, margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        scene=dict(
            xaxis=dict(visible=False), yaxis=dict(visible=False), zaxis=dict(visible=False),
            bgcolor="rgba(0,0,0,0)", aspectmode="manual",
            aspectratio=dict(
                x=(max(xs) - min(xs)) / span or 0.1,
                y=(max(ys) - min(ys)) / span or 0.1,
                z=(max(zs) - min(zs)) / span or 0.1),
            camera=dict(eye=dict(x=1.5, y=1.5, z=1.1)),
        ))
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
    sidebar_step(1, "제품 입력")

    with st.expander("🏷️ 기본 정보 (고객사·품명) · 선택"):
        customer = st.text_input("고객사", value="", placeholder="예: 현대모비스")
        part_name = st.text_input("품명 (제품명/품번)", value="",
                                  placeholder="예: SW-CONN-0250")

    # --- 3D 파일(STL/OBJ)에서 치수 자동 추출 ---
    for _k, _dv in (("prod_l", 50.0), ("prod_w", 30.0), ("prod_h", 15.0)):
        st.session_state.setdefault(_k, _dv)

    with st.expander("📐 3D 파일로 치수 자동입력 (NX → STL)"):
        if not mesh_loader.available():
            st.caption("⚠️ 3D 라이브러리(trimesh) 미설치. requirements.txt 반영 후 재배포하세요.")
        up3d = st.file_uploader("STL / OBJ / PLY 업로드",
                                type=mesh_loader.SUPPORTED, key="mesh_up",
                                help="NX: 파일 → 내보내기 → STL 로 저장 후 업로드하세요. "
                                     "제품 좌표축 기준으로 L/W/H를 측정합니다 (NX 측정값과 동일).")
        if up3d is not None:
            sig = (up3d.name, up3d.size)
            if st.session_state.get("_mesh_sig") != sig:
                data = up3d.getvalue()
                res = mesh_loader.load_mesh(data, up3d.name)
                if res is None:
                    st.session_state["_mesh_geom"] = None
                    st.error("3D 파일을 읽지 못했습니다. STL/OBJ/PLY 인지 확인하세요. "
                             "(STEP/STP는 미지원 — NX에서 STL로 내보내세요)")
                else:
                    L, W, H = res["dims"]
                    st.session_state["prod_l"] = float(L)
                    st.session_state["prod_w"] = float(W)
                    st.session_state["prod_h"] = float(H)
                    st.session_state["_mesh_geom"] = res
                    st.session_state["_mesh_name"] = up3d.name
                st.session_state["_mesh_sig"] = sig
                st.rerun()
        _mg = st.session_state.get("_mesh_geom")
        if _mg:
            L, W, H = _mg["dims"]
            st.success(f"📥 {st.session_state.get('_mesh_name','')} → "
                       f"**{L}×{W}×{H} mm** 자동입력됨")
            if _mg.get("note"):
                st.caption("ℹ️ " + _mg["note"])

    st.markdown("**제품 외경 (mm)**")
    c1, c2, c3 = st.columns(3)
    pl = c1.number_input("L", min_value=0.1, step=1.0, key="prod_l")
    pw = c2.number_input("W", min_value=0.1, step=1.0, key="prod_w")
    ph = c3.number_input("H", min_value=0.1, step=1.0, key="prod_h")
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

    # --- 트레이 설정 (범용/제품 전용) ---
    tray_gap = tray_pitch_x = tray_pitch_y = 0.0
    tray_l = tray_w = tray_thickness = 0.0
    tray_custom = False
    tray_cells_manual = 0
    tray_material = tray_maker = tray_color = ""
    if is_tray:
        with st.expander("🔧 트레이 설정", expanded=True):
            tray_kind = st.radio(
                "트레이 종류", ["범용 (칸수 자동 계산)", "제품 전용 (칸수 직접 입력)"],
                help="제품 전용 트레이는 도면상 적재수량(칸수)이 정해져 있으므로 직접 입력합니다.")
            tray_custom = tray_kind.startswith("제품 전용")

            t1, t2 = st.columns(2)
            tray_l = t1.number_input("트레이 가로 (mm)", min_value=0.0, value=315.0, step=5.0)
            tray_w = t2.number_input("트레이 세로 (mm)", min_value=0.0, value=410.0, step=5.0)
            tray_thickness = st.number_input(
                "트레이 두께/높이 (mm)", min_value=0.0, value=15.0, step=1.0,
                help="트레이 1장 높이 → 박스에 몇 단 쌓이는지 계산에 사용")

            if tray_custom:
                tray_cells_manual = st.number_input(
                    "트레이 1장당 적재수량 (칸)", min_value=1, value=50, step=1,
                    help="제품 전용 트레이 도면의 칸 수를 직접 입력하세요.")
                st.markdown("**🆕 신규 트레이 제작 정보**")
                tm1, tm2 = st.columns(2)
                tray_material = tm1.text_input("재질", placeholder="예: PS 전도성")
                tray_color = tm2.text_input("색상", placeholder="예: 블랙")
                tray_maker = st.text_input("제작업체", placeholder="예: ○○트레이")
            else:
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
            # 전체 봉투 비교 (참고값) — 장황한 안내 대신 selectbox 툴팁 속으로
            _caps = [(bg["박스명"], bg["size"], bag_layer_capacity(product, bg))
                     for bg in _bags]
            _pos = [c for _, _, c in _caps if c > 0]
            _range_txt = (f" · 전체 봉투 입수 {min(_pos):,}~{max(_pos):,}개"
                          if _pos else "")
            _sel = st.selectbox(
                "지퍼백 규격 (자동 추천 · 변경 가능)", _names, index=_idx,
                help=f"제품 {pl:g}×{pw:g}×{ph:g} 기준, 들어가는 최소 규격을 자동 추천합니다."
                     f"{_range_txt}. 규격이 클수록 한 봉지에 많이 들어갑니다.")
            _chosen = _bags[_names.index(_sel)]
            _cap_sel = bag_layer_capacity(product, _chosen)   # 선택 봉투 입수
            _per_layer = _cap_sel or 1
            _size_disp = _chosen["size"].replace("*", "×")   # 마크다운 * 이탤릭 방지

            # ── 한 줄 요약 캡션 (규격 변경 시 즉시 재계산·반영) ──
            if _cap_sel > 0:
                st.caption(f"✅ 한 봉지 **{_cap_sel}개** · {_chosen['박스명']} ({_size_disp})")
            elif _pos:
                st.caption(f"⚠️ 이 규격엔 안 들어감 · 더 큰 규격을 선택하세요")
            else:
                st.caption("⚠️ 들어가는 지퍼백이 없습니다 · 규격을 확인하세요")

            if _pos and st.checkbox("봉투별 입수 상세 보기", value=False):
                st.dataframe(
                    pd.DataFrame([{"지퍼백": n, "규격": s,
                                   "입수(개)": c if c > 0 else "안 들어감",
                                   "선택": "✅" if n == _chosen["박스명"] else ""}
                                  for n, s, c in _caps]),
                    hide_index=True, use_container_width=True)

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
        safety_pct = st.number_input(
            "적재 여유율 (%)", min_value=0.0, max_value=90.0, value=0.0, step=5.0,
            help="이론 적재량은 실제보다 많을 수 있어요. 예: 10 → 이론 수량의 90%로 반영")

# ---------------------------------------------------------------------------
# 메인 - 계산 (제품 → 포장재 → 박스)
# ---------------------------------------------------------------------------
outer_boxes = BOX_CATALOG[outer_group]

# 트레이 칸수 (범용=자동 계산 / 제품 전용=직접 입력)
tray_cells, tray_grid = 0, (0, 0)
if is_tray and tray_l > 0 and tray_w > 0:
    if tray_custom:
        tray_cells = int(tray_cells_manual)
        # 직접 입력 칸수를 배치도용 격자로 분해 (정사각에 가까운 약수쌍 = 정확히 일치)
        _gc = 1
        for _a in range(1, int(tray_cells ** 0.5) + 1):
            if tray_cells % _a == 0:
                _gc = _a
        tray_grid = (_gc, tray_cells // _gc)
    else:
        tray_cells, tray_grid = tray_cell_count(
            product, {"inner_l": tray_l, "inner_w": tray_w},
            gap=tray_gap, pitch_x=tray_pitch_x, pitch_y=tray_pitch_y)

rows = build_packaging_rows(
    product, outer_boxes, inner_mode=inner_mode, outer_group=outer_group,
    unit_weight_g=unit_weight_g, part_name=part_name, wall_margin=wall_margin,
    use_best=use_best, tray_cells=tray_cells, tray_grid=tray_grid,
    tray_l=tray_l, tray_w=tray_w, tray_thickness=tray_thickness,
    bag_name=bag_name, bag_count=bag_count, bag_l=bag_l, bag_w=bag_w, bag_h=bag_h,
    weight_limit_kg=weight_limit_kg, safety_pct=safety_pct)

best_row = max(rows, key=lambda r: r["박스당 총 제품"]) if rows else None

# 사용할 박스 선택(탭 간 공유) — session_state('sel_box')로 유지
# (드롭다운/비교표 클릭 둘 다 이 값을 갱신 → 위젯 key 미사용, 수동 관리)
_rowmap = {r["박스명"]: r for r in rows}
if rows and st.session_state.get("sel_box") not in _rowmap:
    st.session_state["sel_box"] = best_row["박스명"]
sel_row = _rowmap.get(st.session_state.get("sel_box")) if rows else None
if rows and sel_row is None:
    sel_row = best_row


def pick_box(name):
    """박스 선택 갱신 (드롭다운·표 공용)."""
    if name and name != st.session_state.get("sel_box"):
        st.session_state["sel_box"] = name
        st.rerun()
# 트레이 제작정보(견적서용)
tray_info = ""
if is_tray and tray_custom:
    _parts = [x for x in [tray_material, f"두께 {tray_thickness:g}mm" if tray_thickness else "",
                          tray_maker, tray_color] if x]
    tray_info = " · ".join(_parts)

# ---------------------------------------------------------------------------
# 탭 구성
# ---------------------------------------------------------------------------
VIEWS = ["📊 적재 효율 계산", "📄 구매팀 견적 양식",
         "📚 기록 관리", "🗄️ 기준 데이터(박스 리스트)"]
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

        # 사용할 박스 선택 (드롭다운 · 견적/기록 탭과 공유)
        _names = [r["박스명"] for r in
                  sorted(rows, key=lambda r: r["박스당 총 제품"], reverse=True)]
        _picked = st.selectbox(
            "📦 사용할 박스 선택 (추천 자동선택 · 변경 가능)", _names,
            index=_names.index(sel_row["박스명"]),
            format_func=lambda n: f'{n}  ·  {_rowmap[n]["박스당 총 제품"]:,}개  ·  '
                                  f'{_rowmap[n]["규격(Size)"]}'
                                  + ("  🏆추천" if n == best_row["박스명"] else ""),
            help="추천 박스가 기본 선택됩니다. 드롭다운 또는 아래 비교표에서 골라도 됩니다.")
        pick_box(_picked)
        if sel_row is not best_row:
            st.caption(f"ℹ️ 추천은 **{best_row['박스명']}**({best_row['박스당 총 제품']:,}개)"
                       f"이지만 **{sel_row['박스명']}**({sel_row['박스당 총 제품']:,}개)를 "
                       "선택했습니다.")

        _c, _r = int(sel_row["_cols"]), int(sel_row["_rows"])
        _lay, _u = int(sel_row["_layers"]), sel_row["_unit"]
        _total = sel_row["박스당 총 제품"]
        _per_layer = _c * _r
        _geom = _per_layer * _lay                       # 부피 기준 최대
        _capped = _per_layer > 0 and _total < _geom     # 무게 한도 등으로 줄었는지
        _eff_layers = -(-_total // _per_layer) if _per_layer > 0 else 0
        # 선택 박스 실제 치수(mm)
        _pbox = next((b for b in outer_boxes if b["박스명"] == sel_row["박스명"]), None)
        _bl, _bw, _bh = (_pbox["inner_l"], _pbox["inner_w"], _pbox["inner_h"]) \
            if _pbox else (0, 0, 0)
        _layer_h = (_bh / _lay) if _lay else 0

        # 적재 배치도 (좌: Plotly 시각화 / 우: 수치 카드)
        section(f"적재 배치도 · {sel_row['박스명']}")
        st.markdown(
            '<div class="badges">'
            f'<span class="badge">가로 <b>{_bl:g}</b></span>'
            f'<span class="badge">세로 <b>{_bw:g}</b></span>'
            f'<span class="badge g">높이 <b>{_bh:g}</b></span>'
            '<span class="badge unit">mm</span>'
            '</div>', unsafe_allow_html=True)
        # 업로드한 실제 제품 3D 형상 미리보기 (있을 때만)
        _mg = st.session_state.get("_mesh_geom")
        if _mg and _mg.get("verts"):
            with st.expander(f"🧩 업로드한 실제 제품 3D 보기 · "
                             f"{st.session_state.get('_mesh_name','')}"):
                mp, mc = st.columns([2, 1], gap="large")
                with mp:
                    st.plotly_chart(mesh_preview_fig(_mg),
                                    use_container_width=True,
                                    config={"displayModeBar": False})
                _L, _W, _H = _mg["dims"]
                _vol_cc = round(_mg["volume"] / 1000, 1) if _mg.get("volume") else 0
                with mc:
                    st.markdown(
                        stat_card("측정 L×W×H", f"{_L}×{_W}×{_H}", "mm", hi=True)
                        + stat_card("부피", f"{_vol_cc:g}" if _vol_cc else "-",
                                    "cm³" if _vol_cc else ""),
                        unsafe_allow_html=True)
                    st.caption("제품 좌표축 기준(NX 측정값과 동일) · 3D 파일에서 자동 측정")
        viz, stt = st.columns([2, 1], gap="large")
        with viz:
            # 슬라이더를 3D 차트와 같은 컬럼에 묶고, 3D가 실제로 그리는 층수까지만 허용
            # (그림·슬라이더 최대치를 일치시켜 '슬라이더는 24인데 그림은 8' 혼란 제거)
            _ZCAP = 40                               # 3D가 표시하는 최대 층수(packing_fig_3d 와 동일)
            _disp_layers = min(_eff_layers, _ZCAP)   # 그림에 그려지는 실제 층수
            sel = _disp_layers
            if _c > 0 and _r > 0 and _disp_layers > 1:
                sel = st.slider("적층 선택 (1층 ~ N층 누적)", 1, _disp_layers, _disp_layers,
                                help="1층부터 선택한 층까지 채워 표시 · 맨 위 층 초록 강조 · "
                                     "드래그 회전 / 스크롤 확대")
            if _c > 0 and _r > 0:
                cum_qty = min(_per_layer * sel, _total)
                cum_h = round(sel * _layer_h)
                # 선택한 맨 위(sel) 층에 실제로 놓이는 블록 수 (자투리 층이면 < 층당개수)
                _top_cnt = cum_qty - _per_layer * (sel - 1)
                st.plotly_chart(
                    packing_fig_3d(_c, _r, _disp_layers, active_layers=sel,
                                   top_count=_top_cnt,
                                   box_l=_bl, box_w=_bw, box_h=_bh),
                    use_container_width=True, config={"displayModeBar": False})
                _grid_trunc = _c > 8 or _r > 8      # 바닥 격자만 8×8로 축약
                _layer_trunc = _eff_layers > _ZCAP  # 층수가 상한 초과 시에만
                _partial = _top_cnt < _per_layer
                cap_note = (f"  ·  ⚠️ 부피상 {_geom:,}개까지 가능하지만 "
                            f"'{sel_row['제한 요인']}'으로 {_total:,}개에서 적재를 멈춥니다") if _capped else ""
                partial_note = (f"  ·  맨 위 {sel}층은 자투리 <b>{_top_cnt}개</b>만 적재"
                                if _partial else "")
                grid_note = "  ·  바닥 격자가 커서 그림은 8×8로 축약" if _grid_trunc else ""
                layer_note = (f"  ·  {_eff_layers}층 중 {_ZCAP}층까지만 그림 표시"
                              if _layer_trunc else "")
                st.markdown(
                    f'<div style="font-size:12.5px;color:#8b98a5;margin-top:2px;">'
                    f'🧊 {_c}×{_r} 격자 × {_eff_layers}층'
                    f'{partial_note}{grid_note}{layer_note}{cap_note}</div>',
                    unsafe_allow_html=True)
                st.markdown(
                    '<div class="cta" style="margin-top:4px;padding:13px 18px;">'
                    f'<div><div class="t">📐 1층 ~ {sel}층 누적</div>'
                    f'<div class="d">예상 높이 <b style="color:#7ec8f3;">{cum_h:,} mm</b>'
                    f' / 박스 {_bh:g}mm &nbsp;·&nbsp; 누적 수량 '
                    f'<b style="color:#40d6a0;">{cum_qty:,} 개</b></div></div></div>',
                    unsafe_allow_html=True)
            else:
                st.info("이 조합은 적재되지 않습니다.")
        with stt:
            # 상단 '추천 박스'·배지줄과 중복되는 카드는 빼고, 배치도를 설명하는 2개만
            st.markdown(
                stat_card(f"층당 개수 ({_c}×{_r})", f"{_per_layer:,}", _u, hi=True)
                + stat_card("총 적층 단수", f"{_eff_layers:,}", "층"),
                unsafe_allow_html=True)
        st.markdown("")

        # 🚚 물류 적재 (파렛트 · 컨테이너) — 보조 정보라 기본 접힘
        if _total > 0 and _bl and _bw and _bh:
            with st.expander("🚚 물류 적재 (파렛트 · 컨테이너)"):
                lc1, lc2, lc3 = st.columns(3)
                _pal_name = lc1.selectbox("파렛트 규격", list(PALLETS.keys()))
                _con_name = lc2.selectbox("컨테이너", list(CONTAINERS.keys()))
                _pmargin = lc3.number_input("적재 높이 여유 (mm)", min_value=0.0, value=0.0,
                                            step=10.0, help="파렛트 적재 시 상단 여유 높이")
                _pal, _con = PALLETS[_pal_name], CONTAINERS[_con_name]
                _bpp, _pbase, _players = boxes_per_pallet(_bl, _bw, _bh, _pal, box_margin=_pmargin)
                _bpc, _cgrid = boxes_per_container(_bl, _bw, _bh, _con)
                _boxcbm = cbm(_bl, _bw, _bh)
                _fill = (_bpc * _boxcbm / cbm(_con["l"], _con["w"], _con["h"]) * 100) \
                    if _bpc else 0
                kpi_row([
                    {"label": f"파렛트당 박스 · {_pal_name.split(' ')[0]}", "value": f"{_bpp:,}",
                     "unit": "박스", "sub": f"바닥 {_pbase} × {_players}단", "w": 1},
                    {"label": "파렛트당 총 제품", "value": f"{_bpp * _total:,}", "unit": "개",
                     "op": "→", "variant": "total", "w": 1.2},
                    # 아래 컨테이너 행(3칸)과 열을 맞추기 위한 빈 칸
                    {"label": "", "value": "", "variant": "ghost", "w": 0.8},
                ])
                kpi_row([
                    {"label": f"{_con_name} 당 박스", "value": f"{_bpc:,}", "unit": "박스",
                     "sub": f"3D 배열 {_cgrid[0]}×{_cgrid[1]}×{_cgrid[2]} · 박스 {_boxcbm:.3f} CBM",
                     "w": 1},
                    {"label": f"{_con_name} 당 총 제품", "value": f"{_bpc * _total:,}", "unit": "개",
                     "op": "→", "variant": "total", "w": 1.2},
                    {"label": "컨테이너 적입률", "value": f"{_fill:.0f}", "unit": "%",
                     "sub": "박스 부피 합 ÷ 컨테이너 부피", "w": 0.8},
                ])
                st.caption("※ 파렛트/컨테이너 표준 규격 기준. 실제 규격은 `data.py`에서 조정 가능합니다.")

        # 💰 원가 · 중량 (선택)
        with st.expander("💰 원가 · 출하중량 계산 (선택)"):
            oc1, oc2, oc3 = st.columns(3)
            box_price = oc1.number_input("박스 단가 (원)", min_value=0.0, value=0.0, step=50.0)
            pack_price = oc2.number_input("포장재 단가/박스 (원)", min_value=0.0, value=0.0, step=10.0)
            box_self_g = oc3.number_input("박스 자체 무게 (g)", min_value=0.0, value=0.0, step=10.0)
            _unit_cost = (box_price + pack_price) / _total if _total else 0
            _ship_kg = (unit_weight_g * _total + box_self_g) / 1000 \
                if (unit_weight_g or box_self_g) else 0
            kpi_row([
                {"label": "제품 1개당 포장원가", "value": f"{_unit_cost:,.1f}", "unit": "원",
                 "sub": "(박스+포장재 단가) ÷ 박스당 제품"},
                {"label": "박스 출하 총중량", "value": f"{_ship_kg:,.2f}", "unit": "kg",
                 "sub": "제품 무게 합 + 박스 자체 무게"},
            ])

        df = pd.DataFrame(rows).drop(columns=["_cols", "_rows", "_layers", "_unit"])
        _mx = df["박스당 총 제품"].max()
        df.insert(0, "선택", df["박스명"].apply(
            lambda n: "✅" if n == sel_row["박스명"] else ""))
        df.insert(1, "추천", df["박스당 총 제품"].apply(
            lambda v: "🏆" if v == _mx and v > 0 else ""))
        show_cols = ["선택", "추천", "박스명", "규격(Size)", "포장재", "적재 방식",
                     "박스당 총 제품", "제한 요인", "비고"]
        if unit_weight_g:
            show_cols.insert(7, "박스 총중량(kg)")
        with st.expander(f"📦 다른 박스 규격 및 적재량 비교하기  ({len(rows)}종) · "
                         "행을 클릭하면 그 박스로 선택됩니다 ▾", expanded=False):
            _ev = st.dataframe(
                df[show_cols], use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row", key="cmp_table",
                column_config={
                    "박스당 총 제품": st.column_config.NumberColumn(format="%d 개"),
                },
            )
            if _ev.selection.rows:
                _r0 = _ev.selection.rows[0]
                # 표 클릭이 '새로 바뀐' 경우에만 적용 (드롭다운 선택과 충돌 방지)
                if st.session_state.get("_last_tbl_row") != _r0:
                    st.session_state["_last_tbl_row"] = _r0
                    pick_box(df.iloc[_r0]["박스명"])

            # 박스별 적재량 비교 차트 (선택 박스=에메랄드, 나머지=차콜)
            _bars = df[df["박스당 총 제품"] > 0]
            if not _bars.empty:
                _sel_name = sel_row["박스명"]
                _colors = ["#40d6a0" if n == _sel_name else "#39424f"
                           for n in _bars["박스명"]]
                _cfig = go.Figure(go.Bar(
                    x=_bars["박스명"], y=_bars["박스당 총 제품"],
                    marker=dict(color=_colors, cornerradius=5, line=dict(width=0)),
                    text=_bars["박스당 총 제품"], texttemplate="%{text:,}",
                    textposition="outside", textfont=dict(size=10, color="#aeb9c5"),
                    cliponaxis=False,
                    hovertemplate="%{x}<br>%{y:,}개<extra></extra>"))
                _cfig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    height=280, margin=dict(l=0, r=0, t=22, b=0), bargap=0.4,
                    font=dict(color="#8b98a5", size=11, family="Inter, sans-serif"),
                    uniformtext=dict(minsize=8, mode="hide"),
                    yaxis=dict(visible=False, showgrid=False, zeroline=False),
                    xaxis=dict(type="category", showgrid=False, zeroline=False,
                               showline=False, ticks="", color="#8b98a5",
                               tickfont=dict(size=10.5)))
                st.plotly_chart(_cfig, use_container_width=True,
                                config={"displayModeBar": False})

        # 기록 저장 (추천 vs 실제 적용)
        st.markdown("")
        section("이 사양 기록 저장  ·  추천 ↔ 실제 적용")
        _hist = store.load_df()
        _sim = store.similar_records(_hist, product)
        if _sim is not None and not _sim.empty:
            st.info(f"💡 비슷한 사이즈 과거 기록 **{len(_sim)}건**이 있어요. "
                    "'📚 기록 관리'에서 실제 적용 이력을 확인하세요.")
        with st.form("record_form", border=True):
            st.caption("대시보드 추천값이 기본으로 채워집니다. 현장에서 실제 적용한 값이 "
                       "다르면 수정 후 저장하세요.")
            rc1, rc2, rc3 = st.columns(3, gap="large")
            real_inner = rc1.text_input("실제 포장재", value=inner_mode)
            real_box = rc2.text_input("실제 적용 박스", value=sel_row["박스명"])
            real_qty = rc3.number_input("실제 적용 수량", min_value=0,
                                        value=int(sel_row["박스당 총 제품"]), step=1)
            memo = st.text_area("현실화 메모 / 비고", value="", height=84,
                                placeholder="예: 추천은 55-2였으나 재고 문제로 T-10 적용")
            _saved = st.form_submit_button("💾 기록 저장", type="primary",
                                           use_container_width=True)
        if _saved:
            if not (customer or part_name):
                st.warning("고객사 또는 품명을 입력한 뒤 저장하세요.")
            else:
                match = "✅ 동일" if (real_box == best_row["박스명"]
                                    and real_inner == inner_mode) else "⚠️ 변경"
                store.append_record({
                    "저장일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "고객사": customer or "-", "제품명": part_name or "-",
                    "L": pl, "W": pw, "H": ph, "무게(g)": unit_weight_g,
                    "추천 포장재": inner_mode, "추천 박스": best_row["박스명"],
                    "추천 수량": best_row["박스당 총 제품"],
                    "실제 포장재": real_inner, "실제 박스": real_box,
                    "실제 수량": real_qty, "일치": match,
                    "트레이재질": tray_material, "트레이두께": tray_thickness if is_tray else "",
                    "제작업체": tray_maker, "색상": tray_color, "비고": memo,
                })
                st.success(f"✅ 기록 저장됨 ({match}). '📚 기록 관리'에서 확인하세요.")

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
    _sb = sel_row["박스명"] if sel_row else ""
    _ss = sel_row["규격(Size)"] if sel_row else ""
    _sq = sel_row["박스당 총 제품"] if sel_row else 0
    header_info = default_header_info(
        customer, outer_group, product, part_name=part_name,
        unit_weight_g=unit_weight_g, inner_mode=inner_mode, bag_name=bag_name,
        sel_box=_sb, sel_size=_ss, sel_qty=_sq, tray_info=tray_info)

    # 메타 정보를 칩으로
    chips = "".join(f'<span class="chip"><b>{k}</b> {v}</span>'
                    for k, v in header_info.items())
    st.markdown(f'<div class="ctxbar">{chips}</div>', unsafe_allow_html=True)

    if sel_row:
        st.success(f"📦 선택 박스 **{_sb}** ({_ss}) · 박스당 **{_sq:,}개** 기준으로 견적을 만듭니다.")
    scope = st.radio("견적 범위", ["선택 박스만", "전체 박스 비교"], horizontal=True,
                     help="기본은 선택한 박스 1종. 여러 박스 단가를 한 번에 받으려면 '전체'.")
    quote_source = [sel_row] if (scope == "선택 박스만" and sel_row) else rows

    st.markdown("##### 견적 항목  ·  구매팀이 **‘구매 확정 단가’** 열에 입력해 회신")
    if not quote_source:
        st.warning("표시할 항목이 없습니다.")
    else:
        editable = pd.DataFrame(quote_source).drop(
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

# --- 탭3: 기록 관리 ---
elif view == VIEWS[2]:
    section("📚 기록 관리")
    hist = store.load_df()

    if hist.empty:
        st.info("아직 저장된 기록이 없습니다. '📊 적재 효율 계산'에서 "
                "'💾 기록 저장'을 눌러 첫 기록을 남겨보세요.")
    else:
        # 추천 vs 실제 적용 요약
        if "일치" in hist.columns:
            _same = (hist["일치"].astype(str).str.contains("동일")).sum()
            _chg = len(hist) - _same
            mm = st.columns(3)
            mm[0].metric("총 기록", f"{len(hist)}건")
            mm[1].metric("추천대로 적용", f"{_same}건")
            mm[2].metric("변경 적용", f"{_chg}건")

        # 비슷한 사이즈 조회
        sim = store.similar_records(hist, product)
        section(f"🔍 현재 제품({pl:g}×{pw:g}×{ph:g})과 비슷한 사이즈")
        if sim is not None and not sim.empty:
            st.caption(f"±15%(최소 5mm) 이내 재원 기록 {len(sim)}건 · 추천 ↔ 실제 적용 비교")
            st.dataframe(sim, use_container_width=True, hide_index=True)
        else:
            st.caption("비슷한 사이즈의 과거 기록이 없습니다.")

        # 추천 ↔ 실제 비교 (핵심 열만)
        section("추천 ↔ 실제 적용 비교")
        cmp_cols = ["저장일시", "고객사", "제품명", "추천 포장재", "추천 박스",
                    "추천 수량", "실제 포장재", "실제 박스", "실제 수량", "일치"]
        cmp_cols = [c for c in cmp_cols if c in hist.columns]
        st.dataframe(hist[cmp_cols][::-1], use_container_width=True, hide_index=True)

        section("전체 기록 (제작정보 포함)")
        kw = st.text_input("검색 (고객사·제품명·박스·업체)", value="",
                           placeholder="비워두면 전체 표시")
        shown = hist
        if kw:
            m = hist.apply(lambda r: kw.lower() in
                           " ".join(str(v) for v in r.values).lower(), axis=1)
            shown = hist[m]
        st.dataframe(shown[::-1], use_container_width=True, hide_index=True)
        st.caption(f"총 {len(hist)}건 저장됨 · {len(shown)}건 표시")

    st.markdown("")
    section("보관 · 복원 (영구 저장)")
    st.caption("⚠️ 클라우드는 재배포 시 기록이 초기화됩니다. 주기적으로 내려받아 보관하고, "
               "필요 시 다시 올려 복원하세요.")
    b1, b2 = st.columns(2)
    b1.download_button("⬇️ 전체 기록 내려받기 (CSV)",
                       data=hist.to_csv(index=False).encode("utf-8-sig"),
                       file_name="포장기록.csv", mime="text/csv",
                       use_container_width=True, disabled=hist.empty)
    up = b2.file_uploader("⬆️ CSV 올려 복원/병합", type=["csv"],
                          label_visibility="collapsed")
    if up is not None:
        try:
            new = pd.read_csv(up)
            merged = pd.concat([hist, new], ignore_index=True).drop_duplicates()
            store.save_df(merged)
            st.success(f"복원 완료 · 현재 {len(merged)}건")
        except Exception as e:
            st.error(f"CSV를 읽을 수 없습니다: {e}")
    if not hist.empty and st.button("🗑️ 전체 기록 삭제", type="secondary"):
        store.save_df(hist.iloc[0:0])
        st.warning("모든 기록을 삭제했습니다.")

# --- 탭4: 기준 데이터 열람 ---
else:
    section("표준 포장재 카탈로그 (기준 데이터)")
    st.caption("실제 값 수정은 `data.py`의 리스트(CARTONS·DANPLA·PLASTIC·ZIPPERS·TRAYS)에서 관리합니다.")
    _grp = st.selectbox("분류 선택", list(BOX_CATALOG.keys()))
    st.dataframe(pd.DataFrame(BOX_CATALOG[_grp]), use_container_width=True, hide_index=True)

st.divider()
st.caption("© 성우 개발팀 · 포장 사양 견적 대시보드")
