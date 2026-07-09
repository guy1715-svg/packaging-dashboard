"""
3D 파일(STL/OBJ/PLY) → 제품 치수·형상 추출
================================================

NX 등 CAD에서 내보낸 3D 메쉬 파일을 읽어
  ① 최소 바운딩박스 기준 L/W/H (제품이 비스듬히 놓여 있어도 가장 타이트한 치수)
  ② 3D 미리보기용 정점/면 데이터
를 반환합니다.

STEP(.stp)은 B-rep(곡면 수식) 포맷이라 CAD 커널이 필요해 지원하지 않습니다.
NX에서 '파일 → 내보내기 → STL' 로 저장해 업로드하세요.

trimesh 미설치 시에도 앱이 죽지 않도록 안전하게 폴백합니다.
"""

import io

SUPPORTED = ["stl", "obj", "ply", "glb", "off"]
MAX_PREVIEW_FACES = 40000   # 미리보기 성능 상한 (초과 시 데시메이션 시도)


def available():
    """trimesh 사용 가능 여부."""
    try:
        import trimesh  # noqa: F401
        return True
    except Exception:
        return False


def _to_single_mesh(loaded):
    """trimesh load 결과(Scene/Mesh)를 하나의 Trimesh로 병합."""
    import trimesh
    if isinstance(loaded, trimesh.Trimesh):
        return loaded
    if isinstance(loaded, trimesh.Scene):
        geoms = [g for g in loaded.geometry.values()
                 if isinstance(g, trimesh.Trimesh) and len(g.faces)]
        if not geoms:
            return None
        return trimesh.util.concatenate(geoms)
    return None


def load_mesh(data_bytes, filename):
    """
    업로드 바이트 + 파일명 → 결과 dict 또는 None.

    반환 dict:
      dims   : (L, W, H) 내림차순 정렬 정수 mm  ← 최소 바운딩박스 기준
      aabb   : (x, y, z) 축정렬 바운딩박스 (참고용)
      volume : 부피(mm³, 닫힌 메쉬만 유효, 아니면 0)
      verts  : [[x,y,z], ...]  (미리보기용, 원점 기준 정렬)
      faces  : [[i,j,k], ...]
      note   : 처리 관련 안내 문구
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in SUPPORTED:
        return None
    try:
        import trimesh
    except Exception:
        return None

    try:
        loaded = trimesh.load(io.BytesIO(data_bytes), file_type=ext, force="mesh")
    except Exception:
        try:
            loaded = trimesh.load(io.BytesIO(data_bytes), file_type=ext)
        except Exception:
            return None

    mesh = _to_single_mesh(loaded)
    if mesh is None or len(mesh.faces) == 0:
        return None

    note = ""

    # --- 치수: 제품 좌표축 기준 바운딩박스(AABB) = NX/CAD 측정값과 동일 ---
    # 최소회전박스(OBB)는 부피는 작지만 제품을 기울인 값이라 실제 포장 방향과
    # 어긋나고(특히 높이 과소평가) NX 측정값과 불일치 → 사용하지 않는다.
    dims = tuple(sorted((round(float(x), 1) for x in mesh.extents), reverse=True))
    aabb = dims

    try:
        volume = float(mesh.volume) if mesh.is_watertight else 0.0
    except Exception:
        volume = 0.0

    # --- 미리보기용 지오메트리 (원점 기준 이동, 필요시 데시메이션) ---
    prev = mesh
    if len(prev.faces) > MAX_PREVIEW_FACES:
        try:
            prev = prev.simplify_quadric_decimation(MAX_PREVIEW_FACES)
            note = (note + " ").strip() + f"미리보기 면수를 {MAX_PREVIEW_FACES:,}개로 단순화했습니다."
        except Exception:
            pass

    v = prev.vertices - prev.vertices.min(axis=0)   # 원점으로 이동
    verts = v.tolist()
    faces = prev.faces.tolist()

    return {
        "dims": dims, "aabb": aabb, "volume": volume,
        "verts": verts, "faces": faces, "note": note,
    }
