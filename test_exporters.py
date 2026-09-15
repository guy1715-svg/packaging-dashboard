"""
exporters.py 회귀 테스트
===========================
Excel/PDF 견적서 생성이 예외 없이 유효한 파일 바이트를 만들어내는지,
빈 결과(적재 불가) 케이스도 안전하게 처리하는지 검증한다.

실행:  pytest -q          (프로젝트 루트에서)
"""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook

from exporters import to_excel_bytes, to_pdf_bytes, default_header_info

HEADER = default_header_info("성우전자", "박스(Carton)", (50, 30, 15),
                             part_name="테스트부품", unit_weight_g=12.5,
                             inner_mode="없음(벌크)", sel_box="55-2",
                             sel_size="410*345*412", sel_qty=100)

ROWS = [{
    "품명": "테스트부품", "박스종류": "박스(Carton)", "박스명": "55-2",
    "규격(Size)": "410*345*412", "포장재": "없음(벌크)",
    "적재 방식": "벌크 3D 8×10×20", "박스당 총 제품": 1600,
    "제한 요인": "-", "박스 총중량(kg)": 20.0, "비고": "",
    "구매 확정 단가": None,
}]


def test_default_header_info_has_expected_keys():
    assert HEADER["고객사"] == "성우전자"
    assert HEADER["제품 사이즈(L×W×H mm)"] == "50×30×15"
    assert HEADER["선택 박스"] == "55-2 (410*345*412)"


def test_to_excel_bytes_basic():
    data = to_excel_bytes(HEADER, ROWS)
    assert isinstance(data, bytes)
    assert len(data) > 0
    wb = load_workbook(io.BytesIO(data))
    ws = wb.active
    assert ws["A1"].value == "포장자재 견적 요청서 (Packaging Quotation Request)"


def test_to_excel_bytes_empty_rows_does_not_crash():
    data = to_excel_bytes(HEADER, [])
    assert isinstance(data, bytes)
    wb = load_workbook(io.BytesIO(data))
    ws = wb.active
    values = [c.value for row in ws.iter_rows() for c in row if c.value]
    assert any("적재 가능한 박스가 없습니다" in str(v) for v in values)


def test_to_excel_bytes_bilingual_headers():
    data = to_excel_bytes(HEADER, ROWS, bilingual=True)
    wb = load_workbook(io.BytesIO(data))
    ws = wb.active
    header_cells = [c.value for row in ws.iter_rows() for c in row if c.value]
    assert any("Tên SP" in str(v) for v in header_cells)


def test_to_pdf_bytes_basic():
    data = to_pdf_bytes(HEADER, ROWS)
    assert isinstance(data, bytes)
    assert data.startswith(b"%PDF")


def test_to_pdf_bytes_empty_rows_does_not_crash():
    data = to_pdf_bytes(HEADER, [])
    assert isinstance(data, bytes)
    assert data.startswith(b"%PDF")
