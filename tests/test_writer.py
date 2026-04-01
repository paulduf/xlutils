import os
import tempfile

from openpyxl import load_workbook

from xlutils import style

DATA = [{"product": "Alpha", "revenue": 100}, {"product": "Beta", "revenue": 200}]


def _save_and_load(styler, **kwargs):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        styler.save(path, **kwargs)
        wb = load_workbook(path)
        return wb
    finally:
        os.unlink(path)


def test_one_liner_writes_file():
    wb = _save_and_load(style(DATA))
    ws = wb.active
    assert ws["A1"].value == "product"
    assert ws["B1"].value == "revenue"
    assert ws["A2"].value == "Alpha"
    assert ws["B3"].value == 200


def test_header_is_bold():
    wb = _save_and_load(style(DATA))
    ws = wb.active
    assert ws["A1"].font.bold is True


def test_banded_rows_have_different_fills():
    wb = _save_and_load(style(DATA))
    ws = wb.active
    fill_row2 = ws["A2"].fill.fgColor.rgb
    fill_row3 = ws["A3"].fill.fgColor.rgb
    assert fill_row2 != fill_row3


def test_minimal_theme():
    wb = _save_and_load(style(DATA).apply_theme("minimal"))
    ws = wb.active
    assert ws["A1"].font.bold is True


def test_custom_sheet_name():
    wb = _save_and_load(style(DATA), sheet_name="Sales")
    assert "Sales" in wb.sheetnames


def test_fixed_column_width():
    wb = _save_and_load(style(DATA).expand_columns(30))
    ws = wb.active
    assert ws.column_dimensions["A"].width == 30


def test_no_column_expand():
    # expand_columns(None) should leave widths at default
    wb = _save_and_load(style(DATA).expand_columns(None))
    ws = wb.active
    # openpyxl default is None/0; just verify no crash
    assert ws["A1"].value == "product"
