from __future__ import annotations

from pydantic import BaseModel


class FontStyle(BaseModel):
    name: str = "Calibri"
    size: int = 11
    bold: bool = False
    italic: bool = False
    color: str = "000000"  # hex, no leading #


class FillStyle(BaseModel):
    fg_color: str | None = None  # hex, no leading #
    pattern: str = "solid"


class BorderSide(BaseModel):
    style: str | None = None  # "thin", "medium", "thick", None
    color: str = "000000"


class BorderStyle(BaseModel):
    top: BorderSide = BorderSide()
    bottom: BorderSide = BorderSide()
    left: BorderSide = BorderSide()
    right: BorderSide = BorderSide()


class CellStyle(BaseModel):
    font: FontStyle = FontStyle()
    fill: FillStyle = FillStyle()
    border: BorderStyle = BorderStyle()
    wrap_text: bool = False
    horizontal_align: str = "left"


class Theme(BaseModel):
    name: str
    header: CellStyle
    even_row: CellStyle
    odd_row: CellStyle
    column_padding: int = 2
