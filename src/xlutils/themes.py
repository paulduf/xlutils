from .models import BorderSide, BorderStyle, CellStyle, FillStyle, FontStyle, Theme

_DEFAULT = Theme(
    name="default",
    header=CellStyle(
        font=FontStyle(bold=True, color="FFFFFF", size=11),
        fill=FillStyle(fg_color="1F3864"),
        border=BorderStyle(bottom=BorderSide(style="thin", color="FFFFFF")),
        horizontal_align="center",
    ),
    even_row=CellStyle(
        fill=FillStyle(fg_color="DCE6F1"),
    ),
    odd_row=CellStyle(
        fill=FillStyle(fg_color="FFFFFF"),
    ),
)

_MINIMAL = Theme(
    name="minimal",
    header=CellStyle(
        font=FontStyle(bold=True, color="1F1F1F", size=11),
        fill=FillStyle(fg_color=None),
        border=BorderStyle(bottom=BorderSide(style="medium", color="1F1F1F")),
        horizontal_align="left",
    ),
    even_row=CellStyle(
        fill=FillStyle(fg_color=None),
    ),
    odd_row=CellStyle(
        fill=FillStyle(fg_color="F5F5F5"),
    ),
)

THEME_REGISTRY: dict[str, Theme] = {
    "default": _DEFAULT,
    "minimal": _MINIMAL,
}
