from .models import CellStyle, Theme
from .styler import Styler
from .themes import THEME_REGISTRY


def style(data) -> Styler:
    """Entry point. Returns a chainable Styler object.

    Examples
    --------
    # One-liner with defaults
    style(df).save("report.xlsx")

    # Fluent API
    (style(df)
         .apply_theme("minimal")
         .expand_columns("fit")
         .color_gradient("revenue")
         .save("report.xlsx"))
    """
    return Styler(data)


__all__ = ["style", "Styler", "Theme", "CellStyle", "THEME_REGISTRY"]
