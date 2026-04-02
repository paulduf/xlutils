from .models import CellStyle, Theme
from .sheet_spec import SheetSpec
from .styler import Styler
from .themes import THEME_REGISTRY


def style(data=None) -> Styler:
    """Entry point. Returns a chainable `Styler` object.

    Pass data for a single-sheet workbook, or no data for a multi-sheet workbook.

    Examples
    --------
    Single sheet (one-liner):

    ```python
    style(df).save("report.xlsx")
    ```

    Single sheet (fluent):

    ```python
    (style(df)
         .apply_theme("minimal")
         .expand_columns("fit")
         .color_gradient("revenue")
         .save("report.xlsx"))
    ```

    Multi-sheet workbook:

    ```python
    (style()
        .add_sheet(sales_df, "Sales")
            .apply_theme("default")
            .color_gradient("Revenue")
            .back()
        .add_sheet(costs_df, "Costs")
            .apply_theme("minimal")
            .back()
        .save("report.xlsx"))
    ```
    """
    return Styler(data)


__all__ = ["style", "Styler", "SheetSpec", "Theme", "CellStyle", "THEME_REGISTRY"]
