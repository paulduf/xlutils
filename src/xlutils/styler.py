from __future__ import annotations

from typing import Any, Literal

from .models import Theme
from .normalizer import normalize
from .sheet_spec import SheetSpec
from .themes import THEME_REGISTRY
from .writer import SheetConfig, write


class Styler:
    """Chainable Excel styling object. Accumulates config, writes on `.save()`.

    **Single-sheet** (unchanged existing API):

    ```python
    (style(data)
         .apply_theme("minimal")
         .expand_columns("fit")
         .color_gradient("Sales")
         .save("report.xlsx"))
    ```

    **Multi-sheet**: call `style()` with no data, then add sheets with
    `.add_sheet()`. Each sheet is a `SheetSpec` with the same styling methods.
    Call `.back()` on a `SheetSpec` to return here for the next sheet.

    ```python
    (style()
        .add_sheet(sales_df, "Sales")
            .apply_theme("default")
            .color_gradient("Revenue")
            .back()
        .add_sheet(costs_df, "Costs")
            .apply_theme("minimal")
            .freeze_header(False)
            .back()
        .save("report.xlsx"))
    ```

    Calling styling methods directly on `Styler` in multi-sheet mode sets
    **workbook-level defaults** that all sheets inherit unless overridden.
    """

    def __init__(self, data: Any = None) -> None:
        self._sheets: list[SheetSpec] = []
        self._primary: bool = False
        self._default_theme: Theme = THEME_REGISTRY["default"]
        self._default_column_mode: str | int | None = "fit"
        self._default_freeze_header: bool = True

        if data is not None:
            self._primary = True
            sheet = SheetSpec(data, "Sheet1", self)
            self._sheets.append(sheet)

    def _resolve_theme(self, theme: str | Theme) -> Theme:
        if isinstance(theme, Theme):
            return theme
        if theme not in THEME_REGISTRY:
            available = ", ".join(f'"{k}"' for k in THEME_REGISTRY)
            raise ValueError(f"Unknown theme {theme!r}. Available: {available}")
        return THEME_REGISTRY[theme]

    # -- Styling methods (workbook defaults in multi-sheet mode, proxied to sheet in primary mode) --

    def apply_theme(self, theme: str | Theme) -> Styler:
        """Set the theme.

        In single-sheet mode, applies to the sheet. In multi-sheet mode, sets a
        workbook-level default inherited by all sheets unless overridden per-sheet.
        Pass a name (`"default"`, `"minimal"`) or a `Theme` object.
        """
        resolved = self._resolve_theme(theme)
        if self._primary:
            self._sheets[0]._theme = resolved
        else:
            self._default_theme = resolved
        return self

    def expand_columns(self, mode: Literal["fit"] | int | None = "fit") -> Styler:
        """Set column width.

        In single-sheet mode, applies to the sheet. In multi-sheet mode, sets a
        workbook-level default. `"fit"` auto-sizes, int for fixed width, `None` to disable.
        """
        if self._primary:
            self._sheets[0]._column_mode = mode
        else:
            self._default_column_mode = mode
        return self

    def freeze_header(self, freeze: bool = True) -> Styler:
        """Freeze the top row.

        In single-sheet mode, applies to the sheet. In multi-sheet mode, sets a
        workbook-level default inherited by all sheets unless overridden per-sheet.
        """
        if self._primary:
            self._sheets[0]._freeze_header = freeze
        else:
            self._default_freeze_header = freeze
        return self

    def color_gradient(
        self,
        column: str | int,
        low_color: str = "#FFFFFF",
        high_color: str = "#4472C4",
    ) -> Styler:
        """Apply a two-color gradient to a column.

        Only available in single-sheet mode. In multi-sheet mode, call
        `.color_gradient()` on the individual `SheetSpec` returned by `.add_sheet()`.
        """
        if not self._primary:
            raise TypeError(
                "color_gradient() cannot be called on Styler in multi-sheet mode. "
                "Call it on the SheetSpec returned by add_sheet() instead."
            )
        self._sheets[0].color_gradient(column, low_color, high_color)
        return self

    # -- Sheet management --

    def add_sheet(self, data: Any, sheet_name: str = "Sheet1") -> SheetSpec:
        """Add a sheet to the workbook and return its `SheetSpec` for per-sheet configuration.

        Args:
            data: Table data for this sheet (list[dict], pandas DataFrame, polars DataFrame).
            sheet_name: Name of the worksheet tab.

        Returns:
            A `SheetSpec` for chaining per-sheet styling. Call `.back()` on it to
            return to this `Styler`, or `.save(path)` to write the workbook directly.

        Raises:
            TypeError: If called on a `Styler` created with `style(data)` (primary mode).
            ValueError: If `sheet_name` is already used by another sheet.
        """
        if self._primary:
            raise TypeError(
                "Cannot mix style(data) and add_sheet(). "
                "Use style() with no data for multi-sheet workbooks."
            )
        if any(s._sheet_name == sheet_name for s in self._sheets):
            raise ValueError(f"Duplicate sheet name: {sheet_name!r}")
        sheet = SheetSpec(data, sheet_name, self)
        self._sheets.append(sheet)
        return sheet

    def _defaults(self) -> dict:
        return {
            "theme": self._default_theme,
            "column_mode": self._default_column_mode,
            "freeze_header": self._default_freeze_header,
        }

    def save(self, path: str, sheet_name: str | None = None) -> None:
        """Write the styled workbook to an `.xlsx` file.

        Args:
            path: Destination file path (must end in `.xlsx`).
            sheet_name: Worksheet name, only valid for single-sheet workbooks.
                Ignored in multi-sheet mode (use the `sheet_name` arg of `add_sheet()`).

        Raises:
            ValueError: If there are no sheets to write.
            ValueError: If `sheet_name` is given but there are multiple sheets.
        """
        if not self._sheets:
            raise ValueError(
                "Cannot save a workbook with no sheets. Call add_sheet() at least once."
            )
        if sheet_name is not None and len(self._sheets) > 1:
            raise ValueError(
                "sheet_name kwarg is only valid for single-sheet workbooks. "
                "Use the sheet_name argument of add_sheet() to name sheets."
            )
        if sheet_name is not None:
            self._sheets[0]._sheet_name = sheet_name

        defaults = self._defaults()
        sheet_configs: list[SheetConfig] = []
        for sheet in self._sheets:
            headers, rows = normalize(sheet._data)
            resolved = sheet._resolve(defaults)
            sheet_configs.append(
                SheetConfig(
                    sheet_name=sheet._sheet_name,
                    headers=headers,
                    rows=rows,
                    theme=resolved["theme"],
                    column_mode=resolved["column_mode"],
                    gradient_rules=resolved["gradient_rules"],
                    freeze_header=resolved["freeze_header"],
                )
            )
        write(path, sheet_configs)
