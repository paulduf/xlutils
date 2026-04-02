from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal

from .models import Theme
from .themes import THEME_REGISTRY

if TYPE_CHECKING:
    from .styler import Styler


class _Unset:
    """Sentinel for settings not explicitly set on a SheetSpec."""


_UNSET = _Unset()


class SheetSpec:
    """Per-sheet configuration within a multi-sheet workbook.

    Returned by `Styler.add_sheet()`. Supports the same fluent styling methods
    as `Styler`, scoped to a single sheet. Call `.back()` to return to the
    parent `Styler` for workbook-level operations.

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

    Settings not explicitly set on a sheet inherit from the workbook-level
    defaults set on the parent `Styler`.
    """

    def __init__(self, data: Any, sheet_name: str, styler: Styler) -> None:
        self._data = data
        self._sheet_name = sheet_name
        self._styler = styler
        self._theme: Theme | _Unset = _UNSET
        self._column_mode: str | int | None | _Unset = _UNSET
        self._freeze_header: bool | _Unset = _UNSET
        self._gradient_rules: list[dict] = []

    def apply_theme(self, theme: str | Theme) -> SheetSpec:
        """Set the theme for this sheet. Pass a name or a `Theme` object."""
        if isinstance(theme, Theme):
            self._theme = theme
        else:
            if theme not in THEME_REGISTRY:
                available = ", ".join(f'"{k}"' for k in THEME_REGISTRY)
                raise ValueError(f"Unknown theme {theme!r}. Available: {available}")
            self._theme = THEME_REGISTRY[theme]
        return self

    def expand_columns(self, mode: Literal["fit"] | int | None = "fit") -> SheetSpec:
        """Set column width for this sheet. `"fit"` auto-sizes, int for fixed, `None` to disable."""
        self._column_mode = mode
        return self

    def freeze_header(self, freeze: bool = True) -> SheetSpec:
        """Freeze the top row for this sheet."""
        self._freeze_header = freeze
        return self

    def color_gradient(
        self,
        column: str | int,
        low_color: str = "#FFFFFF",
        high_color: str = "#4472C4",
    ) -> SheetSpec:
        """Apply a two-color gradient to a column of this sheet."""
        self._gradient_rules.append(
            {"column": column, "low": low_color, "high": high_color}
        )
        return self

    def back(self) -> Styler:
        """Return the parent `Styler` to continue workbook-level configuration."""
        return self._styler

    def save(self, path: str) -> None:
        """Write the workbook to an `.xlsx` file (delegates to parent `Styler`)."""
        self._styler.save(path)

    def _resolve(self, defaults: dict) -> dict:
        """Merge per-sheet settings over workbook-level defaults."""
        return {
            "theme": self._theme if not isinstance(self._theme, _Unset) else defaults["theme"],
            "column_mode": self._column_mode if not isinstance(self._column_mode, _Unset) else defaults["column_mode"],
            "freeze_header": self._freeze_header if not isinstance(self._freeze_header, _Unset) else defaults["freeze_header"],
            "gradient_rules": self._gradient_rules,
        }
