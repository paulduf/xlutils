from __future__ import annotations

from typing import Any, Literal

from .models import Theme
from .normalizer import normalize
from .themes import THEME_REGISTRY
from .writer import write


class Styler:
    """Chainable Excel styling object. Accumulates config, writes on `.save()`.

    All methods return `self`, enabling fluent chaining:

    ```python
    (style(data)
         .apply_theme("minimal")
         .expand_columns("fit")
         .color_gradient("Sales")
         .save("report.xlsx"))
    ```
    """

    def __init__(self, data: Any) -> None:
        self._data = data
        self._theme: Theme = THEME_REGISTRY["default"]
        self._column_mode: str | int | None = "fit"
        self._freeze_header: bool = True
        self._gradient_rules: list[dict] = []

    def apply_theme(self, theme: str | Theme) -> Styler:
        """Set the theme. Pass a name (`"default"`, `"minimal"`) or a `Theme` object."""
        self._theme = self._resolve_theme(theme)
        return self

    def _resolve_theme(self, theme: str | Theme) -> Theme:
        if isinstance(theme, Theme):
            return theme
        if theme not in THEME_REGISTRY:
            available = ", ".join(f'"{k}"' for k in THEME_REGISTRY)
            raise ValueError(f"Unknown theme {theme!r}. Available: {available}")
        return THEME_REGISTRY[theme]

    # -- Layout --

    def expand_columns(self, mode: Literal["fit"] | int = "fit") -> Styler:
        """Set column width. `"fit"` auto-sizes, or pass an int for fixed."""
        self._column_mode = mode
        return self

    def freeze_header(self, freeze: bool = True) -> Styler:
        """Freeze the top row so headers stay visible while scrolling."""
        self._freeze_header = freeze
        return self

    def color_gradient(
        self,
        column: str | int,
        low_color: str = "#FFFFFF",
        high_color: str = "#4472C4",
    ) -> Styler:
        """Apply a two-color gradient to a column. Chainable."""
        self._gradient_rules.append(
            {"column": column, "low": low_color, "high": high_color}
        )
        return self

    def save(self, path: str, sheet_name: str = "Sheet1") -> None:
        """Write the styled workbook to an `.xlsx` file."""
        headers, rows = normalize(self._data)
        write(
            path=path,
            headers=headers,
            rows=rows,
            theme=self._theme,
            column_mode=self._column_mode,
            gradient_rules=self._gradient_rules,
            freeze_header=self._freeze_header,
            sheet_name=sheet_name,
        )
