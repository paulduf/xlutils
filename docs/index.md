# xlutils

Minimal boilerplate Excel styling for dataframes.

## Installation

```bash
uv add xlutils
```

## Quick Start

```python
from xlutils import style

data = [
    {"Region": "North", "Product": "Widget A", "Sales": 14200},
    {"Region": "South", "Product": "Widget B", "Sales": 6300},
]

# One-liner with nice defaults
style(data).save("report.xlsx")

# Fluent API for more control
(style(data)
     .apply_theme("minimal")
     .expand_columns("fit")
     .color_gradient("Sales")
     .save("report.xlsx"))
```

Works with **pandas**, **polars**, or plain `list[dict]` — powered by [narwhals](https://github.com/narwhals-dev/narwhals).

## Built-in Themes

| Theme | Description |
|---|---|
| `"default"` | Dark blue header, light blue banded rows |
| `"minimal"` | Bold header with bottom border, subtle grey bands |

See the [Demo](demo.md) page for rendered previews.
