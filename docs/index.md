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

## Multi-Sheet Workbooks

Call `style()` with no data, then add sheets with `.add_sheet()`. Each sheet
gets its own data, theme, and settings. Use `.back()` to return to the workbook
level and chain the next sheet.

```python
from xlutils import style

(style()
    .add_sheet(detail_data, "Sales")
        .apply_theme("default")
        .expand_columns("fit")
        .color_gradient("Q1 Sales")
        .color_gradient("Q2 Sales")
        .back()
    .add_sheet(summary_data, "Summary")
        .apply_theme("minimal")
        .expand_columns(18)
        .freeze_header(False)
        .back()
    .save("report.xlsx"))
```

You can also set **workbook-level defaults** that all sheets inherit unless
overridden individually:

```python
(style()
    .apply_theme("minimal")          # default for every sheet
    .add_sheet(q1_data, "Q1")
        .color_gradient("Revenue")
        .back()
    .add_sheet(q2_data, "Q2")
        .apply_theme("default")      # overrides the workbook default
        .color_gradient("Revenue")
        .back()
    .save("quarterly.xlsx"))
```

## Built-in Themes

| Theme | Description |
|---|---|
| `"default"` | Dark blue header, light blue banded rows |
| `"minimal"` | Bold header with bottom border, subtle grey bands |

See the [Demo](demo.md) page for rendered previews.
