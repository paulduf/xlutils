# Demo

All examples use this dataset:

```python
data = [
    {"Region": "North", "Product": "Widget A", "Q1 Sales": 14200, "Q2 Sales": 16800},
    {"Region": "North", "Product": "Widget B", "Q1 Sales": 8700,  "Q2 Sales": 9100},
    {"Region": "South", "Product": "Widget A", "Q1 Sales": 11500, "Q2 Sales": 13400},
    {"Region": "South", "Product": "Widget B", "Q1 Sales": 6300,  "Q2 Sales": 7800},
    {"Region": "East",  "Product": "Widget A", "Q1 Sales": 19000, "Q2 Sales": 21200},
    {"Region": "East",  "Product": "Widget B", "Q1 Sales": 10400, "Q2 Sales": 11900},
    {"Region": "West",  "Product": "Widget A", "Q1 Sales": 15600, "Q2 Sales": 17100},
    {"Region": "West",  "Product": "Widget B", "Q1 Sales": 9200,  "Q2 Sales": 10500},
]
```

---

## Default Theme

One-liner with no configuration:

```python
from xlutils import style

style(data).save("demo_default.xlsx")
```

**Preview:**

--8<-- "docs/_snippets/demo_default.html"

<a href="/downloads/demo_default.xlsx" download data-wm-adjusted="done" target="_top">Download demo_default.xlsx</a>

---

## Minimal Theme + Color Gradient

Fluent API with theme, auto-fit columns, and gradient on sales columns.
The color gradient uses Excel conditional formatting — download the file to see it in action:

```python
(style(data)
    .apply_theme("minimal")
    .expand_columns("fit")
    .color_gradient("Q1 Sales")
    .color_gradient("Q2 Sales")
    .save("demo_styled.xlsx"))
```

**Preview:**

--8<-- "docs/_snippets/demo_styled.html"

<a href="/downloads/demo_styled.xlsx" download data-wm-adjusted="done" target="_top">Download demo_styled.xlsx</a>
