# Future Features

## v0.2 — Conditional Formatting
- `color_gradient(col, low, mid, high)` — three-color scale (low/mid/high)
- `highlight_min_max(col)` — mark min/max cells in a column
- `conditional_style(col, fn)` — apply a `CellStyle` via a lambda, e.g. `lambda v: v < 0`

## v0.3 — Layout & Structure
- `freeze_panes(row, col)` — general pane freezing beyond just row 1
- `add_excel_table(name, style)` — native Excel ListObject with autofilter/total row
- `number_format(col, fmt)` — e.g. `"#,##0.00"` for currency columns

## v0.4 — Themes & Portability
- `apply_theme(path)` — load a `Theme` from a JSON/YAML file
- `theme.save(path)` — export the current theme to JSON/YAML
- Additional built-in themes: corporate, pastel, high-contrast

## v0.5 — Data Sources
- `list[list]` input with explicit `headers=` argument: `style(rows, headers=["a","b"])`
- Multi-sheet: `style(df1).to_sheet("Sales").add_sheet(df2, "Costs").save(...)`
- Broader narwhals support: lazy frames, chunked interchange, nullable columns

## v1.0 — Engine & CLI
- XlsxWriter engine option (faster for large write-only files)
- Engine abstraction ABC so both engines are interchangeable with a single argument
- CLI: `xlutils style data.csv --theme minimal report.xlsx`
- StyleFrame migration guide
