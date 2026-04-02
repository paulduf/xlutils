# xlutils

Minimal boilerplate Excel styling for dataframes — fluent API, Polars/Pandas support via [narwhals](https://github.com/narwhals-dev/narwhals).

## Install

```bash
pip install xlutils
# with pandas or polars support
pip install "xlutils[pandas]"
pip install "xlutils[polars]"
```

## Quick start

```python
from xlutils import XlStyler

(XlStyler(df)
 .freeze_header()
 .auto_width()
 .save("report.xlsx"))
```

## Documentation

Full docs are deployed at the GitHub Pages site for this repo.

## Roadmap

See [PLAN_FEATURES.md](PLAN_FEATURES.md) for planned features (v0.2 → v1.0).

## Requirements

Python 3.10+, `openpyxl`, `pydantic`, `narwhals`.
