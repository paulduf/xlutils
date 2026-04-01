from __future__ import annotations

from typing import Any

import narwhals as nw


def normalize(data: Any) -> tuple[list[str], list[list]]:
    """Return (headers, rows) from a pandas/polars DataFrame or list[dict]."""
    if isinstance(data, list):
        if not data:
            return [], []
        headers = list(data[0].keys())
        rows = [list(d.values()) for d in data]
        return headers, rows

    df = nw.from_native(data, eager_only=True)
    headers = list(df.columns)
    rows = [list(row) for row in df.rows()]
    return headers, rows
