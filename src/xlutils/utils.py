from __future__ import annotations


def calc_col_widths(
    headers: list[str],
    rows: list[list],
    padding: int = 2,
    max_width: int = 60,
) -> list[int]:
    """Return one integer width per column."""
    widths = []
    for col_idx, header in enumerate(headers):
        col_values = [row[col_idx] for row in rows]
        max_data = max((len(str(v)) for v in col_values), default=0)
        width = min(max(len(header), max_data) + padding, max_width)
        widths.append(width)
    return widths
