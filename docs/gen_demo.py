"""Generate demo Excel files and their HTML previews for the docs site."""

from __future__ import annotations

import sys
from io import StringIO
from pathlib import Path

# ensure the package is importable even when run standalone
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from xlutils import style

DOCS_DIR = Path(__file__).resolve().parent
DOWNLOADS_DIR = DOCS_DIR / "downloads"
SNIPPETS_DIR = DOCS_DIR / "_snippets"

DEMO_DATA = [
    {"Region": "North", "Product": "Widget A", "Q1 Sales": 14200, "Q2 Sales": 16800},
    {"Region": "North", "Product": "Widget B", "Q1 Sales": 8700, "Q2 Sales": 9100},
    {"Region": "South", "Product": "Widget A", "Q1 Sales": 11500, "Q2 Sales": 13400},
    {"Region": "South", "Product": "Widget B", "Q1 Sales": 6300, "Q2 Sales": 7800},
    {"Region": "East", "Product": "Widget A", "Q1 Sales": 19000, "Q2 Sales": 21200},
    {"Region": "East", "Product": "Widget B", "Q1 Sales": 10400, "Q2 Sales": 11900},
    {"Region": "West", "Product": "Widget A", "Q1 Sales": 15600, "Q2 Sales": 17100},
    {"Region": "West", "Product": "Widget B", "Q1 Sales": 9200, "Q2 Sales": 10500},
]

SUMMARY_DATA = [
    {"Region": "North", "Total Q1": 22900, "Total Q2": 25900, "YoY Growth": "+13%"},
    {"Region": "South", "Total Q1": 17800, "Total Q2": 21200, "YoY Growth": "+19%"},
    {"Region": "East",  "Total Q1": 29400, "Total Q2": 33100, "YoY Growth": "+13%"},
    {"Region": "West",  "Total Q1": 24800, "Total Q2": 27600, "YoY Growth": "+11%"},
]


def _xlsx_to_html_fragment(xlsx_path: Path) -> str:
    """Convert xlsx to just the <table> fragment (no DOCTYPE/html/body wrapper)."""
    import re

    from xlsx2html import xlsx2html

    buf = StringIO()
    xlsx2html(str(xlsx_path), output=buf)
    full = buf.getvalue()
    match = re.search(r"<body>\s*(.*?)\s*</body>", full, re.DOTALL)
    return match.group(1) if match else full


def generate() -> None:
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    SNIPPETS_DIR.mkdir(parents=True, exist_ok=True)

    # --- Demo 1: one-liner with defaults ---
    path_default = DOWNLOADS_DIR / "demo_default.xlsx"
    style(DEMO_DATA).save(str(path_default))

    # --- Demo 2: minimal theme + gradient ---
    path_styled = DOWNLOADS_DIR / "demo_styled.xlsx"
    (
        style(DEMO_DATA)
        .apply_theme("minimal")
        .expand_columns("fit")
        .color_gradient("Q1 Sales")
        .color_gradient("Q2 Sales")
        .save(str(path_styled))
    )

    # --- Demo 3: multi-sheet workbook with per-sheet styling ---
    path_multi = DOWNLOADS_DIR / "demo_multisheet.xlsx"
    (
        style()
        .add_sheet(DEMO_DATA, "Sales")
            .apply_theme("default")
            .expand_columns("fit")
            .color_gradient("Q1 Sales")
            .color_gradient("Q2 Sales")
            .back()
        .add_sheet(SUMMARY_DATA, "Summary")
            .apply_theme("minimal")
            .expand_columns(18)
            .freeze_header(False)
            .back()
        .save(str(path_multi))
    )

    # --- Generate HTML table fragments (for snippet inclusion only) ---
    for xlsx_path in [path_default, path_styled, path_multi]:
        fragment = _xlsx_to_html_fragment(xlsx_path)
        html_path = SNIPPETS_DIR / xlsx_path.with_suffix(".html").name
        html_path.write_text(fragment)

    print(f"Generated demos: downloads in {DOWNLOADS_DIR}, snippets in {SNIPPETS_DIR}")


if __name__ == "__main__":
    generate()
