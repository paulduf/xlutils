# xlutils
Util library to style dataframes when exporting to Excel

# Plan 

Building a modern Excel utility library requires a clear separation between **Data Ingestion**, **Style Logic**, and **Engine Execution**. This prevents the library from becoming obsolete when a new dataframe library (like Polars) or a new Excel engine emerges.

Below is a technical implementation plan for a library we'll call **`XlShaper`**.

---

## 🏗️ Phase 1: Core Architecture & Ingestion
The goal is to decouple the data source from the formatting logic.

* **Data Interchange Layer:** Implement a `DataHandler` class that detects the input type.
    * Support for **PEP 660** (Dataframe Interchange Protocol).
    * Fallbacks for `list[dict]` and `list[list]`.
* **The Schema Model:** Use **Pydantic** for the styling engine. This ensures type safety and allows users to export/import "Themes" as JSON/YAML.
* **Engine Wrapper:** A unified interface for `XlsxWriter` (best for new files) and `Openpyxl` (best for editing).

---

## 🎨 Phase 2: The Styling Engine (The "Brain")
Instead of applying styles cell-by-cell, we define **Rules**.

### **Key Components:**
* **Styler Object:** A Pydantic model containing:
    * `font`: (size, name, color, bold, italic)
    * `fill`: (pattern_type, fg_color, bg_color)
    * `border`: (top, bottom, left, right)
    * `alignment`: (horizontal, vertical, wrap_text)
* **Rule Mapper:**
    * `HeaderRule`: Styles for the first row.
    * `BandRule`: Logic for alternating row colors.
    * `ConditionalRule`: A lambda or string-based check (e.g., `if value < 0: apply Styler(color='red')`).

---

## 📏 Phase 3: Layout & UX Utilities
These are the "human-readable" features that `StyleFrame` lacks or handles clumsily.

* **Auto-Column Width:**
    * Calculate width using: $Width = \max(\text{len}(str(x))) \times \text{font\_modifier} + \text{padding}$
* **Fluent API (Chaining):**
    ```python
    (XlShaper(data)
     .use_theme("corporate_blue")
     .set_columns(width="auto", align="center")
     .add_excel_table(name="SalesData", style="TableStyleMedium9")
     .freeze_panes(row=1)
     .save("report.xlsx"))
    ```

---

## 🛠️ Phase 4: Modern Technical Specs

### **1. Library Metadata**
* **Language:** Python 3.10+ (utilizing `|` for Union types and `Self` return types).
* **Core Dependencies:** `openpyxl`, `pydantic`, `typing_extensions`.
* **Optional Dependencies:** `pandas`, `polars`, `xlsxwriter`.

### **2. Project Structure**
```text
xlshaper/
├── engines/
│   ├── base.py          # Abstract Base Class for engines
│   ├── xlsxwriter.py    # XlsxWriter implementation
│   └── openpyxl.py      # Openpyxl implementation
├── models/
│   ├── style.py         # Pydantic styling schemas
│   └── theme.py         # Pre-defined color palettes
├── handlers/
│   ├── dataframe.py     # Protocol-based data ingestion
│   └── sequence.py      # List/Tuple ingestion
├── formatter.py         # The main Fluent API entry point
└── utils/
    └── calc.py          # Column width & coordinate helpers
```

---

## 🚀 Implementation Roadmap

| Milestone | Deliverable | Key Feature |
| :--- | :--- | :--- |
| **v0.1.0** | Core Ingestion | Support Polars/Pandas via Interchange Protocol. |
| **v0.2.0** | Styling Model | Pydantic-based `Style` and `Theme` objects. |
| **v0.3.0** | Engine Layer | Basic `openpyxl` writer with auto-column width. |
| **v0.4.0** | The "Excel Table" | Native `ListObject` support with filters/total rows. |
| **v1.0.0** | Production Ready | Documentation, CI/CD, and "StyleFrame" migration guide. |

---

### **Why this wins over StyleFrame:**
1.  **Maintenance:** By using Pydantic and ABCs (Abstract Base Classes), the code is easier to test and extend.
2.  **Performance:** Polars users can export to Excel without ever converting to Pandas (saving memory).
3.  **Flexibility:** You can swap the underlying engine (Openpyxl vs XlsxWriter) with a single argument while keeping your styles identical.
