Starbucks vs Dunkins — Financial Analysis
=========================================

This repository contains financial statements and analysis comparing Starbucks and Dunkin'.

Contents
- `excel/` — original Excel workbooks (source data)
- `report/converted/` — HTML and PDF exports generated from the workbook (per-sheet)
- `tools/convert_xlsx.py` — script to convert workbooks into color-preserving HTML (and optional PDFs)
- `requirements.txt` — Python dependencies to run the converter

How to reproduce the exports

1. Create and activate a Python virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install Python dependencies:

```bash
pip install -r requirements.txt
```

3. Convert the workbook (example):

```bash
python tools/convert_xlsx.py "excel/Starbucks vs. Dunkins Project - Statements (etc).xlsx" --out report/converted --pdf
```

Notes
- The converter preserves cell background colors, bold/italic, and basic alignment, then writes per-sheet `.html` files.
- PDF generation uses `weasyprint` or `wkhtmltopdf` if available. `weasyprint` may require additional system libraries on macOS (cairo, pango).

If you want, I can now run the converter and push the generated HTML and PDFs to GitHub for you.
