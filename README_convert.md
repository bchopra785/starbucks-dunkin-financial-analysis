Export Excel sheets to HTML/PDF
=================================

This project includes a small tool to export Excel workbooks into per-sheet
HTML files (preserves basic colors, bold/italic, alignment) and optionally
generate PDFs for each sheet.

Files added
- `tools/convert_xlsx.py` — converter script
- `requirements.txt` — Python dependencies (weasyprint optional)

Quick usage

1. Create a Python environment and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Note: `weasyprint` requires additional system libraries on macOS (cairo, pango). If you prefer a binary, install `wkhtmltopdf` and the script will try that as a fallback.

2. Run the converter for the workbook in `excel/`:

```bash
python tools/convert_xlsx.py "excel/Starbucks vs. Dunkins Project - Statements (etc).xlsx" --out report/converted --pdf
```

3. Outputs will be under `report/converted/<workbook-stem>/` as `.html` (and `.pdf` if created).

Pushing to GitHub
- Add the `report/converted/` folder to your repository (committed files) before pushing.
- GitHub will render PDFs and HTML files linked from the repo; HTML files may not render inline on GitHub directly, so include an index or link to the PDF for best compatibility.

If you want, I can run the conversion now and add the generated HTML/PDF files to the repo so you can push them. Which would you prefer: generate HTML-only, or HTML+PDF (may need you to install `weasyprint` or `wkhtmltopdf`)?
