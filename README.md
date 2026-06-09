# CALG Coefficient Generator

Flask application for generating CALG regression summaries, coefficients, equations, and Excel workbooks from uploaded Excel files.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

Open:

```text
http://127.0.0.1:5000
```

## Workflow

1. Upload a `.xlsx` or `.xls` CALG file.
2. Enter the number of output columns.
3. Click Generate.
4. The Flask backend generates `CALG_Results.xlsx`.
5. The browser downloads the generated workbook automatically.

The browser only handles upload, progress display, and download. Regression calculations and workbook generation run in Python.
