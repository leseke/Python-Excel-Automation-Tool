# Python Excel Automation Tool

A practical Python automation project that consolidates multiple CSV files into a clean, structured Excel workbook with automated validation, data-quality checks, KPI reporting and a dashboard.

This repository is the public source-code companion to the **Python Excel Automation & Reporting** portfolio project.

## What it does

- Imports multiple CSV files from a folder
- Normalizes column names and common text fields
- Converts dates and numeric values to consistent types
- Detects missing values and duplicate rows
- Consolidates the cleaned records into one dataset
- Generates an Excel workbook automatically
- Adds KPI formulas and a summary dashboard
- Creates an exception sheet for records that need review
- Keeps the workflow reproducible and easy to run

## Project structure

```text
Python-Excel-Automation-Tool/
├── data/
│   ├── north_sales.csv
│   ├── south_sales.csv
│   └── west_sales.csv
├── output/
│   └── .gitkeep
├── src/
│   ├── __init__.py
│   ├── cleaner.py
│   ├── excel_report.py
│   └── pipeline.py
├── tests/
│   ├── test_cleaner.py
│   └── test_pipeline.py
├── .github/
│   └── workflows/
│       └── tests.yml
├── .gitignore
├── LICENSE
├── README.md
├── main.py
└── requirements.txt
```

## Quick start

```bash
python -m venv .venv
```

Activate the virtual environment, then install dependencies:

```bash
pip install -r requirements.txt
```

Run the automation:

```bash
python main.py
```

The generated workbook is written to:

```text
output/automated_sales_report.xlsx
```

## Excel deliverable

The generated workbook contains:

- **Dashboard** — headline KPIs and charts
- **Consolidated Data** — cleaned, analysis-ready records
- **Exceptions** — rows that require manual review

The dashboard includes total records, valid records, records to review, total quantity, total revenue and unique products, plus visual summaries by region and product.

## Data-quality controls

A record is flagged for review when required information is missing or when values cannot be converted to the expected data type. Duplicate source rows are removed during consolidation.

## Tests

```bash
pytest
```

Tests cover the core cleaning and consolidation behavior. GitHub Actions runs the test suite automatically on pushes and pull requests.

## Tech stack

- Python 3.11+
- pandas
- openpyxl
- pytest
- Microsoft Excel
- GitHub Actions

## Portfolio purpose

This project demonstrates a typical freelance workflow: taking fragmented business data, validating and consolidating it, then delivering a professional Excel report that can be refreshed from new source files with a single command.

## License

Released under the MIT License.
