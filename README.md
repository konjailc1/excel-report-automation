# Excel Report Automation

Turn a flat sales export into a formatted weekly Excel report with one command.

## Problem

A retail team receives a CSV of product-line sales and rebuilds the same workbook every week: headline numbers, a week-over-week comparison, breakdowns by region and category, and charts for the weekly meeting. Doing that by hand is slow, and a wrong total is easy to miss when the figures are in VND.

This project builds that workbook in Python. Pandas calculates the totals and pivot tables. openpyxl writes the sheets, charts, and formatting.

## Features

- **Summary** — KPI cards for total revenue, orders (product lines), average order value, and week-over-week revenue change. Increases are green and decreases are red.
- **By Region** and **By Category** — revenue by week from `pandas.pivot_table`, with a total row and a total column.
- **Weekly Trend** — a line chart of revenue by week, and a bar chart of the top 10 products in the selected week.
- **Raw Data** — the rows behind the report as an Excel Table with filters and a frozen header.
- VND amounts use the `#,##0` format. Columns on the data sheets are sized to fit their values.

## Screenshot

> **Placeholder.** Open `sample_output/weekly_report_week8.xlsx`, capture the Summary sheet, and add it here:
>
> `![Summary sheet](sample_output/screenshot.png)`

## How to run

From this folder, activate the virtual environment first:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python make_sample.py
python main.py sample_input/sales_data.csv --week 8
```

The report is written to `output/weekly_report_week8.xlsx`.

A finished copy is already in `sample_output/weekly_report_week8.xlsx`, so you can open the workbook before generating your own. `output/` is gitignored. `sample_output/` is the example that stays with the project.

Change the week to report a different slice. Weeks after the one you pick are left out, and the comparison is against the previous week:

```powershell
python main.py sample_input/sales_data.csv --week 5 --output output/weekly_report_week5.xlsx
```

### Your own CSV

`report.py` expects these columns:

`date`, `week`, `region`, `store`, `product_category`, `product`, `quantity`, `unit_price`, `revenue`, `salesperson`

`week` is a whole number (1, 2, 3, …). Each row is one product line. The orders KPI is a count of those rows. Dates should look like `2026-10-12`.

### Sample data

`make_sample.py` writes 1,000 synthetic rows for a Vietnamese retail chain (Miền Bắc, Miền Trung, Miền Nam) across 8 weeks starting 24 Aug 2026. The random seed is fixed, so generating the file again produces the same rows. The product list at the top of `make_sample.py` is the catalog.

## Project layout

| File | Purpose |
| --- | --- |
| `make_sample.py` | Generate `sample_input/sales_data.csv` |
| `report.py` | Calculations and Excel formatting |
| `main.py` | Command-line interface |
| `sample_input/` | Sample CSV |
| `sample_output/` | Example workbook |
| `output/` | Reports you generate (not committed) |

## Requirements

- Windows
- Python 3.10 or newer
- pandas and openpyxl (listed in `requirements.txt`)
