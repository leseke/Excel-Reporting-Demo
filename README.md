# Excel Reporting Demo

A reproducible proof-of-concept for turning raw monthly business data into a refreshable Excel reporting workbook.

## What this demonstrates

- clean separation between raw data, calculations and dashboard output;
- budget vs actual comparisons;
- monthly KPI calculations;
- variance percentages;
- rolling 3-month averages;
- automatic workbook generation from synthetic CSV data;
- chart-ready Excel output;
- automated checks for key calculations.

## Quick start

```bash
python -m pip install -r requirements.txt
python build_report.py
pytest -q
```

The script creates `output/business_report.xlsx` from `data/monthly_finance.csv`.

## Workbook structure

- **Raw Data** — source rows kept in tabular form.
- **Analysis** — monthly revenue, cost, profit, budget variance and rolling metrics.
- **Dashboard** — KPI summary and charts generated from the analysis sheet.

## Example use cases

- recurring management reporting;
- budget vs actual tracking;
- monthly sales or finance dashboards;
- replacing repeated copy/paste reporting steps;
- adapting an existing workbook into a repeatable process.

## Scope and limitations

This repository uses synthetic data and demonstrates the engineering pattern, not accounting advice. A client project would adapt formulas, source columns, business rules and refresh logic to the real workbook and reporting requirements.

## Author

Ylan Bitang — automation and data tools.