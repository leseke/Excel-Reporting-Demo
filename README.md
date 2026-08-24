# Excel Reporting Demo

A client-style portfolio project showing how raw monthly business data can be turned into a repeatable, management-ready Excel reporting workbook.

## Business problem

A small company tracks revenue, costs and budgets in recurring monthly files. Management needs a clear view of performance without rebuilding formulas, charts and summaries by hand every month.

This demo automates that reporting flow from a simple CSV source into a structured Excel workbook with KPIs, budget comparisons, trends and reusable analysis sheets.

## What the generated workbook contains

### Dashboard

- total revenue;
- total cost;
- net profit;
- profit margin;
- revenue vs budget variance;
- latest revenue growth;
- revenue vs budget chart;
- monthly net profit chart;
- revenue growth trend;
- rolling 3-month profit trend.

### Analysis

A structured monthly table containing:

- revenue and budget revenue;
- revenue variance and variance %;
- cost and budget cost;
- cost variance;
- net profit;
- profit margin;
- month-over-month revenue growth;
- rolling 3-month profit.

### Raw Data

The original synthetic source data is preserved in a clean Excel table so the reporting pipeline remains auditable and easy to adapt.

## Demo results

The included 2026 synthetic dataset contains 12 reporting periods and produces the following annual summary:

| KPI | Result |
| --- | ---: |
| Revenue | €267,600 |
| Cost | €151,600 |
| Net profit | €116,000 |
| Profit margin | 43.3% |
| Revenue vs budget | +€11,100 |

## How it works

```text
monthly_finance.csv
        ↓
Python validation & calculations
        ↓
Raw Data sheet
        ↓
Analysis sheet
        ↓
Management Dashboard
        ↓
business_report.xlsx
```

The workbook is generated with Python and `openpyxl`; no manual copy/paste is required to rebuild the report.

## Quick start

Requirements: Python 3.12+.

```bash
python -m pip install -r requirements.txt
python build_report.py
pytest -q
```

The script creates:

```text
output/business_report.xlsx
```

## Automated validation

The test suite verifies:

- the full 12-month source dataset;
- core profit and variance calculations;
- annual summary totals;
- workbook generation;
- expected workbook sheets;
- KPI values written to the dashboard;
- chart generation.

GitHub Actions runs the test suite automatically on pushes and pull requests.

## Skills demonstrated

`Excel Reporting` · `Python` · `openpyxl` · `Data Processing` · `KPI Dashboards` · `Budget vs Actual` · `Automation` · `Automated Testing` · `GitHub Actions`

## Typical client adaptations

This pattern can be adapted to:

- recurring sales reporting;
- budget vs actual dashboards;
- expense and profitability reporting;
- inventory or operational KPI reports;
- multiple CSV/Excel inputs;
- existing company workbook templates;
- custom formulas, categories and business rules;
- scheduled or one-click report generation.

## Scope

All data in this repository is synthetic. The project demonstrates a reporting and automation pattern, not accounting or financial advice.

---

**Have a repetitive Excel reporting process?** A client version can be adapted to your existing files, KPIs and reporting workflow so the same report can be rebuilt consistently from fresh data.
