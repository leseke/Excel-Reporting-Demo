# Excel Reporting Demo

A client-style portfolio project that turns recurring monthly business data into a repeatable, management-ready Excel workbook.

## Client outcome

Instead of rebuilding formulas, checks, charts and KPI summaries by hand every month, the user provides a structured CSV source and runs one Python command. The project generates a complete Excel reporting package with an auditable source layer, calculated analysis, management dashboard and pre-delivery data-quality controls.

## Workbook experience

### START HERE
A client-facing landing sheet explains the workflow, workbook map and how to rebuild the report.

### Dashboard
- total revenue, cost and net profit;
- profit margin;
- revenue vs budget;
- latest revenue growth;
- revenue vs budget trend;
- monthly net profit;
- revenue growth trend;
- rolling 3-month profit.

### Analysis
A structured monthly reporting table with revenue/cost variances, profitability, growth and rolling metrics, including conditional formatting for variance review.

### Raw Data
The original synthetic source is preserved in an Excel table so the reporting pipeline stays auditable.

### Data Quality
A dedicated control sheet checks row availability, missing periods, non-positive revenue/budget values and invalid margins before the workbook is used or delivered.

## Demo results

The included synthetic 2026 dataset contains 12 reporting periods:

| KPI | Result |
| --- | ---: |
| Revenue | €267,600 |
| Cost | €151,600 |
| Net profit | €116,000 |
| Profit margin | 43.3% |
| Revenue vs budget | +€11,100 |

## Automated workflow

```text
monthly_finance.csv
        ↓
Python validation & calculations
        ↓
START HERE / Raw Data / Analysis
        ↓
Management Dashboard
        ↓
Data Quality controls
        ↓
business_report.xlsx
```

## Quick start

Requirements: Python 3.12+.

```bash
python -m pip install -r requirements.txt
python build_report.py
python -m pytest -q
```

Generated file:

```text
output/business_report.xlsx
```

## Engineering quality

The automated test suite covers the 12-month source dataset, profit and variance calculations, annual summaries, workbook generation, required sheets, dashboard KPI values and chart generation. GitHub Actions runs the suite on pushes and pull requests.

## Skills demonstrated

`Excel Reporting` · `Python` · `openpyxl` · `Data Processing` · `KPI Dashboards` · `Budget vs Actual` · `Automation` · `Data Validation` · `Automated Testing` · `GitHub Actions`

## Typical client adaptations

The same pattern can be adapted to recurring sales reports, expense/profitability reporting, inventory KPIs, multiple CSV/Excel inputs, existing company templates, custom business rules and one-click report generation.

## Scope

All repository data is synthetic. This is a portfolio demonstration of reporting automation and delivery practices, not accounting or financial advice.

---

**Have a repetitive Excel reporting process?** A client version can be adapted to existing files, KPIs and reporting workflows so fresh reports can be rebuilt consistently without repetitive manual work.
