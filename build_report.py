from __future__ import annotations

from csv import DictReader
from pathlib import Path
from statistics import mean

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "data" / "monthly_finance.csv"
OUTPUT = ROOT / "output" / "business_report.xlsx"


def load_rows(path: Path = SOURCE):
    with path.open(encoding="utf-8", newline="") as f:
        return list(DictReader(f))


def calculate_metrics(rows):
    result = []
    profits = []
    for row in rows:
        revenue = float(row["revenue"])
        budget_revenue = float(row["budget_revenue"])
        cost = float(row["cost"])
        budget_cost = float(row["budget_cost"])
        profit = revenue - cost
        profits.append(profit)
        result.append({
            "month": row["month"],
            "revenue": revenue,
            "budget_revenue": budget_revenue,
            "revenue_variance": revenue - budget_revenue,
            "revenue_variance_pct": (revenue - budget_revenue) / budget_revenue if budget_revenue else 0,
            "cost": cost,
            "budget_cost": budget_cost,
            "cost_variance": cost - budget_cost,
            "profit": profit,
            "rolling_3m_profit": mean(profits[-3:]),
        })
    return result


def build_workbook(rows, output_path: Path = OUTPUT):
    metrics = calculate_metrics(rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    raw = wb.active
    raw.title = "Raw Data"
    raw_headers = ["Month", "Revenue", "Budget Revenue", "Cost", "Budget Cost"]
    raw.append(raw_headers)
    for row in rows:
        raw.append([row["month"], float(row["revenue"]), float(row["budget_revenue"]), float(row["cost"]), float(row["budget_cost"])])

    analysis = wb.create_sheet("Analysis")
    headers = ["Month", "Revenue", "Budget Revenue", "Revenue Variance", "Revenue Variance %", "Cost", "Budget Cost", "Cost Variance", "Profit", "Rolling 3M Profit"]
    analysis.append(headers)
    for item in metrics:
        analysis.append([
            item["month"], item["revenue"], item["budget_revenue"], item["revenue_variance"], item["revenue_variance_pct"],
            item["cost"], item["budget_cost"], item["cost_variance"], item["profit"], item["rolling_3m_profit"]
        ])
    for cell in analysis[1]:
        cell.font = Font(bold=True)
    for cell in raw[1]:
        cell.font = Font(bold=True)
    for row_idx in range(2, analysis.max_row + 1):
        analysis.cell(row_idx, 5).number_format = "0.0%"

    dashboard = wb.create_sheet("Dashboard")
    dashboard["A1"] = "Business Performance Dashboard"
    dashboard["A1"].font = Font(bold=True, size=16)
    dashboard["A3"] = "Total Revenue"
    dashboard["B3"] = sum(x["revenue"] for x in metrics)
    dashboard["A4"] = "Total Profit"
    dashboard["B4"] = sum(x["profit"] for x in metrics)
    dashboard["A5"] = "Average Revenue Variance %"
    dashboard["B5"] = mean(x["revenue_variance_pct"] for x in metrics)
    dashboard["B5"].number_format = "0.0%"

    line = LineChart()
    line.title = "Revenue vs Budget"
    line.y_axis.title = "Amount"
    line.x_axis.title = "Month"
    line.add_data(Reference(analysis, min_col=2, max_col=3, min_row=1, max_row=analysis.max_row), titles_from_data=True)
    line.set_categories(Reference(analysis, min_col=1, min_row=2, max_row=analysis.max_row))
    dashboard.add_chart(line, "A8")

    bar = BarChart()
    bar.title = "Monthly Profit"
    bar.add_data(Reference(analysis, min_col=9, min_row=1, max_row=analysis.max_row), titles_from_data=True)
    bar.set_categories(Reference(analysis, min_col=1, min_row=2, max_row=analysis.max_row))
    dashboard.add_chart(bar, "J8")

    for sheet in (raw, analysis, dashboard):
        for column in sheet.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column) + 2
            sheet.column_dimensions[column[0].column_letter].width = min(width, 28)

    wb.save(output_path)
    return output_path, metrics


if __name__ == "__main__":
    path, metrics = build_workbook(load_rows())
    print(f"Created {path} with {len(metrics)} reporting periods")
