from pathlib import Path

from build_report import calculate_metrics, load_rows, build_workbook


def test_metrics_are_calculated():
    rows = load_rows()
    metrics = calculate_metrics(rows)
    assert len(metrics) == 6
    assert metrics[0]["profit"] == 7400
    assert round(metrics[0]["revenue_variance_pct"], 4) == 0.04


def test_workbook_is_generated(tmp_path: Path):
    output = tmp_path / "report.xlsx"
    path, metrics = build_workbook(load_rows(), output)
    assert path.exists()
    assert path.stat().st_size > 0
    assert len(metrics) == 6
