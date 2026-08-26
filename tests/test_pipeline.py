from pathlib import Path

from src.pipeline import consolidate_csv_files


def test_consolidate_csv_files_removes_duplicates(tmp_path: Path):
    csv = "Date,Region,Product,Quantity,Unit Price\n2026-01-01,North,Laptop,2,1000\n"
    (tmp_path / "a.csv").write_text(csv, encoding="utf-8")
    (tmp_path / "b.csv").write_text(csv, encoding="utf-8")

    result = consolidate_csv_files(tmp_path)

    assert len(result) == 1
    assert result.loc[0, "revenue"] == 2000
