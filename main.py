from pathlib import Path

from src.excel_report import build_excel_report
from src.pipeline import consolidate_csv_files


def main() -> None:
    root = Path(__file__).resolve().parent
    data = consolidate_csv_files(root / "data")
    output = build_excel_report(data, root / "output" / "automated_sales_report.xlsx")
    print(f"Report generated: {output}")
    print(f"Records processed: {len(data)}")
    print(f"Records to review: {(data['status'] == 'Review').sum()}")


if __name__ == "__main__":
    main()
