from __future__ import annotations

from pathlib import Path

import pandas as pd

from .cleaner import clean_sales_data


def consolidate_csv_files(data_dir: str | Path) -> pd.DataFrame:
    """Load, clean and consolidate every CSV file in *data_dir*."""
    data_path = Path(data_dir)
    files = sorted(data_path.glob("*.csv"))

    if not files:
        raise FileNotFoundError(f"No CSV files found in {data_path}")

    frames = [
        clean_sales_data(pd.read_csv(path), source_file=path.name)
        for path in files
    ]
    combined = pd.concat(frames, ignore_index=True)

    duplicate_subset = ["date", "region", "product", "quantity", "unit_price"]
    combined = combined.drop_duplicates(subset=duplicate_subset, keep="first")
    combined = combined.sort_values(["date", "region", "product"], na_position="last")
    return combined.reset_index(drop=True)
