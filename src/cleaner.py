from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

REQUIRED_COLUMNS = [
    "date",
    "region",
    "product",
    "quantity",
    "unit_price",
]


def normalize_column_name(value: str) -> str:
    """Convert a source column name to a predictable snake_case name."""
    value = str(value).strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def clean_sales_data(frame: pd.DataFrame, source_file: str = "") -> pd.DataFrame:
    """Normalize and validate a sales DataFrame.

    The returned frame keeps invalid rows so they can be surfaced in the
    Excel exception report instead of being silently discarded.
    """
    df = frame.copy()
    df.columns = [normalize_column_name(column) for column in df.columns]

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    for column in ("region", "product"):
        df[column] = df[column].astype("string").str.strip()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")
    df["revenue"] = df["quantity"] * df["unit_price"]
    df["source_file"] = source_file

    required_values: Iterable[str] = ("date", "region", "product", "quantity", "unit_price")
    invalid = df[list(required_values)].isna().any(axis=1)
    invalid |= df["region"].eq("") | df["product"].eq("")
    invalid |= df["quantity"].le(0).fillna(False)
    invalid |= df["unit_price"].lt(0).fillna(False)

    df["status"] = "Valid"
    df.loc[invalid, "status"] = "Review"

    return df
