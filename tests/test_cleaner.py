import pandas as pd

from src.cleaner import clean_sales_data, normalize_column_name


def test_normalize_column_name():
    assert normalize_column_name(" Unit Price ") == "unit_price"


def test_clean_sales_data_calculates_revenue_and_status():
    source = pd.DataFrame(
        {
            "Date": ["2026-01-01", "bad-date"],
            "Region": [" North ", "South"],
            "Product": ["Laptop", "Mouse"],
            "Quantity": [2, 3],
            "Unit Price": [1000, 20],
        }
    )

    result = clean_sales_data(source, "sample.csv")

    assert result.loc[0, "region"] == "North"
    assert result.loc[0, "revenue"] == 2000
    assert result.loc[0, "status"] == "Valid"
    assert result.loc[1, "status"] == "Review"
    assert result.loc[0, "source_file"] == "sample.csv"
