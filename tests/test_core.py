"""Smoke tests for the shared report logic.

Run from the project folder:  python -m unittest discover -s tests -v
"""

from __future__ import annotations

import io
import sys
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import (  # noqa: E402
    DataValidationError,
    build_report,
    clean_sales_data,
    summarize,
    write_report,
)


def sample_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Product": ["Laptop", "Mouse", "Laptop", "Keyboard"],
            "Region": ["North", "South", "North", "East"],
            "Quantity": [2, 10, 1, 5],
            "Price": [800, 20, 800, 50],
        }
    )


class CleanTests(unittest.TestCase):
    def test_happy_path_adds_revenue(self) -> None:
        df = clean_sales_data(sample_df())
        self.assertEqual(len(df), 4)
        self.assertIn("Revenue", df.columns)
        self.assertEqual(df["Revenue"].iloc[0], 1600)

    def test_missing_column_raises(self) -> None:
        with self.assertRaises(DataValidationError) as ctx:
            clean_sales_data(pd.DataFrame({"Product": ["A"]}))
        self.assertIn("Missing required columns", str(ctx.exception))

    def test_empty_frame_raises(self) -> None:
        with self.assertRaises(DataValidationError):
            clean_sales_data(pd.DataFrame())

    def test_blank_product_is_dropped_not_named_nan(self) -> None:
        # regression: astype(str) used to turn a missing Product into "nan"
        df = sample_df()
        df.loc[0, "Product"] = None
        df.loc[1, "Region"] = "   "
        cleaned = clean_sales_data(df)
        self.assertEqual(len(cleaned), 2)
        self.assertNotIn("nan", set(cleaned["Product"]))
        self.assertNotIn("nan", set(cleaned["Region"]))

    def test_non_numeric_and_non_positive_rows_are_dropped(self) -> None:
        df = sample_df()
        df["Quantity"] = df["Quantity"].astype(object)  # allow a bad value below
        df.loc[0, "Quantity"] = "abc"
        df.loc[1, "Price"] = 0
        cleaned = clean_sales_data(df)
        self.assertEqual(len(cleaned), 2)

    def test_all_rows_invalid_raises(self) -> None:
        df = sample_df()
        df["Quantity"] = ["not a number"] * len(df)
        with self.assertRaises(DataValidationError):
            clean_sales_data(df)


class SummaryTests(unittest.TestCase):
    def test_metrics(self) -> None:
        s = summarize(clean_sales_data(sample_df()))
        # 1600 + 200 + 800 + 250
        self.assertEqual(s.total_revenue, 2850)
        self.assertEqual(s.transactions, 4)
        self.assertEqual(s.best_product, "Laptop")
        self.assertEqual(s.best_product_revenue, 2400)
        self.assertEqual(s.best_region, "North")
        self.assertEqual(s.weighted_price, 2850 / 18)
        self.assertEqual(list(s.top_products(2).index), ["Laptop", "Keyboard"])

    def test_ties_resolve_deterministically(self) -> None:
        df = pd.DataFrame(
            {"Product": ["A", "B"], "Region": ["X", "X"], "Quantity": [1, 1], "Price": [10, 10]}
        )
        s = summarize(clean_sales_data(df))
        self.assertEqual(s.total_revenue, 20)


class ReportTests(unittest.TestCase):
    def test_workbook_structure(self) -> None:
        df = clean_sales_data(sample_df())
        summary = summarize(df)
        data = write_report(df, summary, None)
        self.assertIsInstance(data, bytes)

        wb = load_workbook(io.BytesIO(data))
        self.assertEqual(
            wb.sheetnames,
            ["Dashboard", "Sales Data", "By Product", "By Region", "Business Insights"],
        )
        self.assertEqual(len(wb["Dashboard"]._charts), 2)
        # header of the grouped sheets must not be blank
        self.assertEqual(wb["By Product"]["A1"].value, "Product")
        self.assertEqual(wb["By Region"]["A1"].value, "Region")

    def test_custom_currency(self) -> None:
        df = clean_sales_data(sample_df())
        data = write_report(df, summarize(df), None, currency="€")
        wb = load_workbook(io.BytesIO(data))
        self.assertIn("€", wb["Sales Data"]["E2"].number_format)

    def test_build_report_writes_file(self) -> None:
        source = io.BytesIO()
        sample_df().to_excel(source, index=False)
        source.seek(0)

        target = Path(__file__).resolve().parent / "_tmp_report.xlsx"
        try:
            summary, data = build_report(source, target)
            self.assertTrue(target.exists())
            self.assertEqual(target.stat().st_size, len(data))
            self.assertEqual(summary.transactions, 4)
        finally:
            target.unlink(missing_ok=True)

    def test_build_report_rejects_unreadable_file(self) -> None:
        with self.assertRaises(DataValidationError):
            build_report(io.BytesIO(b"this is not an xlsx file"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
