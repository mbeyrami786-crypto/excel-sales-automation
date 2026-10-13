"""End-to-end tests for the Streamlit UI (runs the real app.py headlessly).

Run from the project folder:  python -m unittest discover -s tests -v
"""

from __future__ import annotations

import io
import sys
import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def workbook_bytes(rows: int = 4) -> bytes:
    import pandas as pd

    df = pd.DataFrame(
        {
            "Product": ["Laptop", "Mouse", "Laptop", "Keyboard"][:rows],
            "Region": ["North", "South", "North", "East"][:rows],
            "Quantity": [2, 10, 1, 5][:rows],
            "Price": [800, 20, 800, 50][:rows],
        }
    )
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    return buffer.getvalue()


def run_app(upload: bytes | None = None) -> AppTest:
    app = AppTest.from_file(str(PROJECT / "app.py"), default_timeout=60)
    app.run()
    if upload is not None:
        app.file_uploader[0].upload("sales.xlsx", upload, XLSX_MIME)
        app.run()
    return app


class AppTests(unittest.TestCase):
    def test_no_file_shows_instructions(self) -> None:
        app = run_app()
        self.assertFalse(app.exception)
        self.assertEqual([t.value for t in app.title], ["AI Sales Automation Dashboard"])
        self.assertIn("No file selected", app.info[0].value)
        self.assertEqual(len(app.metric), 0)

    def test_valid_upload_renders_metrics_and_download(self) -> None:
        app = run_app(workbook_bytes())
        self.assertFalse(app.exception)
        self.assertFalse(app.error)

        metrics = {m.label: m.value for m in app.metric}
        self.assertEqual(metrics["Total Revenue"], "$2,850.00")
        self.assertEqual(metrics["Total Quantity"], "18")
        self.assertEqual(metrics["Transactions"], "4")
        self.assertEqual(metrics["Best Product"], "Laptop")
        self.assertEqual(metrics["Best Region"], "North")
        self.assertEqual(metrics["Rows Discarded"], "0")

        # preview + top-5 tables
        self.assertGreaterEqual(len(app.dataframe), 3)
        # the report download must be offered
        self.assertEqual(len(app.download_button), 1)

    def test_missing_columns_shows_error_not_traceback(self) -> None:
        import pandas as pd

        buffer = io.BytesIO()
        pd.DataFrame({"Color": ["Red"], "Size": [1]}).to_excel(buffer, index=False)

        app = run_app(buffer.getvalue())
        self.assertFalse(app.exception)  # no crash
        self.assertTrue(app.error)
        self.assertIn("Missing required columns", app.error[0].value)
        self.assertEqual(len(app.metric), 0)

    def test_corrupt_file_shows_error(self) -> None:
        app = run_app(b"this is not a spreadsheet")
        self.assertFalse(app.exception)
        self.assertTrue(app.error)
        self.assertEqual(len(app.metric), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
