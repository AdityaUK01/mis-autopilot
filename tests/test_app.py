import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest


APP = Path(__file__).resolve().parents[1] / "app.py"


class StreamlitSmokeTests(unittest.TestCase):
    def test_demo_app_renders_without_errors(self):
        app = AppTest.from_file(str(APP)).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertEqual(app.title[0].value, "MIS Autopilot")
        self.assertEqual(len(app.metric), 4)
        self.assertEqual(len(app.download_button), 3)

    def test_app_prompts_for_upload_when_optional_demo_file_is_missing(self):
        original_is_file = Path.is_file

        def hide_optional_sample(path):
            if path == APP.parent / "sample_sales.csv":
                return False
            return original_is_file(path)

        with patch.object(Path, "is_file", hide_optional_sample):
            app = AppTest.from_file(str(APP)).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertTrue(any("optional demo file is not present" in message.value for message in app.info))

    def test_malformed_upload_shows_actionable_error_without_crashing(self):
        app = AppTest.from_file(str(APP)).run(timeout=15)
        app.file_uploader[0].set_value(("bad.csv", b'"unterminated header\n', "text/csv")).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertTrue(any("valid CSV with consistent rows" in message.value for message in app.error))

    def test_csv_upload(self):
        app = AppTest.from_file(str(APP)).run(timeout=15)
        csv_data = b"Order ID,Date,Product,Quantity,Unit Price,Total Sales,Region\nA1,2026-01-01,Soap,2,50,100,Delhi\n"
        app.file_uploader[0].set_value(("sales.csv", csv_data, "text/csv")).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertEqual(app.metric[0].value, "₹100.00")

    def test_excel_upload(self):
        excel_data = BytesIO()
        pd.DataFrame({
            "Order ID": ["A1"], "Date": ["2026-01-01"], "Product": ["Soap"],
            "Quantity": [2], "Unit Price": [50], "Total Sales": [100], "Region": ["Delhi"],
        }).to_excel(excel_data, index=False)
        app = AppTest.from_file(str(APP)).run(timeout=15)
        app.file_uploader[0].set_value(("sales.xlsx", excel_data.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertEqual(app.metric[0].value, "₹100.00")

    def test_repeated_transaction_schema_renders_cleaned_quality_and_download_paths(self):
        csv_data = (
            b"Record ID,Order ID,Date,Product,Quantity,Unit Price,Total Sales,Region\n"
            b"REC-1,A1,2026-01-01,Soap,1,10,10,Delhi\n"
            b"REC-2,A1,2026-01-01,Soap,1,10,10,Delhi\n"
        )
        app = AppTest.from_file(str(APP)).run(timeout=15)
        app.file_uploader[0].set_value(("repeated.csv", csv_data, "text/csv")).run(timeout=15)
        self.assertFalse(app.exception)
        # Streamlit renders all tab bodies in this script run, including the
        # Cleaned data and Data Quality bodies that previously raised KeyError.
        self.assertEqual(app.metric[3].value, "1")
        self.assertEqual(len(app.download_button), 3)

    def test_finance_template_runs_on_non_sales_columns(self):
        csv_data = b"Expense Amount,Category,Budget,Actual\n100,Rent,120,100\n200,Supplies,250,200\n"
        app = AppTest.from_file(str(APP)).run(timeout=15)
        app.selectbox[0].set_value("Finance")
        app.file_uploader[0].set_value(("finance.csv", csv_data, "text/csv"))
        app.run(timeout=15)
        self.assertFalse(app.exception)
        self.assertTrue(any(metric.label == "Mapped expense total" and metric.value == "₹300.00" for metric in app.metric))

    def test_general_template_renders_generic_profile_for_unmatched_dataset(self):
        csv_data = b"Team,Monthly Cost,Status\nA,1200,Open\nB,900,Closed\n"
        app = AppTest.from_file(str(APP)).run(timeout=15)
        app.selectbox[0].set_value("General spreadsheet analysis")
        app.file_uploader[0].set_value(("operations.csv", csv_data, "text/csv"))
        app.run(timeout=15)
        self.assertFalse(app.exception)
        self.assertTrue(any(metric.label == "Source rows" and metric.value == "2" for metric in app.metric))
        self.assertEqual(len(app.download_button), 3)

    def test_hr_inventory_and_operations_templates_render(self):
        examples = [
            ("HR and attendance", "hr.csv", b"Employee ID,Attendance Status\nE1,Present\nE2,Absent\n"),
            ("Inventory", "inventory.csv", b"SKU,Stock Quantity,Reorder Level\nX,3,5\nY,10,4\n"),
            ("Operations", "operations.csv", b"Actual Output,Target Output,Task Status,Owner\n8,10,Open,A\n12,10,Done,B\n"),
        ]
        for template, filename, csv_data in examples:
            with self.subTest(template=template):
                app = AppTest.from_file(str(APP)).run(timeout=15)
                app.selectbox[0].set_value(template)
                app.file_uploader[0].set_value((filename, csv_data, "text/csv"))
                app.run(timeout=15)
                self.assertFalse(app.exception)
                self.assertEqual(len(app.download_button), 3)


if __name__ == "__main__":
    unittest.main()
