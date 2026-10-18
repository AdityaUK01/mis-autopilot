import unittest
from io import BytesIO
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import Workbook

from spreadsheet_core import load_spreadsheet, profile_spreadsheet


class SpreadsheetCoreTests(unittest.TestCase):
    def test_csv_loader_keeps_duplicate_columns_and_values(self):
        frame, sheets = load_spreadsheet(b"Name,Amount,Amount\nTea,10,12\n", "dup.csv")
        self.assertEqual(sheets, [])
        self.assertEqual(list(frame.columns), ["Name", "Amount", "Amount [2]"])
        self.assertEqual(frame.loc[0].tolist(), ["Tea", "10", "12"])

    def test_excel_loader_selects_sheet_and_keeps_source_values(self):
        output = BytesIO()
        workbook = Workbook()
        first = workbook.active
        first.title = "Notes"
        first.append(["Description"])
        first.append(["ignore"])
        sales = workbook.create_sheet("Actual Data")
        sales.append(["Item", "Amount"])
        sales.append(["Tea", 42])
        workbook.save(output)

        first_frame, sheets = load_spreadsheet(output.getvalue(), "workbook.xlsx")
        selected_frame, selected_sheets = load_spreadsheet(output.getvalue(), "workbook.xlsx", "Actual Data")
        self.assertEqual(sheets, ["Notes", "Actual Data"])
        self.assertEqual(selected_sheets, sheets)
        self.assertEqual(first_frame.iloc[0, 0], "ignore")
        self.assertEqual(selected_frame.iloc[0].tolist(), ["Tea", 42])

    def test_legacy_xls_uses_xlrd_engine(self):
        excel = Mock()
        excel.sheet_names = ["Legacy"]
        headerless = pd.DataFrame([["Amount"], [7]])
        with patch("spreadsheet_core.pd.ExcelFile", return_value=excel) as excel_file, patch(
            "spreadsheet_core.pd.read_excel", return_value=headerless
        ) as read_excel:
            frame, sheets = load_spreadsheet(b"mock legacy workbook bytes", "legacy.xls")
        self.assertEqual(excel_file.call_args.kwargs["engine"], "xlrd")
        self.assertEqual(read_excel.call_args.kwargs["sheet_name"], "Legacy")
        self.assertEqual(frame.loc[0, "Amount"], 7)
        self.assertEqual(sheets, ["Legacy"])

    def test_unsupported_and_empty_files_have_clear_errors(self):
        with self.assertRaisesRegex(ValueError, "Unsupported file format"):
            load_spreadsheet(b"x", "input.json")
        with self.assertRaisesRegex(ValueError, "empty"):
            load_spreadsheet(b"", "input.csv")

    def test_malformed_csv_and_excel_have_clear_errors(self):
        with self.assertRaisesRegex(ValueError, "valid CSV"):
            load_spreadsheet(b'"unterminated header\n', "bad.csv")
        with self.assertRaisesRegex(ValueError, "valid, unprotected workbook"):
            load_spreadsheet(b"not an Excel workbook", "bad.xlsx")

    def test_profile_flags_invalid_dates_numbers_missing_values_and_duplicates_without_dropping(self):
        source = pd.DataFrame([
            {"Work Date": "2026-01-01", "Amount": "100", "Team": "A"},
            {"Work Date": "not-a-date", "Amount": "oops", "Team": ""},
            {"Work Date": "2026-01-01", "Amount": "100", "Team": "A"},
            {"Work Date": "", "Amount": "", "Team": ""},
        ])
        profile = profile_spreadsheet(source, "quality.csv")
        self.assertEqual(profile["row_count"], 4)
        self.assertEqual(profile["blank_row_count"], 1)
        self.assertEqual(profile["duplicate_count"], 1)
        self.assertEqual(len(profile["original_data"]), 4)
        self.assertTrue(profile["original_data"].equals(source))
        kinds = set(profile["issues"]["Issue"])
        self.assertIn("Invalid date", kinds)
        self.assertIn("Invalid numeric value", kinds)
        self.assertIn("Missing value", kinds)
        self.assertIn("Exact duplicate row", kinds)
        self.assertEqual(profile["detected_types"]["Work Date"], "date")
        self.assertEqual(profile["detected_types"]["Amount"], "numeric")
        self.assertEqual(profile["issues"].loc[profile["issues"]["Issue"].eq("Exact duplicate row"), "Row"].iloc[0], 4)


if __name__ == "__main__":
    unittest.main()
