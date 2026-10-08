from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from md_xlsx_convertor.excel_to_markdown import create_markdown


class ExcelToMarkdownTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.workbook_path = self.root / "spec.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Cases"
        sheet.append(["Title"])
        sheet.append(["Metadata"])
        sheet.append(["#", "Name", "Name"])
        sheet.append([1, "first", ""])
        sheet.append([2, "second", "duplicate"])
        workbook.save(self.workbook_path)
        workbook.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_generates_unique_headers_and_omits_blank_cells(self):
        output, count, markdown = create_markdown(self.workbook_path)
        self.assertEqual(output.name, "Cases.md")
        self.assertEqual(count, 2)
        self.assertIn("### Name (2)\n\nduplicate", markdown)
        self.assertNotIn("### Name (2)\n\n\n", markdown)
        first_case = markdown.split("---")[0]
        self.assertNotIn("### Name (2)", first_case)

    def test_respects_conversion_start_row(self):
        _, count, markdown = create_markdown(self.workbook_path, start_row=5)
        self.assertEqual(count, 1)
        self.assertIn("### #\n\n2", markdown)
        self.assertNotIn("### #\n\n1", markdown)

    def test_rejects_start_row_at_or_above_header(self):
        with self.assertRaises(ValueError):
            create_markdown(self.workbook_path, start_row=3)


if __name__ == "__main__":
    unittest.main()
