from datetime import date
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from pdf_to_xlsx.pdf2xlsx import create_workbook
from pdf_to_xlsx.pdf2xlsx_Daten import extract_data_from_text
from pdf_to_xlsx.pdf2xlsx_Resultate import (
    RESULT_COLUMNS,
    ResultsParseError,
    _parse_value,
    extract_results,
    write_results_sheet,
)
from sj_paths import DATA_DIRECTORY


SAMPLE_PDF_DIRECTORY = DATA_DIRECTORY / "pdf"
PDF_NAME = re.compile(r"^\d{4}JP\d{4}(?:RLT|RLQ|RTRIA|RL)\.pdf$", re.I)


def sample_pdfs() -> list[Path]:
    if not SAMPLE_PDF_DIRECTORY.exists():
        return []
    return sorted(p for p in SAMPLE_PDF_DIRECTORY.rglob("*.pdf") if PDF_NAME.match(p.name))


class ParseValueTests(unittest.TestCase):
    def test_gate_without_change(self):
        self.assertEqual(_parse_value("Gate", "12"), {"Gate": 12, "Coach Gate": None})

    def test_ewoxx_coach_gate_change(self):
        self.assertEqual(
            _parse_value("Gate", "13➔©10"),
            {"Gate": 10, "Gate From": 13, "Coach Gate": "©"},
        )

    def test_swiss_timing_coach_gate_change(self):
        self.assertEqual(
            _parse_value("Gate", "©10→09"),
            {"Gate": 9, "Gate From": 10, "Coach Gate": "©"},
        )

    def test_coach_gate_without_arrow(self):
        self.assertEqual(_parse_value("Gate", "©5"), {"Gate": 5, "Coach Gate": "©"})

    def test_distance_with_fall(self):
        self.assertEqual(_parse_value("Distance", "*106.5"), {"Distance": 106.5, "Fall": "*"})
        self.assertEqual(_parse_value("Distance", "106.5"), {"Distance": 106.5})

    def test_shared_rank_is_numeric(self):
        self.assertEqual(_parse_value("Speed Rank", "=22."), {"Speed Rank": 22})
        self.assertEqual(_parse_value("Round Rank", "3."), {"Round Rank": 3})

    def test_unexpected_text_is_rejected(self):
        self.assertIsNone(_parse_value("Wind", "fis-ski.com"))
        self.assertIsNone(_parse_value("Distance", "12a"))


class ResultsSheetTests(unittest.TestCase):
    def test_sheet_uses_fixed_columns(self):
        workbook = Workbook()
        write_results_sheet(
            workbook,
            "2027JP0281RL",
            [
                {"Rank": 1, "Bib": 65, "Name": "STEINBERGER Jakob", "Distance 1": 113.0,
                 "Wind 1": 0.32, "Date of Birth": date(2005, 8, 18)},
                {"Bib": 31, "Name": "CRAWFORD Dyllon", "Status": "DNS"},
            ],
        )
        worksheet = workbook["2027JP0281RL"]
        header = [cell.value for cell in worksheet[1]]
        self.assertEqual(header, list(RESULT_COLUMNS))
        column = {name: index + 1 for index, name in enumerate(header)}
        self.assertEqual(worksheet.cell(row=2, column=column["Distance 1"]).value, 113.0)
        self.assertEqual(worksheet.cell(row=2, column=column["Wind 1"]).number_format, "0.00")
        self.assertIsNone(worksheet.cell(row=3, column=column["Rank"]).value)
        self.assertEqual(worksheet.cell(row=3, column=column["Status"]).value, "DNS")

    def test_unreadable_results_get_error_sheet_and_warning(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "results.xlsx"
            source_pdf = Path(temporary_directory) / "2027JP0281RL.pdf"
            with (
                patch(
                    "pdf_to_xlsx.pdf2xlsx.extract_data",
                    return_value=extract_data_from_text("FIS CUP MEN\n"),
                ),
                patch(
                    "pdf_to_xlsx.pdf2xlsx.extract_results",
                    side_effect=ResultsParseError("2027JP0281RL.pdf: result check failed"),
                ),
            ):
                warnings = create_workbook([source_pdf], output_file)

            self.assertIn("2027JP0281RL.pdf: result check failed", warnings)
            self.assertTrue(warnings[-1].startswith("Ranglisten: not created"))
            workbook = load_workbook(output_file)
            self.assertEqual(workbook.sheetnames, ["Notes", "Daten", "Ranglisten", "2027JP0281RL"])
            self.assertIn("not created", workbook["Ranglisten"]["A1"].value)
            self.assertIn("result check failed", workbook["2027JP0281RL"]["A1"].value)
            workbook.close()


@unittest.skipUnless(sample_pdfs(), "no FIS sample PDFs in the data folder")
class RealPdfTests(unittest.TestCase):
    """All PDFs in data/pdf (including archiv) must pass every result check."""

    def test_all_sample_pdfs_are_read(self):
        for pdf_path in sample_pdfs():
            with self.subTest(pdf=str(pdf_path.relative_to(SAMPLE_PDF_DIRECTORY))):
                rows = extract_results(pdf_path)
                self.assertTrue(rows)
                for row in rows:
                    self.assertTrue(set(row) <= set(RESULT_COLUMNS), set(row) - set(RESULT_COLUMNS))
                    self.assertTrue(row.get("Name"))
                    self.assertTrue(row.get("NSA"))


if __name__ == "__main__":
    unittest.main()
