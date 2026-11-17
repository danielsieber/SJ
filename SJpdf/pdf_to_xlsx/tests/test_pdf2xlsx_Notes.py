from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from pdf_to_xlsx.pdf2xlsx import create_workbook, find_result_pdfs
from pdf_to_xlsx.pdf2xlsx_Notes import FILE_FORMULA, FOLDER_FORMULA, Note


VALUES = {
    "Titel": "FIS CUP MEN",
    "Gender": "MEN",
    "Place": "Einsiedeln (SUI)",
    "Hill Size": 117,
    "Datum": date(2026, 9, 12),
    "Start Time": "13:30",
}
ROWS = [{"Bib": 1, "Name": "ATHLETE 1", "NSA": "SUI", "Distance 1": 90.0}]


def run(folder: Path, side_effect=None):
    output_file = folder / "out" / "latest event.xlsx"
    with (
        patch("pdf_to_xlsx.pdf2xlsx.extract_data", return_value=VALUES),
        patch("pdf_to_xlsx.pdf2xlsx.extract_results", return_value=ROWS, side_effect=side_effect),
    ):
        warnings = create_workbook(find_result_pdfs(folder), output_file)
    return warnings, output_file


def notes_rows(worksheet):
    return [tuple(v for v in row if v is not None) for row in worksheet.iter_rows(values_only=True)]


class NotesTests(unittest.TestCase):
    def test_notes_is_active_and_shows_paths_and_links(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            (folder / "2027JP0281RLT.pdf").touch()
            (folder / "2027JP0281RTRIA.pdf").touch()
            (folder / "pdf_links.csv").write_text(
                "pdf,url\n2027JP0281RLT.pdf,https://www.fis-ski.com/a.pdf\n", encoding="utf-8"
            )
            _warnings, output_file = run(folder)

            workbook = load_workbook(output_file)
            self.assertEqual(workbook.active.title, "Notes")
            self.assertTrue(workbook.calculation.fullCalcOnLoad)
            self.assertEqual(
                [ws.title for ws in workbook.worksheets if ws.sheet_view.tabSelected], ["Notes"]
            )
            rows = notes_rows(workbook["Notes"])
            self.assertIn(("xlsx-Ordner", FOLDER_FORMULA), rows)
            self.assertIn(("xlsx-Datei", FILE_FORMULA), rows)
            self.assertIn(("PDF-Ordner", str(folder)), rows)
            self.assertIn(
                ("2027JP0281RLT", str(folder / "2027JP0281RLT.pdf"), "https://www.fis-ski.com/a.pdf"),
                rows,
            )
            self.assertIn(("Info", "Kein FIS-Link in pdf_links.csv für 2027JP0281RTRIA."), rows)

            daten = workbook["Daten"]
            linked = {c.value: c.hyperlink.target for c in daten[1] if c.hyperlink}
            self.assertEqual(linked, {"2027JP0281RLT": "https://www.fis-ski.com/a.pdf"})
            workbook.close()

    def test_warnings_appear_in_notes(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import ResultsParseError

        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            (folder / "2027JP0281RLT.pdf").touch()
            warnings, output_file = run(
                folder, side_effect=ResultsParseError("2027JP0281RLT.pdf page 1: broken")
            )
            workbook = load_workbook(output_file)
            rows = notes_rows(workbook["Notes"])
            self.assertIn(("Warnung", "2027JP0281RLT", "2027JP0281RLT.pdf page 1: broken"), rows)
            self.assertEqual(len(warnings), 2)
            workbook.close()

    def test_source_of_note(self):
        self.assertEqual(Note("Warnung", "2027JP0281RL.pdf page 2: x").source, "2027JP0281RL")
        self.assertEqual(Note("Warnung", "Ranglisten: x").source, "Ranglisten")
        self.assertEqual(Note("Info", "Kopie gespeichert").source, "")


if __name__ == "__main__":
    unittest.main()
