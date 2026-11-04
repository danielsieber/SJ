from datetime import date
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

from pdf_to_xlsx.pdf2xlsx import create_workbook, find_result_pdfs
from pdf_to_xlsx.pdf2xlsx_1Daten import PdfParseError, extract_data_from_text


SAMPLE_TEXT = """\
FIS SKI JUMPING GRAND PRIX
7th Grand Prix Competition                                             FRI 23 OCT 2026
Klingenthal (GER)
Men Individual Large Hill                                                          Start Time:        17:01
Finish Time:       18:37
Results Official Training
Hill Data
HILL SIZE HS / (95%)                                         140m (133m)    Gate Factor                                                            7.67 points / m
K-POINT                                                       125m    Wind Factors - head / tail                                    10.80 / 16.20 points / m/s
METER VALUE                                                              1.8 points / m
"""

FIS_CUP_TEXT = """\
FIS CUP MEN
7th FIS Cup Competition                                                                   SAT 12 SEP 2026
Einsiedeln (SUI)                                                                                                                                                                    Start Time 16:00
LARGE HILL INDIVIDUAL                                                                                                 End Time 16:48
INTERMEDIATE RESULTS 1ST ROUND
Hill Data / Calculation
HILL-SIZE / (95%)                                         117m (111m)                                 GATE FACTOR                                                    6.32 points per m
K-POINT                                                    105m                             WIND FACTORS - HEAD / TAIL                            7.29 / 10.94 points per m/s
METER VALUE                                                        1.8 points per m
"""


class ExtractDataTests(unittest.TestCase):
    def test_extracts_general_data_and_hill_values(self):
        values = extract_data_from_text(SAMPLE_TEXT, "2027JP0276RLT")

        self.assertEqual(values["Titel"], "FIS SKI JUMPING GRAND PRIX")
        self.assertEqual(values["Typ"], "GRAND PRIX")
        self.assertEqual(values["Datum"], date(2026, 10, 23))
        self.assertEqual(values["Start Time"], "17:01")
        self.assertEqual(values["Finish Time"], "18:37")
        self.assertEqual(values["Competition"], "7th Grand Prix Competition")
        self.assertEqual(values["Place"], "Klingenthal (GER)")
        self.assertEqual(values["Hill"], "Individual Large Hill")
        self.assertEqual(values["Gender"], "MEN")
        self.assertEqual(values["Rangliste"], "Results Official Training")
        self.assertEqual(values["Hill Size"], 140)
        self.assertEqual(values["HS 95%"], 133)
        self.assertEqual(values["K-Punkt"], 125)
        self.assertEqual(values["Meter Value"], 1.8)
        self.assertEqual(values["Gate Factor"], 7.67)
        self.assertEqual(values["HWC"], 10.8)
        self.assertEqual(values["TWC"], 16.2)

    def test_type_removes_sport_words_and_gender(self):
        text = SAMPLE_TEXT.replace(
            "FIS SKI JUMPING GRAND PRIX",
            "FIS CONTINENTAL CUP SKI JUMPING MEN",
        )
        self.assertEqual(
            extract_data_from_text(text)["Typ"],
            "CONTINENTAL CUP",
        )

    def test_fis_cup_header_without_colons(self):
        values = extract_data_from_text(FIS_CUP_TEXT, "2027JP0281RL")

        self.assertEqual(values["Titel"], "FIS CUP MEN")
        self.assertEqual(values["Typ"], "FIS CUP")
        self.assertEqual(values["Datum"], date(2026, 9, 12))
        self.assertEqual(values["Start Time"], "16:00")
        self.assertEqual(values["Finish Time"], "16:48")
        self.assertEqual(values["Competition"], "7th FIS Cup Competition")
        self.assertEqual(values["Place"], "Einsiedeln (SUI)")
        self.assertEqual(values["Hill"], "LARGE HILL INDIVIDUAL")
        self.assertEqual(values["Gender"], "MEN")
        self.assertEqual(values["Rangliste"], "INTERMEDIATE RESULTS 1ST ROUND")

    def test_missing_optional_fields_are_blank(self):
        values = extract_data_from_text(
            "FIS SKI JUMPING GRAND PRIX\nResults Qualification\n"
        )
        self.assertIsNone(values["Datum"])
        self.assertIsNone(values["Hill Size"])
        self.assertIsNone(values["Gender"])

    def test_missing_title_is_reported(self):
        with self.assertRaises(PdfParseError):
            extract_data_from_text("No result data")


class WorkbookTests(unittest.TestCase):
    def test_only_supported_pdf_suffixes_are_selected(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            pdf_directory = Path(temporary_directory)
            for name in (
                "2027JP0372RL.pdf",
                "2027JP0372RLT.pdf",
                "2027JP0372RLQ.pdf",
                "2027JP0372RTRIA.pdf",
                "2027JP0372SLQ.pdf",
                "otherRL.pdf",
            ):
                (pdf_directory / name).touch()

            self.assertEqual(
                [path.name for path in find_result_pdfs(pdf_directory)],
                [
                    "2027JP0372RL.pdf",
                    "2027JP0372RLQ.pdf",
                    "2027JP0372RLT.pdf",
                    "2027JP0372RTRIA.pdf",
                ],
            )

    def test_workbook_uses_pdf_columns_and_date_format(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "results.xlsx"
            source_pdf = Path(temporary_directory) / "2027JP0276RLT.pdf"

            from unittest.mock import patch

            with patch(
                "pdf_to_xlsx.pdf2xlsx.extract_data",
                return_value=extract_data_from_text(SAMPLE_TEXT),
            ):
                create_workbook([source_pdf], output_file)

            workbook = load_workbook(output_file, data_only=True)
            worksheet = workbook["Daten"]
            self.assertEqual(worksheet["A1"].value, "pdf")
            self.assertEqual(worksheet["B1"].value, "2027JP0276RLT")
            self.assertEqual(worksheet["A2"].value, "Titel")
            self.assertEqual(worksheet["B2"].value, "FIS SKI JUMPING GRAND PRIX")
            self.assertEqual(worksheet["A4"].value, "Datum")
            self.assertEqual(worksheet["B4"].number_format, "yyyy-mm-dd")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
