from datetime import date
from io import StringIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from pdf_to_xlsx.pdf2xlsx import (
    create_workbook,
    default_pdf_directory,
    find_result_pdfs,
    main,
)
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

    def test_hill_size_without_hs_label_and_with_hyphen(self):
        values = extract_data_from_text(FIS_CUP_TEXT, "2027JP0281RL")

        self.assertEqual(values["Hill Size"], 117)
        self.assertEqual(values["HS 95%"], 111)

    def test_hill_size_with_original_label_format(self):
        text = FIS_CUP_TEXT.replace(
            "HILL-SIZE / (95%)",
            "HILL SIZE HS / (95%)",
        )
        values = extract_data_from_text(text, "2027JP0281RL")

        self.assertEqual(values["Hill Size"], 117)
        self.assertEqual(values["HS 95%"], 111)

    def test_hill_size_label_can_vary_between_hill_and_95_percent(self):
        text = FIS_CUP_TEXT.replace(
            "HILL-SIZE / (95%)",
            "HILL DIMENSION / (95%)",
        )
        values = extract_data_from_text(text, "2027JP0281RL")

        self.assertEqual(values["Hill Size"], 117)
        self.assertEqual(values["HS 95%"], 111)

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

    def test_only_supported_pdfs_in_input_folder_are_found(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            pdf_directory = Path(temporary_directory)
            event_directory = pdf_directory / "2026-09-13 Einsiedeln FCM"
            event_directory.mkdir()
            (pdf_directory / "2027JP0373RL.pdf").touch()
            (event_directory / "2027JP0372RL.pdf").touch()
            (event_directory / "2027JP0372RL1.pdf").touch()

            self.assertEqual(
                [path.name for path in find_result_pdfs(pdf_directory)],
                ["2027JP0373RL.pdf"],
            )

    def test_default_input_prefers_synced_project_when_both_have_results(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            project_root = root / "checkout" / "SJpdf"
            project_pdf_directory = project_root / "data" / "pdf"
            project_pdf_directory.mkdir(parents=True)
            (project_pdf_directory / "2027JP0373RL.pdf").touch()
            synced_pdf_directory = (
                root / "CLOUD-DANIEL" / "SJA" / "SJpdf" / "data" / "pdf"
            )
            archived_event = synced_pdf_directory / "archiv" / "event"
            archived_event.mkdir(parents=True)
            (archived_event / "2027JP0372RL.pdf").touch()
            (synced_pdf_directory / "2027JP0374RL.pdf").touch()

            self.assertEqual(
                default_pdf_directory(project_root, root),
                synced_pdf_directory,
            )

    def test_default_input_falls_back_to_checkout_when_synced_has_no_results(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            project_pdf_directory = root / "checkout" / "SJpdf" / "data" / "pdf"
            project_pdf_directory.mkdir(parents=True)
            (project_pdf_directory / "2027JP0372RL.pdf").touch()
            synced_pdf_directory = (
                root / "CLOUD-DANIEL" / "SJA" / "SJpdf" / "data" / "pdf"
            )
            synced_pdf_directory.mkdir(parents=True)

            self.assertEqual(
                default_pdf_directory(root / "checkout" / "SJpdf", root),
                project_pdf_directory,
            )

    def test_workbook_uses_pdf_columns_and_date_format(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "results.xlsx"
            workbook = Workbook()
            workbook.active.title = "Daten"
            workbook["Daten"]["A1"] = "existing field"
            workbook["Daten"]["B1"] = "old PDF"
            workbook.create_sheet("Old results")
            workbook.save(output_file)
            workbook.close()

            source_pdf = Path(temporary_directory) / "2027JP0276RLT.pdf"

            from unittest.mock import patch

            with patch(
                "pdf_to_xlsx.pdf2xlsx.extract_data",
                return_value=extract_data_from_text(SAMPLE_TEXT),
            ):
                create_workbook([source_pdf], output_file)

            workbook = load_workbook(output_file, data_only=True)
            self.assertEqual(workbook.sheetnames, ["Daten"])
            worksheet = workbook["Daten"]
            self.assertEqual(worksheet["A1"].value, "pdf")
            self.assertEqual(worksheet["B1"].value, "2027JP0276RLT")
            self.assertEqual(worksheet["A2"].value, "Titel")
            self.assertEqual(worksheet["B2"].value, "FIS SKI JUMPING GRAND PRIX")
            self.assertEqual(worksheet["A4"].value, "Datum")
            self.assertEqual(worksheet["B4"].number_format, "yyyy-mm-dd")
            self.assertEqual(worksheet.max_column, 2)
            workbook.close()

    def test_no_matching_pdfs_still_clears_old_sheets_and_columns(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "results.xlsx"
            workbook = Workbook()
            workbook.active.title = "Daten"
            workbook["Daten"]["A1"] = "keep"
            workbook["Daten"]["B1"] = "remove"
            workbook.create_sheet("Old results")
            workbook.save(output_file)
            workbook.close()

            with self.assertRaises(FileNotFoundError):
                create_workbook([], output_file)

            workbook = load_workbook(output_file, data_only=True)
            self.assertEqual(workbook.sheetnames, ["Daten"])
            self.assertEqual(workbook["Daten"]["A1"].value, "keep")
            self.assertEqual(workbook["Daten"].max_column, 1)
            workbook.close()

    def test_unavailable_fields_have_light_red_background(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "results.xlsx"
            source_pdf = Path(temporary_directory) / "2027JP0276RLT.pdf"

            with patch(
                "pdf_to_xlsx.pdf2xlsx.extract_data",
                return_value={"Titel": "FIS SKI JUMPING GRAND PRIX"},
            ):
                create_workbook([source_pdf], output_file)

            workbook = load_workbook(output_file)
            worksheet = workbook["Daten"]
            self.assertEqual(worksheet["B2"].value, "FIS SKI JUMPING GRAND PRIX")
            self.assertIsNone(worksheet["B2"].fill.fill_type)
            self.assertIsNone(worksheet["B3"].value)
            self.assertEqual(worksheet["B3"].fill.fill_type, "solid")
            self.assertEqual(worksheet["B3"].fill.fgColor.rgb, "FFFFC7CE")
            workbook.close()

    def test_startup_clears_workbook_before_input_directory_check(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            output_file = root / "results.xlsx"
            workbook = Workbook()
            workbook.active.title = "Daten"
            workbook["Daten"]["A1"] = "keep"
            workbook["Daten"]["B1"] = "remove"
            workbook.create_sheet("Old results")
            workbook.save(output_file)
            workbook.close()
            error_output = StringIO()

            with patch("pdf_to_xlsx.pdf2xlsx.sys.stderr", error_output):
                result = main(
                    ["--pdf-dir", str(root / "missing"), "--output", str(output_file)]
                )

            self.assertEqual(result, 1)
            workbook = load_workbook(output_file, data_only=True)
            self.assertEqual(workbook.sheetnames, ["Daten"])
            self.assertEqual(workbook["Daten"]["A1"].value, "keep")
            self.assertEqual(workbook["Daten"].max_column, 1)
            workbook.close()

    def test_locked_output_reports_how_to_fix_permission_error(self):
        error_output = StringIO()
        with (
            patch("pdf_to_xlsx.pdf2xlsx.find_result_pdfs", return_value=[Path("sample.pdf")]),
            patch(
                "pdf_to_xlsx.pdf2xlsx.create_workbook",
                side_effect=PermissionError(13, "Permission denied"),
            ),
            patch("pdf_to_xlsx.pdf2xlsx.sys.stderr", error_output),
        ):
            result = main(["--pdf-dir", ".", "--output", "locked.xlsx"])

        self.assertEqual(result, 1)
        self.assertIn("Close the workbook in Excel", error_output.getvalue())
        self.assertIn("locked.xlsx", error_output.getvalue())


if __name__ == "__main__":
    unittest.main()
