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
    _State,
    _check_ranks,
    _parse_value,
    _validate,
    extract_results,
    looks_like_athlete,
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

    def test_older_coach_gate_notations(self):
        # Swiss Timing 2024: '15->©14'; local organisers: '©11/13' (gate 11, before 13)
        self.assertEqual(
            _parse_value("Gate", "15->©14"),
            {"Gate": 14, "Gate From": 15, "Coach Gate": "©"},
        )
        self.assertEqual(
            _parse_value("Gate", "©11/13"),
            {"Gate": 11, "Gate From": 13, "Coach Gate": "©"},
        )

    def test_name_without_bold_print(self):
        for name in ("CHERVET Emma", "D'AGOSTINO Ava", "MITU-COSCA Alessia", "DEL MORO Lucrezia"):
            self.assertTrue(looks_like_athlete(name), name)
        for club in ("CS CHAMONIX", "Aioi Nissay Dowa Insurance Co., Ltd", "31 DEC 2003"):
            self.assertFalse(looks_like_athlete(club), club)

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


class RankTests(unittest.TestCase):
    def test_dense_ranks_of_winspringen_training(self):
        rows = [{"Name": n, "Speed 1": v, "Speed Rank 1": r}
                for n, v, r in (("A", 88.0, 1), ("B", 88.0, 1), ("C", 87.5, 2))]
        errors = []
        _check_ranks(rows, "Speed 1", "Speed Rank 1", errors, dense=True)
        self.assertEqual(errors, [])
        _check_ranks(rows, "Speed 1", "Speed Rank 1", errors)
        self.assertEqual(errors, ["C: Speed Rank 1 2 expected 3"])

    def test_ko_competition_ranks_winners_before_losers(self):
        # Four Hills: the loser of a pair (rank 3) may have more points than a winner.
        state = _State("test.pdf", layouts={"competition"})
        state.rows = [
            {"Name": name, "Rank": rank, "Total": 110.0 + wind, "Distance 1": 125.0,
             "Distance Points 1": 60.0, "Judge Points 1": 50.0, "Wind Points 1": wind, **extra}
            for name, rank, wind, extra in (
                ("A", 1, 30.0, {}), ("B", 2, 10.0, {}), ("C", 3, 20.0, {"Status": "Not Qualified"}),
            )
        ]
        _validate(state, "Men Individual KO Large Hill\nK-POINT 125m\nMETER VALUE 1.8\n")
        with self.assertRaises(ResultsParseError):
            _validate(state, "Men Large Hill Individual\nK-POINT 125m\nMETER VALUE 1.8\n")


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


class LegacyLayoutTests(unittest.TestCase):
    """Swiss Timing until 2023: old header labels renamed to the current ones."""

    def test_competition_header(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import Span, _legacy_labels

        def span(text, x0, x1, y):
            return Span(text, x0, x1, y - 4, y + 4, True)

        spans = [
            span("Rk", 22, 32, 239), span("Bib", 42, 54, 239), span("Name", 66, 87, 227),
            span("Nation", 180, 204, 227), span("Speed", 212, 232, 239), span("Dist", 240, 252, 239),
            span("DP", 266, 275, 239), *(span(j, 289 + 18 * i, 293 + 18 * i, 239) for i, j in enumerate("ABCDE")),
            span("JP", 380, 388, 239), span("Gate", 399, 417, 239), span("Pts", 429, 441, 239),
            span("m/s", 457, 473, 239), span("Pts", 486, 498, 239), span("Round", 510, 534, 227),
            span("Points", 510, 534, 237), span("Total", 549, 567, 239),
        ]
        renamed = _legacy_labels(spans)
        row = [s.text for s in sorted(renamed, key=lambda s: s.x0) if abs(s.y - 239) < 1 and s.x0 > 210]
        self.assertEqual(row, ["[km/h]", "[m]", "POINTS", "A", "B", "C", "D", "E", "POINTS",
                               "GATE", "POINTS", "[m/s]", "POINTS", "POINTS", "RANK"])
        self.assertIn("TOTAL", [s.text for s in renamed])
        self.assertEqual({s.text for s in renamed if s.x0 < 60}, {"RANK", "BIB"})

    def test_current_layout_unchanged(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import Span, _legacy_labels

        spans = [Span("[km/h]", 200, 220, 100, 108, True), Span("Speed", 200, 220, 90, 98, True)]
        self.assertIs(_legacy_labels(spans), spans)
