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
        # Grand Prix 2025: not qualified for the final round, ranked apart too.
        _validate(state, "Men Large Hill Individual\nK-POINT 125m\nMETER VALUE 1.8\n")
        del state.rows[2]["Status"]
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


class HillAndLabelTests(unittest.TestCase):
    def test_hill_from_distances_when_not_printed(self):
        # Ruhpolding: '40.0 m Schanze', 2.4 points per metre, no K-point in the PDF.
        from pdf_to_xlsx.pdf2xlsx_Resultate import _fit_hill

        rows = [{"Distance 1": 38.5, "Distance Points 1": 56.4, "Distance 2": 41.0, "Distance Points 2": 62.4},
                {"Distance 1": 35.5, "Distance Points 1": 49.2}]
        self.assertEqual(_fit_hill(rows), (40.0, 2.4))
        rows[1]["Distance Points 1"] = 50.0  # does not fit
        self.assertIsNone(_fit_hill(rows))

    def test_labels_printed_as_one_text_are_split(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import Span, _split_labels

        spans = _split_labels([Span("[m/s] Points", 100, 160, 10, 18, True), Span("Judges Marks", 0, 40, 10, 18, True)])
        self.assertEqual([s.text for s in spans], ["[m/s]", "Points", "Judges Marks"])
        self.assertLess(spans[0].x1, spans[1].x0)

    def test_footer_texts_are_no_values(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import FOOTER_PATTERN, NO_ROW_PATTERN

        for text in ("FIS Data Provider", "Data & Timing Sponsor", "Page 1/2"):
            self.assertTrue(FOOTER_PATTERN.match(text), text)
        self.assertTrue(NO_ROW_PATTERN.match("Reason for disqualification:"))

    def test_team_name_as_country_in_capitals(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import _identity, _is_team_name, Span

        for text in ("AUT  -  Austria", "AUSTRIA", "SLOVENIA II", "UNITED STATES OF AMERICA"):
            self.assertTrue(_is_team_name(text), text)
        for text in ("USA", "HAAGEN David"):
            self.assertFalse(_is_team_name(text), text)
        team = _identity([Span("SLOVENIA II", 66, 120, 240, 248, True)])
        self.assertEqual(team["NSA"], "SLO")

    def test_bib_and_name_in_one_text(self):
        from pdf_to_xlsx.pdf2xlsx_Resultate import _identity, Span

        row = _identity([Span("1", 42, 47, 246, 254, False), Span("49 KOBAYASHI Ryoyu", 55, 136, 246, 254, False),
                         Span("JPN", 219, 234, 246, 254, False)])
        self.assertEqual((row["Rank"], row["Bib"], row["Name"], row["NSA"]), (1, 49, "KOBAYASHI Ryoyu", "JPN"))
        row = _identity([Span("16", 38, 47, 246, 254, False), Span("21 TROFIMOV RS", 55, 122, 246, 254, False),
                         Span("ROC", 180, 197, 246, 254, False)])
        self.assertEqual((row["Bib"], row["Name"]), (21, "TROFIMOV RS"))

    def test_round_ranks_without_the_jumpers_out_of_the_next_round(self):
        # Swiss Timing until 2023: DSQ in round 2, left out of the round 1 ranks.
        state = _State("test.pdf", layouts={"competition"})
        state.rows = [
            # (no jumps: only the ranks are checked)
            {"Name": name, "Status": "DSQ", "Round Total 1": first, "Round Rank 1": first_rank,
             **({"Round Total 2": total - first} if rank < 3 else {})}
            for name, rank, total, first, first_rank in (
                ("A", 1, 200.0, 100.0, 1), ("B", 2, 150.0, 70.0, 2), ("C", 3, 80.0, 80.0, 2),
            )
        ]
        _validate(state, "K-POINT 95m\nMETER VALUE 2.0\n")

    def test_rank_qualified_by_the_fall_rule(self):
        self.assertEqual(_parse_value("Round Rank", "f33."),
                         {"Round Rank": 33, "Remarks": "f: add. qualified by rule 4.2 (Fall)"})

    def test_unprinted_gate_compensation_in_half_metres(self):
        # Klingenthal 2019: no gate printed, the gate lowered during round 1.
        state = _State("test.pdf", layouts={"competition"})
        state.wind_column = False
        state.rows = [
            {"Name": name, "Distance 1": 77.0, "Distance Points 1": 60.0, "Judge Points 1": 50.0,
             "Round Total 1": total, "Total": total, "Rank": rank}
            for name, total, rank in (("A", 115.6, 1), ("B", 112.8, 2), ("C", 110.0, 3))
        ]
        text = "K-POINT 77m\nMETER VALUE 2.0\nGate Faktor: 5,61 Pkt/m\n"
        _validate(state, text)
        state.rows[0]["Round Total 1"] = state.rows[0]["Total"] = 114.0  # not a gate compensation
        with self.assertRaises(ResultsParseError):
            _validate(state, text)
