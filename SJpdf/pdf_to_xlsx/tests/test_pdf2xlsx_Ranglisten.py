from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pdf_to_xlsx.pdf2xlsx import create_workbook, find_result_pdfs
from pdf_to_xlsx.pdf2xlsx_Startlisten import extract_start_list
from pdf_to_xlsx.pdf2xlsx_Ranglisten import (
    RANGLISTEN_COLUMNS,
    PdfResults,
    RanglistenError,
    build_ranglisten,
    competition_file_name,
)
from sj_paths import DATA_DIRECTORY


EVENT = "2026-09-12 FCM Einsiedeln HS117"


def athlete(bib, **values):
    return {"Bib": bib, "Name": f"ATHLETE {bib}", "NSA": "SUI", **values}


def jump(number, distance, total, rank=None):
    values = {f"Distance {number}": distance, f"Round Total {number}": total}
    if rank is not None:
        values[f"Round Rank {number}"] = rank
    return values


class ColumnTests(unittest.TestCase):
    def test_columns_are_fixed(self):
        letters = {get_column_letter(i): name for i, name in enumerate(RANGLISTEN_COLUMNS, 1)}
        self.assertEqual(letters["A"], "Rank")
        self.assertEqual(letters["H"], "Status")
        self.assertEqual(letters["I"], "Speed")
        self.assertEqual(letters["AB"], "Round Rank")
        self.assertEqual(letters["AC"], "Total")
        self.assertEqual(letters["AD"], "JumpID")
        self.assertEqual(letters["AE"], "JumperID")
        self.assertEqual(letters["AF"], "Date")
        self.assertEqual(letters["AG"], "Time")
        self.assertEqual(letters["AH"], "PDF")
        self.assertEqual(len(RANGLISTEN_COLUMNS), 34)


class BuildTests(unittest.TestCase):
    def competition(self):
        rl1 = [
            athlete(1, Rank=1, Total=120.0, **jump(1, 110.0, None)),
            athlete(2, Rank=2, Total=115.0, **jump(1, 108.0, None)),
            athlete(3, Rank=2, Total=115.0, **jump(1, 108.0, None)),
            athlete(4, Rank=4, Total=100.0, **jump(1, 100.0, None)),
            athlete(5, Status="DSQ", Remarks="YC SCE 4.4"),
        ]
        rl = [
            athlete(1, Rank=1, Total=240.0, **jump(1, 110.0, 120.0, 1), **jump(2, 111.0, 120.0, 1)),
            athlete(3, Rank=2, Total=230.0, **jump(1, 108.0, 115.0, 2), **jump(2, 107.0, 115.0, 2)),
            athlete(4, Rank=3, Status="Not Qualified", Total=100.0, **jump(1, 100.0, 100.0, 4)),
            athlete(2, Status="DSQ", Remarks="YC SCE 4.4.7 - Crotch control"),
            athlete(5, Status="DSQ", Remarks="YC 1st round: SCE 4.4"),
        ]
        # start lists: round 1 in bib order, final round reverse RL1 ranking
        return PdfResults("2027JP0281RL", rl, "2027JP0281RL1", rl1, [1, 2, 3, 4, 5], [4, 3, 2, 1])

    def test_rounds_jump_numbers_and_start_order(self):
        training = PdfResults(
            "2027JP0281RLT",
            [athlete(2, **jump(1, 90.0, 50.0)), athlete(1, Status="DNS")],
        )
        rows, warnings = build_ranglisten([training, self.competition()], EVENT)

        self.assertEqual(warnings, [])
        self.assertEqual(rows[0]["JumpID"], f"{EVENT} 01T")
        summary = [(r["JumpID"][-3:], r["PDF"], r["Bib"], r["Status"]) for r in rows]
        self.assertEqual(
            summary,
            [
                ("01T", "2027JP0281RLT", 1, "DNS"),
                ("01T", "2027JP0281RLT", 2, None),
                ("02W", "2027JP0281RL1", 1, None),
                ("02W", "2027JP0281RL1", 2, None),
                ("02W", "2027JP0281RL1", 3, None),
                ("02W", "2027JP0281RL1", 4, "Not Qualified"),
                ("02W", "2027JP0281RL1", 5, "DSQ"),
                # final-round start list (SLRF)
                ("03W", "2027JP0281RL", 3, None),
                ("03W", "2027JP0281RL", 2, "DSQ"),
                ("03W", "2027JP0281RL", 1, None),
            ],
        )

    def test_first_round_values_come_from_rl1(self):
        rows, _warnings = build_ranglisten([self.competition()], EVENT)
        disqualified = [r for r in rows if r["Bib"] == 2]
        self.assertEqual(disqualified[0]["Distance"], 108.0)
        self.assertEqual(disqualified[0]["Round Total"], 115.0)
        self.assertEqual(disqualified[0]["Round Rank"], 2)
        self.assertIsNone(disqualified[1]["Distance"])

    def test_total_only_in_second_competition_round(self):
        rows, _warnings = build_ranglisten([self.competition()], EVENT)
        self.assertTrue(all(r["Total"] is None for r in rows if r["JumpID"].endswith("01W")))
        winner = [r for r in rows if r["JumpID"].endswith("02W") and r["Bib"] == 1][0]
        self.assertEqual(winner["Total"], 240.0)
        self.assertEqual(winner["Distance"], 111.0)

    def test_not_qualified_must_match_rl(self):
        pdf = self.competition()
        pdf.rows[2]["Status"] = None
        with self.assertRaises(RanglistenError):
            build_ranglisten([pdf], EVENT)

    def test_missing_rl1_falls_back_with_warning(self):
        pdf = self.competition()
        pdf.first_round_rows = None
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual(len(warnings), 1)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("02W")], [3, 1])

    def test_zip_start_order_from_start_lists(self):
        pdf = self.competition()
        pdf.start_order_1 = [5, 1, 4, 2, 3]
        pdf.start_order_2 = [1, 3, 2]
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual(warnings, [])
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("01W")], [5, 1, 4, 2, 3])
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("02W")], [1, 3, 2])

    def test_missing_start_lists_fall_back_with_warnings(self):
        pdf = self.competition()
        pdf.start_order_1 = pdf.start_order_2 = None
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual(len(warnings), 2)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("01W")], [1, 2, 3, 4, 5])
        # reverse order of RL1 (bib 3 is listed after bib 2 there)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("02W")], [3, 2, 1])

    def test_athlete_missing_in_first_start_list_is_inserted_by_bib(self):
        pdf = self.competition()
        pdf.start_order_1 = [1, 2, 4, 5]
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("01W")], [1, 2, 3, 4, 5])
        self.assertEqual(len(warnings), 1)
        self.assertIn("ATHLETE 3 (bib 3)", warnings[0])
        self.assertIn("SLR1", warnings[0])

    def test_athlete_missing_in_descending_start_list_is_inserted_by_bib(self):
        pdf = self.competition()
        pdf.start_order_1 = [5, 4, 2, 1]
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("01W")], [5, 4, 3, 2, 1])
        self.assertEqual(len(warnings), 1)

    def test_starter_missing_in_final_start_list_is_inserted_by_rl1_rank(self):
        pdf = self.competition()
        pdf.start_order_2 = [3, 1]  # bib 2 (RL1 rank 2) missing
        rows, warnings = build_ranglisten([pdf], EVENT)
        self.assertEqual([r["Bib"] for r in rows if r["JumpID"].endswith("02W")], [3, 2, 1])
        self.assertEqual(len(warnings), 1)
        self.assertIn("SLRF", warnings[0])


class FileNameTests(unittest.TestCase):
    def values(self, title, gender="MEN", **extra):
        return {
            "Titel": title,
            "Gender": gender,
            "Place": "Einsiedeln (SUI)",
            "Hill Size": 117,
            "Datum": date(2026, 9, 12),
            "Start Time": "13:30",
            "Finish Time": "14:23",
            **extra,
        }

    def test_examples(self):
        cases = {
            ("FIS CUP MEN", "MEN"): "FCM",
            ("FIS SKI JUMPING WORLD CUP", "WOMEN"): "WCW",
            ("FIS SKI JUMPING GRAND PRIX", "MEN"): "GPM",
            ("FIS SUMMER GRAND PRIX", "MIXED"): "GPA",
            ("FIS CONTINENTAL CUP", "MEN"): "COCM",
            ("FIS INTER-CONTINENTAL CUP", "WOMEN"): "ICOCW",
            ("FIS Inter Continental Cup", "WOMEN"): "ICOCW",
            ("ALPEN CUP", "MEN"): "ACM",
            ("FESA CUP", "WOMEN"): "ACW",
        }
        for (title, gender), code in cases.items():
            with self.subTest(title=title):
                name = competition_file_name(
                    [self.values(title, gender), self.values(title, gender, Datum=date(2026, 9, 13))]
                )
                self.assertEqual(name, f"2026-09-13 {code} Einsiedeln HS117")

    def test_unknown_type_is_reported(self):
        with self.assertRaises(RanglistenError):
            competition_file_name([self.values("FIS SOMETHING")])

    def test_different_competitions_are_reported(self):
        with self.assertRaises(RanglistenError):
            competition_file_name(
                [self.values("FIS CUP MEN"), self.values("FIS CUP MEN", gender="WOMEN")]
            )


class WorkbookTests(unittest.TestCase):
    def test_saves_competition_copy(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "latest event.xlsx"
            source_pdf = Path(temporary_directory) / "2027JP0281RLT.pdf"
            values = FileNameTests().values("FIS CUP MEN")
            with (
                patch("pdf_to_xlsx.pdf2xlsx.extract_data", return_value=values),
                patch(
                    "pdf_to_xlsx.pdf2xlsx.extract_results",
                    return_value=[athlete(1, **jump(1, 90.0, 50.0))],
                ),
            ):
                warnings = create_workbook([source_pdf], output_file)

            self.assertEqual(warnings, [])
            copy_file = Path(temporary_directory) / "2026-09-12 FCM Einsiedeln HS117.xlsx"
            self.assertTrue(copy_file.exists())
            workbook = load_workbook(copy_file)
            self.assertEqual(workbook.sheetnames, ["Notes", "Daten", "Ranglisten", "2027JP0281RLT"])
            worksheet = workbook["Ranglisten"]
            self.assertEqual([c.value for c in worksheet[1]], list(RANGLISTEN_COLUMNS))
            self.assertEqual(worksheet["AD2"].value, "2026-09-12 FCM Einsiedeln HS117 01T")
            self.assertEqual(worksheet["AH2"].value, "2027JP0281RLT")
            self.assertEqual(workbook["Notes"]["B6"].value, copy_file.name)
            workbook.close()


    def test_rl1_column_and_sheet_stand_before_rl(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            folder = Path(temporary_directory)
            output_file = folder / "latest event.xlsx"
            for name in (
                "2027JP0281RLT.pdf", "2027JP0281RL.pdf", "2027JP0281RL1.pdf",
                "2027JP0281SLR1.pdf", "2027JP0281SLRF.pdf",
            ):
                (folder / name).touch()
            event = FileNameTests().values("FIS CUP MEN")
            values = {
                "2027JP0281RLT.pdf": {**event, "Start Time": "13:30"},
                "2027JP0281RL.pdf": {**event, "Start Time": "16:00"},
                "2027JP0281RL1.pdf": {**event, "Start Time": "16:00"},
            }
            with (
                patch("pdf_to_xlsx.pdf2xlsx.extract_data", side_effect=lambda p: values[p.name]),
                patch(
                    "pdf_to_xlsx.pdf2xlsx.extract_results",
                    return_value=[athlete(1, Rank=1, Total=50.0, **jump(1, 90.0, 50.0, 1))],
                ),
                patch(
                    "pdf_to_xlsx.pdf2xlsx.extract_start_list",
                    return_value=[{"Start Order": 1, "Bib": 1, "Name": "ATHLETE 1"}],
                ),
            ):
                create_workbook(find_result_pdfs(folder), output_file)

            workbook = load_workbook(output_file)
            self.assertEqual(
                [c.value for c in workbook["Daten"][1]],
                ["pdf", "2027JP0281RLT", "2027JP0281RL1", "2027JP0281RL"],
            )
            self.assertEqual(
                workbook.sheetnames,
                [
                    "Notes", "Daten", "Ranglisten", "2027JP0281RLT", "2027JP0281SLR1", "2027JP0281RL1",
                    "2027JP0281SLRF", "2027JP0281RL"
                ],
            )
            workbook.close()


ARCHIVE = DATA_DIRECTORY / "pdf" / "archiv"


@unittest.skipUnless(ARCHIVE.exists(), "no archived FIS PDFs in the data folder")
class RealPdfTests(unittest.TestCase):
    def test_archived_events(self):
        for event in sorted(path for path in ARCHIVE.iterdir() if path.is_dir()):
            pdfs = find_result_pdfs(event)
            if not pdfs:
                continue
            with self.subTest(event=event.name), tempfile.TemporaryDirectory() as folder:
                warnings = create_workbook(pdfs, Path(folder) / "latest event.xlsx")
                self.assertEqual(warnings, [])
                copies = [p.name for p in Path(folder).glob("*.xlsx") if p.name != "latest event.xlsx"]
                self.assertEqual(len(copies), 1)
                workbook = load_workbook(Path(folder) / "latest event.xlsx")
                worksheet = workbook["Ranglisten"]
                self.assertEqual([c.value for c in worksheet[1]], list(RANGLISTEN_COLUMNS))
                jump_numbers = []
                for row in worksheet.iter_rows(min_row=2, values_only=True):
                    if not jump_numbers or jump_numbers[-1] != row[29]:
                        jump_numbers.append(row[29])
                prefix = copies[0][: -len(".xlsx")]
                self.assertTrue(all(number.startswith(f"{prefix} ") for number in jump_numbers))
                self.assertEqual(
                    [int(number[-3:-1]) for number in jump_numbers],
                    list(range(1, len(jump_numbers) + 1)),
                )
                workbook.close()

    def test_start_lists(self):
        klingenthal = ARCHIVE / "2026-10-24 Klingenthal GPM HS140"
        if not klingenthal.exists():
            self.skipTest("Klingenthal archive missing")
        # Grand Prix: the first round starts with the highest bib.
        first = extract_start_list(klingenthal / "2027JP0278SLR1.pdf")
        self.assertEqual([row["Bib"] for row in first[:3]], [50, 49, 48])
        self.assertEqual(first[0]["Name"], "KALINICHENKO Vitaliy")
        final = extract_start_list(klingenthal / "2027JP0278SLRF.pdf")
        self.assertEqual([row["Bib"] for row in final[:4]], [46, 47, 31, 9])
        self.assertEqual(len(final), 30)
        for path in sorted(ARCHIVE.glob("*/*SLR[1F].pdf")):
            with self.subTest(pdf=path.name):
                rows = extract_start_list(path)
                self.assertEqual([row["Start Order"] for row in rows], list(range(1, len(rows) + 1)))


if __name__ == "__main__":
    unittest.main()
