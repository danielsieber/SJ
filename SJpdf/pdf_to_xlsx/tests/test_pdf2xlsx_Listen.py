from datetime import date, time
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from pdf_to_xlsx.pdf2xlsx import create_workbook
from pdf_to_xlsx.pdf2xlsx_Listen import (
    Jumper,
    ListenError,
    SjListen,
    hill_name,
    proposed_jumper_id,
    read_sj_listen,
)
from pdf_to_xlsx.pdf2xlsx_Notes import WARNING_FILL


def write_listen(path: Path, jumpers=(), hills=()) -> Path:
    """A small SJ-Listen workbook with the layout of SJ-Listen.xlsm."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Liste_Springer_ID"
    sheet["A1"] = "Springerverzeichnis"
    sheet.append([])
    for _ in range(7):
        sheet.append([])
    sheet.append(["JumperID", "Name", "NSA", "Gender"])
    for jumper in jumpers:
        sheet.append(list(jumper))
    sheet = workbook.create_sheet("Liste_Schanzen_ID")
    sheet["A1"] = "Schanzenverzeichnis"
    for _ in range(8):
        sheet.append([])
    sheet.append(["SchanzenID", "HS", "95%", "K", "Pts./m", "Headwind \nPts./m/s",
                  "Tailwind \nPts./m/s", "Gate \nFactor \nPts./m", "Pts / Gate \nfür DGW"])
    for hill, dgw in hills:
        sheet.append([hill, None, None, None, None, None, None, None, dgw])
    workbook.save(path)
    workbook.close()
    return path


def daten_values():
    return {
        "Titel": "FIS CUP MEN",
        "Gender": "MEN",
        "Place": "Einsiedeln (SUI)",
        "Hill Size": 117,
        "TWC": 10.94,
        "Datum": date(2026, 9, 12),
        "DG 1": "1st Round",
        "Start Time 1": "13:30",
        "End Time 1": "14:23",
    }


class RuleTests(unittest.TestCase):
    def test_proposed_jumper_id(self):
        self.assertEqual(proposed_jumper_id("ACHLEITNER Alexander", "AUT"), "AUT ACHLEI ALEXAN")
        self.assertEqual(proposed_jumper_id("HAGEN Ida Marie", "NOR"), "NOR HAGEN IDA")
        self.assertEqual(proposed_jumper_id("PERCL SERUCNIK Lovro", "SLO"), "SLO PERCL SERUCN")

    def test_rule_id_needs_the_same_surname_else_name_lookup(self):
        listen = SjListen(
            jumper_ids={("del bianco lucrezia", "ita"): "ITA DEL LUCREZ"},
            jumpers=[
                Jumper("AUT STEINB SIMON", "STEINBERGER Simon", "AUT", 8091),
                Jumper("NOR BROBEC HANNAH", "BROBECK Hannah", "NOR", 220607),
                Jumper("ITA DEL LUCREZ", "DEL BIANCO Lucrezia", "ITA", 220365),
            ],
        )
        # second first name in the PDF: same rule ID, same surname
        self.assertEqual(listen.resolve("BROBECK Hannah Christina", "NOR"), ("NOR BROBEC HANNAH", None))
        # special ID: rule ID 'ITA DEL BIANCO' not listed, found by name
        self.assertEqual(listen.resolve("DEL BIANCO Lucrezia", "ITA"), ("ITA DEL LUCREZ", None))
        # rule ID of another jumper (other surname): not taken
        jumper_id, problem = listen.resolve("STEINBERG Simon", "AUT")
        self.assertIsNone(jumper_id)
        self.assertIn("STEINBERGER Simon (AUT)", problem)
        self.assertEqual(
            listen.resolve("MUSTER Hans", "SUI"),
            (None, "JumperID 'SUI MUSTER HANS' fehlt in Liste_Springer_ID"),
        )

    def test_hill_name(self):
        self.assertEqual(hill_name("Einsiedeln (SUI)", 117), "Einsiedeln HS117")
        self.assertEqual(hill_name("Einsiedeln (SUI)", 117.0), "Einsiedeln HS117")
        self.assertIsNone(hill_name(None, 117))


class ReadTests(unittest.TestCase):
    def test_lookup_ignores_case_and_spaces(self):
        with tempfile.TemporaryDirectory() as folder:
            path = write_listen(
                Path(folder) / "SJ-Listen.xlsx",
                jumpers=[("SUI HOESLI LION", "HOESLI Lion", "SUI", "M")],
                hills=[("Einsiedeln HS117", 4.4167)],
            )
            listen = read_sj_listen(path)
        self.assertEqual(listen.jumper_id("HOESLI  Lion", "sui"), "SUI HOESLI LION")
        self.assertIsNone(listen.jumper_id("HOESLI Lion", "AUT"))
        self.assertEqual(listen.dgw("einsiedeln hs117"), 4.4167)

    def test_missing_file_or_sheet_is_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(ListenError):
                read_sj_listen(Path(folder) / "SJ-Listen.xlsx")
            path = Path(folder) / "other.xlsx"
            Workbook().save(path)
            with self.assertRaisesRegex(ListenError, "Liste_Springer_ID"):
                read_sj_listen(path)


class WorkbookTests(unittest.TestCase):
    def convert(self, folder, listen_file):
        output_file = Path(folder) / "latest event.xlsx"
        rows = [
            {"Bib": 1, "Name": "HOESLI Lion", "NSA": "SUI", "Distance 1": 90.0, "Round Total 1": 50.0},
            {"Bib": 2, "Name": "MUSTER Hans Peter", "NSA": "SUI", "Distance 1": 80.0, "Round Total 1": 40.0},
            {"Bib": 3, "Name": "MUSTER Fritz", "NSA": "SUI", "Status": "DNS"},
            {"Bib": 4, "Name": "HOESLI Lion", "NSA": "SUI", "Distance 1": 70.0, "Round Total 1": 30.0},
        ]
        with (
            patch("pdf_to_xlsx.pdf2xlsx.extract_data", return_value=daten_values()),
            patch("pdf_to_xlsx.pdf2xlsx.extract_results", return_value=rows),
        ):
            warnings = create_workbook(
                [Path(folder) / "2027JP0281RLT.pdf"], output_file, listen_file
            )
        return warnings, load_workbook(output_file)

    def test_jumper_ids_and_dgw_are_taken_from_the_lists(self):
        with tempfile.TemporaryDirectory() as folder:
            listen_file = write_listen(
                Path(folder) / "SJ-Listen.xlsx",
                jumpers=[
                    ("SUI HOESLI LION", "HOESLI Lion", "SUI", "M"),
                    ("SUI MUSTER HANS", "MUSTER Hans Peter", "SUI", "M"),
                    ("SUI MUSTER FRITZ", "MUSTER Fritz", "SUI", "M"),
                ],
                hills=[("Einsiedeln HS117", 4.4167)],
            )
            warnings, workbook = self.convert(folder, listen_file)
            self.assertEqual(warnings, [])
            self.assertEqual(workbook.sheetnames[0], "Notes")
            ranglisten = workbook["Ranglisten"]
            self.assertEqual(ranglisten["AE1"].value, "JumperID")
            self.assertEqual(
                [ranglisten.cell(row=row, column=31).value for row in (2, 3, 4)],
                ["SUI HOESLI LION", "SUI MUSTER HANS", "SUI MUSTER FRITZ"],
            )
            daten = workbook["Daten"]
            labels = [cell.value for cell in daten["A"]]
            row = labels.index("Pts / Gate DGW") + 1
            self.assertEqual(labels[row - 2], "TWC")
            self.assertEqual(daten.cell(row=row, column=2).value, 4.4167)
            messages = [cell.value for cell in workbook["Notes"]["C"]]
            hill = messages.index(
                "Daten: SchanzenID Einsiedeln HS117 in Liste_Schanzen_ID gefunden, "
                "Pts / Gate DGW 4.4167."
            )
            jumpers = messages.index("Ranglisten: Alle JumperID in SJ-Listen.xlsm vorhanden.")
            self.assertLess(hill, jumpers)
            workbook.close()

    def test_missing_jumper_and_hill_are_red_warnings(self):
        with tempfile.TemporaryDirectory() as folder:
            listen_file = write_listen(
                Path(folder) / "SJ-Listen.xlsx",
                jumpers=[("SUI HOESLI LION", "HOESLI Lion", "SUI", "M")],
                hills=[("Engelberg HS140", 5.1)],
            )
            warnings, workbook = self.convert(folder, listen_file)
            self.assertEqual(
                warnings,
                [
                    "Daten: SchanzenID Einsiedeln HS117 fehlt in Liste_Schanzen_ID (SJ-Listen.xlsm).",
                    "Ranglisten: JumperID 'SUI MUSTER HANS' fehlt in Liste_Springer_ID: "
                    "MUSTER Hans Peter (SUI).",
                    "Ranglisten: JumperID 'SUI MUSTER FRITZ' fehlt in Liste_Springer_ID: "
                    "MUSTER Fritz (SUI).",
                ],
            )
            ranglisten = workbook["Ranglisten"]
            self.assertIsNone(ranglisten["AE3"].value)
            self.assertEqual(ranglisten["AE3"].fill.fgColor.rgb, WARNING_FILL.fgColor.rgb)
            self.assertNotEqual(ranglisten["AE2"].fill.fgColor.rgb, WARNING_FILL.fgColor.rgb)
            notes = workbook["Notes"]
            warning_rows = [
                row for row in notes.iter_rows() if row[0].value == "Warnung"
            ]
            self.assertEqual(len(warning_rows), 3)
            self.assertEqual(
                [row[1].value for row in warning_rows], ["Daten", "Ranglisten", "Ranglisten"]
            )
            self.assertTrue(all(row[2].fill.fgColor.rgb == WARNING_FILL.fgColor.rgb for row in warning_rows))
            workbook.close()

    def test_jump_times_are_interpolated_between_start_and_end(self):
        with tempfile.TemporaryDirectory() as folder:
            _warnings, workbook = self.convert(folder, None)
            ranglisten = workbook["Ranglisten"]
            header = [cell.value for cell in ranglisten[1]]
            date_column = header.index("Date") + 1
            time_column = header.index("Time") + 1
            dates = [ranglisten.cell(row=row, column=date_column) for row in range(2, 6)]
            times = [ranglisten.cell(row=row, column=time_column).value for row in range(2, 6)]
            self.assertEqual({cell.value.date() for cell in dates}, {date(2026, 9, 12)})
            self.assertEqual(dates[0].number_format, "yyyy-mm-dd")
            self.assertEqual(ranglisten.cell(row=2, column=time_column).number_format, "hh:mm:ss")
            # bib 1, 2 and 4 jumped (13:30 to 14:23), bib 3 did not start
            self.assertEqual(times, [time(13, 30), time(13, 56, 30), None, time(14, 23)])
            workbook.close()

    def test_daten_column_a_keeps_no_old_section_fill(self):
        with tempfile.TemporaryDirectory() as folder:
            output_file = Path(folder) / "latest event.xlsx"
            old = Workbook()
            old.active.title = "Daten"
            old["Daten"]["A19"].fill = WARNING_FILL
            old.save(output_file)
            _warnings, workbook = self.convert(folder, None)
            daten = workbook["Daten"]
            self.assertEqual(daten["A19"].value, "Pts / Gate DGW")
            self.assertIsNone(daten["A19"].fill.fill_type)
            self.assertEqual(daten["A20"].value, "DG 1")
            self.assertEqual(daten["A20"].fill.fgColor.rgb, "FFD9D9D9")
            workbook.close()

    def test_missing_list_file_is_a_warning_and_conversion_continues(self):
        with tempfile.TemporaryDirectory() as folder:
            warnings, workbook = self.convert(folder, Path(folder) / "SJ-Listen.xlsm")
            self.assertEqual(len(warnings), 1)
            self.assertTrue(warnings[0].startswith("SJ-Listen: "))
            self.assertIn("Ranglisten", workbook.sheetnames)
            workbook.close()


if __name__ == "__main__":
    unittest.main()
