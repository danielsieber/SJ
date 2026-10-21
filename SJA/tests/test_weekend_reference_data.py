import tempfile
import unittest
from datetime import date, time
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook

from src.config import load_reference_lists
from src.pipeline.phase_20_preprocessing.format_save import (
    _assign_msg_jump_ids,
    _assign_wind_jump_id,
    _build_jump_blocks,
    format_and_save,
)
from src.views.editing_stage import (
    _build_dg_summary,
    _format_save_error,
    _normalize_types_for_save,
    _validate_dg_summary,
)
from src.views.new_weekend import _build_springer_id_actual, _find_column


def fol_clean_row(*, source_row, name, vorname, nsa):
    row = [None] * 28
    row[0] = "3204"
    row[1] = "fol_1"
    row[2] = source_row
    row[7] = name
    row[8] = vorname
    row[9] = nsa
    return row


class SaveErrorMessageTests(unittest.TestCase):
    def test_permission_error_mentions_open_excel_file_and_path(self):
        output_path = Path(r"C:\Data_Weekend\2026-08-02 GPM Wisla HS134.xlsx")

        message = _format_save_error(
            PermissionError(13, "Permission denied", str(output_path)),
            output_path,
        )

        self.assertIn("noch in Excel geöffnet", message)
        self.assertIn("Bitte schließen Sie die Datei", message)
        self.assertIn(str(output_path), message)
        self.assertNotIn("Permission denied", message)

    def test_unexpected_error_keeps_original_details(self):
        message = _format_save_error(
            RuntimeError("Workbook beschädigt"),
            Path("weekend.xlsx"),
        )

        self.assertEqual(
            message,
            "Fehler beim Speichern: Workbook beschädigt",
        )


class SpringerIdActualTests(unittest.TestCase):
    def test_accepts_jumper_id_as_reference_column_name(self):
        reference = pd.DataFrame(columns=["JumperID", "Gender", "DoB", "JG"])

        self.assertEqual(
            _find_column(
                reference,
                "SpringerID",
                "Springer_ID",
                "JumperID",
                "Jumper_ID",
            ),
            "JumperID",
        )

    def test_builds_uppercase_five_character_ids_and_skips_duplicates(self):
        data = [
            fol_clean_row(
                source_row=17,
                name="Deschwanden",
                vorname="Gregor",
                nsa="sui",
            ),
            fol_clean_row(
                source_row=18,
                name="Deschwanden",
                vorname="Gregor",
                nsa="SUI",
            ),
            fol_clean_row(
                source_row=19,
                name="Kraft",
                vorname="Stefan",
                nsa="AUT",
            ),
        ]

        actual = _build_springer_id_actual(data)

        self.assertEqual(
            actual["SpringerID"].tolist(),
            ["SUI DESCH GREGO", "AUT KRAFT STEFA"],
        )
        self.assertEqual(
            actual.columns.tolist(),
            ["SpringerID", "Name", "Vorname", "NSA"],
        )

    def test_invalid_source_error_contains_original_row_number(self):
        data = [
            fol_clean_row(
                source_row=42,
                name="Deschwanden",
                vorname=None,
                nsa="SUI",
            )
        ]

        with self.assertRaisesRegex(ValueError, r"row nr 42.*Vorname fehlt"):
            _build_springer_id_actual(data)


class ReferenceListLoadingTests(unittest.TestCase):
    def test_detects_jumper_id_header_on_a_non_fixed_row(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            listen_path = Path(temp_dir) / "Listen.xlsm"
            workbook = Workbook()
            schanzen_sheet = workbook.active
            schanzen_sheet.title = "Liste_Schanzen_ID"
            schanzen_sheet.append(["Beschreibung"])
            schanzen_sheet.append([])
            schanzen_sheet.append(["Schanzen-ID", "HS"])
            schanzen_sheet.append(["Oslo HS134", 134])

            springer_sheet = workbook.create_sheet("Liste_Springer_ID")
            springer_sheet.append(["Titel"])
            springer_sheet.append([])
            springer_sheet.append([])
            springer_sheet.append(["Jumper-ID", "Gender", "DoB", "JG"])
            springer_sheet.append(["SUI DESCH GREGO", "M", "1991-02-27", 1991])
            workbook.save(listen_path)

            schanzen, springer = load_reference_lists(listen_path)

            self.assertEqual(schanzen.columns[0], "Schanzen-ID")
            self.assertEqual(schanzen.iloc[0, 0], "Oslo HS134")
            self.assertEqual(springer.columns[0], "Jumper-ID")
            self.assertEqual(springer.iloc[0, 0], "SUI DESCH GREGO")


class DgSaveValidationTests(unittest.TestCase):
    def test_summary_collects_each_codex_once_per_dg_block(self):
        rows = []
        for dg, codex in [(1, "A100"), (None, "A100"), (None, "B200")]:
            row = [None] * 33
            row[0] = dg
            row[27] = codex
            rows.append(row)

        summary = _build_dg_summary(rows)

        self.assertEqual(summary[0][3], "A100, B200")

    def test_summary_calculates_block_duration_in_minutes(self):
        first = [None] * 33
        first[0] = 1
        first[28] = date(2026, 7, 31)
        first[29] = time(23, 55)
        last = [None] * 33
        last[28] = date(2026, 8, 1)
        last[29] = time(0, 7, 30)

        summary = _build_dg_summary([first, last])

        self.assertEqual(summary[0][-2], 12.5)
        self.assertIsNone(summary[0][-1])

    def test_summary_calculates_gap_to_next_block_in_minutes(self):
        first = [None] * 33
        first[0] = 1
        first[28] = date(2026, 7, 31)
        first[29] = time(10, 0)
        first_last = [None] * 33
        first_last[28] = date(2026, 7, 31)
        first_last[29] = time(10, 5)
        second = [None] * 33
        second[0] = 2
        second[28] = date(2026, 7, 31)
        second[29] = time(10, 12, 30)

        summary = _build_dg_summary([first, first_last, second])

        self.assertEqual(summary[0][-1], 7.5)
        self.assertIsNone(summary[1][-1])

    def test_accepts_sequential_integers_and_lowercase_types(self):
        summary = [
            [1, "w"],
            [2.0, " q "],
            ["3", "T"],
            ["🔵 4", "p"],
        ]

        self.assertEqual(_validate_dg_summary(summary), [])

    def test_combines_dg_and_type_errors(self):
        summary = [
            [1, "W"],
            [3, "r"],
            [2.5, None],
        ]

        errors = _validate_dg_summary(summary)

        self.assertTrue(any("Position 2" in error and "erwartet 2" in error for error in errors))
        self.assertTrue(any("DG 3" in error and "'r'" in error for error in errors))
        self.assertTrue(any("Position 3" in error and "keine Ganzzahl" in error for error in errors))
        self.assertTrue(any("Typ bei DG 2.5" in error and "leer" in error for error in errors))

    def test_normalizes_valid_types_to_uppercase_before_save(self):
        first = [None] * 33
        first[0] = 1
        first[1] = " w "
        follow_up = [None] * 33
        second = [None] * 33
        second[0] = 2
        second[1] = "q"
        data = [first, follow_up, second]

        _normalize_types_for_save(data)

        self.assertEqual(data[0][1], "W")
        self.assertIsNone(data[1][1])
        self.assertEqual(data[2][1], "Q")


class FolWindJumpIdTests(unittest.TestCase):
    @staticmethod
    def calc_row(dg, typ, hour, minute):
        row = [None] * 33
        row[0] = dg
        row[1] = typ
        row[28] = date(2026, 3, 15)
        row[29] = time(hour, minute)
        return row

    @staticmethod
    def repair_row(hour=None, minute=None):
        row = [None] * 28
        if hour is not None:
            row[3] = date(2026, 3, 15)
            row[4] = time(hour, minute)
        return row

    def setUp(self):
        self.blocks = _build_jump_blocks(
            [
                self.calc_row(1, "W", 10, 0),
                self.calc_row(None, None, 10, 5),
                self.calc_row(2, "Q", 10, 10),
                self.calc_row(None, None, 10, 15),
            ],
            "weekend",
        )

    def test_assigns_at_start_minus_one_minute(self):
        self.assertEqual(
            _assign_wind_jump_id(self.repair_row(9, 59), self.blocks),
            "weekend 01W",
        )

    def test_assigns_at_end_plus_one_minute(self):
        self.assertEqual(
            _assign_wind_jump_id(self.repair_row(10, 6), self.blocks),
            "weekend 01W",
        )

    def test_does_not_assign_outside_tolerance(self):
        self.assertIsNone(
            _assign_wind_jump_id(self.repair_row(9, 58), self.blocks)
        )

    def test_missing_timestamp_does_not_assign(self):
        self.assertIsNone(
            _assign_wind_jump_id(self.repair_row(), self.blocks)
        )


class MsgCleanJumpIdTests(unittest.TestCase):
    @staticmethod
    def calc_row(dg, typ, codex, hour, minute):
        row = [None] * 33
        row[0] = dg
        row[1] = typ
        row[27] = codex
        row[28] = date(2026, 3, 15)
        row[29] = time(hour, minute)
        return row

    @staticmethod
    def msg_row(codex, hour=None, minute=None):
        row = [None] * 10
        row[7] = codex
        if hour is not None:
            row[0] = date(2026, 3, 15)
            row[1] = time(hour, minute)
        return row

    def setUp(self):
        self.blocks = _build_jump_blocks(
            [
                self.calc_row(1, "W", 3204, 10, 0),
                self.calc_row(None, None, "3204.0", 10, 5),
                self.calc_row(2, "Q", 3204, 10, 15),
                self.calc_row(None, None, 3204, 10, 20),
                self.calc_row(3, "T", 9999, 10, 9),
                self.calc_row(None, None, 9999, 10, 11),
            ],
            "weekend",
        )

    def test_assigns_message_inside_same_codex_block(self):
        assignments = _assign_msg_jump_ids(
            [self.msg_row("3204", 10, 3)], self.blocks
        )

        self.assertEqual(assignments, ["weekend 01W"])

    def test_between_blocks_uses_smaller_time_difference(self):
        assignments = _assign_msg_jump_ids(
            [self.msg_row(3204, 10, 12)], self.blocks
        )

        self.assertEqual(assignments, ["weekend 02Q"])

    def test_codex_has_priority_over_closer_other_codex_block(self):
        assignments = _assign_msg_jump_ids(
            [self.msg_row(3204, 10, 11)], self.blocks
        )

        self.assertEqual(assignments, ["weekend 02Q"])

    def test_equal_distance_uses_earlier_block(self):
        assignments = _assign_msg_jump_ids(
            [self.msg_row(3204, 10, 10)], self.blocks
        )

        self.assertEqual(assignments, ["weekend 01W"])

    def test_missing_codex_block_or_timestamp_stays_empty(self):
        assignments = _assign_msg_jump_ids(
            [self.msg_row(1234, 10, 3), self.msg_row(3204)], self.blocks
        )

        self.assertEqual(assignments, [None, None])


class ReferenceSheetExportTests(unittest.TestCase):
    def test_writes_transposed_schanze_and_enriched_springer_sheets(self):
        schanze = pd.DataFrame(
            [{"SchanzenID": "Oberstdorf HS235", "HS": 235, "Country": "GER"}]
        )
        springer = pd.DataFrame(
            [
                {
                    "SpringerID": "SUI DESCH GREGO",
                    "Name": "Deschwanden",
                    "Vorname": "Gregor",
                    "NSA": "SUI",
                    "Gender": "M",
                    "DoB": "1991-02-27",
                    "JG": 1991,
                }
            ]
        )
        calc_row = [None] * 33
        calc_row[0] = 1
        calc_row[1] = "W"
        calc_row[3] = "  DESCHWANDEN "
        calc_row[4] = "gregor"
        calc_row[5] = " sui "
        calc_row[27] = 3204
        calc_row[28] = date(2025, 6, 1)
        calc_row[29] = time(12, 1)
        fallback_calc_row = [None] * 33
        fallback_calc_row[3] = "Kraft"
        fallback_calc_row[4] = "Stefan"
        fallback_calc_row[5] = "AUT"
        fallback_calc_row[28] = date(2025, 6, 1)
        fallback_calc_row[29] = time(12, 5)
        fallback_calc_row[30] = 732
        repair_with_wind = [None] * 28
        repair_with_wind[3] = date(2025, 6, 1)
        repair_with_wind[4] = time(12, 0)
        repair_with_wind[7] = " Deschwanden "
        repair_with_wind[8] = "Gregor"
        repair_with_wind[9] = "sui"
        repair_with_wind[15] = -1.25
        repair_without_wind = list(repair_with_wind)
        repair_without_wind[4] = time(12, 2)
        repair_without_wind[15] = None
        repair_outside_block = list(repair_with_wind)
        repair_outside_block[4] = time(12, 7)
        repair_outside_block[15] = 0.5
        msg_row = [
            date(2025, 6, 1),
            time(12, 3),
            "message",
            None,
            None,
            None,
            None,
            "3204.0",
            "msg_1",
            7,
        ]
        trialjumper_row = [
            date(2025, 6, 1),
            time(12, 1),
            "Trialjumper",
            None,
            "trialjumper",
            None,
            None,
            "3204.0",
            "msg_1",
            8,
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "weekend.xlsx"
            pdf_output_path = Path(temp_dir) / "Data_pdf"
            with patch(
                "src.pipeline.phase_20_preprocessing.format_save."
                "_open_saved_workbook"
            ) as open_workbook, patch(
                "src.pipeline.phase_20_preprocessing.format_save."
                "start_fis_pdf_downloads"
            ) as start_downloads:
                format_and_save(
                    [],
                    [],
                    [msg_row, trialjumper_row],
                    [
                        repair_with_wind,
                        repair_without_wind,
                        repair_outside_block,
                    ],
                    [],
                    output_path.name,
                    output_path,
                    [calc_row, fallback_calc_row],
                    schanze,
                    springer,
                    pdf_output_path=pdf_output_path,
                )
                open_workbook.assert_called_once_with(output_path)
                start_downloads.assert_called_once_with(
                    {(2026, "3204")}, pdf_output_path
                )

            workbook = load_workbook(output_path, data_only=True)
            self.assertIn("SchanzenID", workbook.sheetnames)
            self.assertIn("SpringerID", workbook.sheetnames)
            fol_calc_position = workbook.sheetnames.index("fol_calc")
            self.assertEqual(
                workbook.sheetnames[fol_calc_position + 1],
                "fol_wind",
            )
            self.assertEqual(workbook.active.title, "fol_calc")
            self.assertEqual(workbook.calculation.calcMode, "auto")
            self.assertTrue(workbook.calculation.fullCalcOnLoad)
            self.assertTrue(workbook.calculation.forceFullCalc)
            self.assertTrue(workbook.calculation.calcOnSave)
            self.assertFalse(workbook.calculation.calcCompleted)
            self.assertEqual(workbook.calculation.calcId, 0)

            fol_calc_sheet = workbook["fol_calc"]
            self.assertTrue(fol_calc_sheet.sheet_view.tabSelected)
            self.assertEqual(fol_calc_sheet["E1"].value, "Vorname")
            self.assertEqual(fol_calc_sheet["AA1"].value, "MSG")
            self.assertEqual(fol_calc_sheet["AA2"].value, "message")
            self.assertIsNone(fol_calc_sheet["AA3"].value)
            self.assertEqual(fol_calc_sheet["AB1"].value, "Remark")
            self.assertIsNone(fol_calc_sheet["AB2"].value)
            self.assertEqual(fol_calc_sheet["AB3"].value, "trialjumper")
            self.assertEqual(fol_calc_sheet["AI1"].value, "Jumper_ID")
            self.assertEqual(fol_calc_sheet["AI2"].value, "SUI DESCH GREGO")
            self.assertEqual(fol_calc_sheet["AI3"].value, "AUT KRAFT STEFA")
            self.assertEqual(fol_calc_sheet["AJ1"].value, "Jump_ID")
            self.assertEqual(fol_calc_sheet.freeze_panes, "E2")
            self.assertEqual(fol_calc_sheet.column_dimensions["A"].width, 7)
            self.assertEqual(fol_calc_sheet.column_dimensions["D"].width, 14)
            self.assertEqual(fol_calc_sheet.column_dimensions["E"].width, 14)
            self.assertEqual(fol_calc_sheet.column_dimensions["Y"].width, 9)
            self.assertEqual(fol_calc_sheet.column_dimensions["Z"].width, 12)
            self.assertEqual(fol_calc_sheet.column_dimensions["AA"].width, 30)
            self.assertEqual(fol_calc_sheet.column_dimensions["AB"].width, 30)
            self.assertEqual(fol_calc_sheet.column_dimensions["AI"].width, 17)
            self.assertEqual(fol_calc_sheet.column_dimensions["AJ"].width, 13)
            self.assertEqual(fol_calc_sheet["AI2"].alignment.horizontal, "left")
            self.assertEqual(fol_calc_sheet["AJ2"].alignment.horizontal, "left")
            self.assertEqual(
                fol_calc_sheet.sheet_view.selection[0].activeCell,
                "A2",
            )
            self.assertEqual(
                fol_calc_sheet.sheet_view.selection[0].sqref,
                "A2",
            )

            msg_clean_sheet = workbook["msg_clean"]
            self.assertEqual(msg_clean_sheet["K1"].value, "Jump_ID")
            self.assertEqual(msg_clean_sheet["K2"].value, "weekend 01W")

            formula_workbook = load_workbook(output_path, data_only=False)
            formula_sheet = formula_workbook["fol_calc"]
            self.assertEqual(
                formula_sheet["AJ2"].value,
                '="weekend "&TEXT(A2,"00")&UPPER(TRIM(B2))',
            )
            self.assertEqual(
                formula_sheet["AJ3"].value,
                '="weekend "&TEXT(A3,"00")&UPPER(TRIM(B3))',
            )

            fol_repair_sheet = workbook["fol_repair"]
            self.assertEqual(fol_repair_sheet["I1"].value, "Vorname")

            fol_wind_sheet = workbook["fol_wind"]
            self.assertEqual(
                [cell.value for cell in fol_wind_sheet[1]],
                ["Datum", "Zeit", "Wind", "Jumper_ID", "Jump_ID"],
            )
            self.assertEqual(fol_wind_sheet.max_row, 3)
            self.assertEqual(
                fol_wind_sheet["A2"].value.date(),
                date(2025, 6, 1),
            )
            self.assertEqual(fol_wind_sheet["B2"].value, time(12, 0))
            self.assertEqual(fol_wind_sheet["C2"].value, -1.25)
            self.assertEqual(fol_wind_sheet["D2"].value, "SUI DESCH GREGO")
            self.assertEqual(fol_wind_sheet["E2"].value, "weekend 01W")
            self.assertEqual(fol_wind_sheet["C3"].value, 0.5)
            self.assertIsNone(fol_wind_sheet["E3"].value)
            self.assertEqual(fol_wind_sheet.freeze_panes, "A2")

            schanzen_sheet = workbook["SchanzenID"]
            self.assertEqual(schanzen_sheet["A1"].value, "SchanzenID")
            self.assertEqual(schanzen_sheet["B1"].value, "Oberstdorf HS235")
            self.assertEqual(schanzen_sheet["A2"].value, "HS")
            self.assertEqual(schanzen_sheet["B2"].value, 235)

            springer_sheet = workbook["SpringerID"]
            self.assertEqual(springer_sheet["A1"].value, "SpringerID")
            self.assertEqual(springer_sheet["A2"].value, "SUI DESCH GREGO")
            self.assertEqual(springer_sheet["E1"].value, "Gender")
            self.assertEqual(springer_sheet["G2"].value, 1991)


if __name__ == "__main__":
    unittest.main()
