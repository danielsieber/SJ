import unittest
from datetime import date
from io import BytesIO

from openpyxl import Workbook, load_workbook

from src.pipeline.phase_20_preprocessing.fol_calc_excel import (
    COLOR_CALCULATED,
    process_fol_calc_worksheet,
)


HEADERS = [
    "Name",
    "DG",
    "Typ",
    "Vorname",
    "NSA",
    "Dist_oben",
    "Dist_m",
    "Dpts",
    "Gate",
    "Gpts",
    "Wind",
    "Wpts",
    "A",
    "B",
    "C",
    "D",
    "E",
    "Jpts",
    "Rpts",
    "Total",
    "km/h",
    "To_beat",
    "BIB",
    "Codex",
    "Date",
    "MSG",
    "Jumper_ID",
    "Jump_ID",
]


def row(**values):
    return [values.get(header) for header in HEADERS]


class FolCalcExcelTests(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "fol_calc"
        self.worksheet.append(HEADERS)
        self.worksheet.append(
            row(
                Name="Alpha",
                DG=1,
                Typ=" q ",
                Vorname="Ada",
                NSA=" sui ",
                Dist_oben=100,
                Gate=9,
                Wind=-1,
                Wpts=-5,
                A=17,
                B=16,
                C=16.5,
                D=18,
                E=16,
                Rpts=0,
                Total=70,
                **{"km/h": 91.25},
            )
        )
        self.worksheet.append(
            row(
                Name="Beta",
                Dist_m=95,
                Dpts=60,
                Gate=9,
                Wind=1,
                Wpts=5,
                Total=60,
            )
        )
        self.worksheet.append(
            row(
                Name="Gamma",
                Dist_m=100,
                Gate=10,
                Gpts=3,
                Wind=0,
                Total=70,
            )
        )
        self.worksheet.append(
            row(
                Name="Delta",
                DG=2,
                Typ="T",
                Dist_m=95,
                Gate=9,
                Gpts=4,
                Wind=2,
                Total=64,
            )
        )

        schanzen = self.workbook.create_sheet("SchanzenID")
        schanzen.append(["SchanzenID", "Kandersteg HS106"])
        schanzen.append(["K", 95])
        schanzen.append(["Pts./m", 2])
        self.schanzen = schanzen

    def test_calculations_are_header_driven_and_preserve_existing_values(self):
        process_fol_calc_worksheet(self.worksheet, self.schanzen)

        columns = {cell.value: cell.column for cell in self.worksheet[1]}
        cell = lambda row_number, header: self.worksheet.cell(
            row_number, columns[header]
        )

        self.assertEqual(cell(3, "DG").value, "=B2")
        self.assertEqual(cell(3, "Typ").value, "=C2")
        self.assertEqual(cell(2, "Gpts").value, 0)
        self.assertEqual(cell(3, "Gpts").value, 0)
        self.assertEqual(cell(4, "Gpts").value, 3)
        self.assertEqual(cell(5, "Gpts").value, 4)

        self.assertEqual(cell(2, "Dist_m").value, "=F2")
        self.assertEqual(
            cell(2, "Dpts").value,
            '=IF(G2="","",ROUND(IF(\'SchanzenID\'!$B$2<=160,60,120)'
            "+(G2-'SchanzenID'!$B$2)*'SchanzenID'!$B$3,1))",
        )
        self.assertEqual(cell(3, "Dpts").value, 60)
        self.assertEqual(cell(4, "Dpts").value, 70)
        self.assertEqual(cell(5, "Dpts").value, 60)

        self.assertEqual(
            cell(2, "Jpts").value,
            "=SUM(M2:Q2)-MIN(M2:Q2)-MAX(M2:Q2)",
        )
        self.assertEqual(cell(4, "Wpts").value, 0)
        self.assertEqual(cell(2, "Wpts").value, -5)
        self.assertEqual(cell(2, "Rpts").value, "=H2+J2+L2+R2")

        calculated_cells = [
            cell(2, "Gpts"),
            cell(2, "Dist_m"),
            cell(2, "Dpts"),
            cell(2, "Jpts"),
            cell(4, "Wpts"),
            cell(2, "Rpts"),
        ]
        self.assertTrue(
            all(
                calculated.fill.fgColor.rgb[-6:] == COLOR_CALCULATED[-6:]
                for calculated in calculated_cells
            )
        )
        self.assertNotEqual(
            cell(3, "DG").fill.fgColor.rgb[-6:], COLOR_CALCULATED[-6:]
        )
        self.assertNotEqual(
            cell(3, "Typ").fill.fgColor.rgb[-6:], COLOR_CALCULATED[-6:]
        )

    def test_missing_total_falls_back_to_current_rpts(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(HEADERS)
        worksheet.append(row(DG=1, Typ="Q", Rpts=10, Total=10))
        worksheet.append(row(Rpts=20))
        worksheet.append(row(Rpts=30, Total=30))
        worksheet.append(row(DG=2, Typ="W", Rpts=10, Total=10))
        worksheet.append(row(Rpts=20, Total=21))
        worksheet.append(row(Rpts=30))

        process_fol_calc_worksheet(worksheet)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        total_column = columns["Total"]
        rpts_letter = worksheet.cell(1, columns["Rpts"]).column_letter

        self.assertEqual(
            worksheet.cell(3, total_column).value,
            f"={rpts_letter}3",
        )
        self.assertEqual(
            worksheet.cell(3, total_column).fill.fgColor.rgb[-6:],
            COLOR_CALCULATED[-6:],
        )
        self.assertEqual(worksheet.cell(7, total_column).value, f"={rpts_letter}7")

    def test_blank_distance_gets_a_deferred_distance_points_formula(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(HEADERS)
        worksheet.append(row(DG=1, Typ="Q", Name="Missing distance"))

        schanzen = workbook.create_sheet("SchanzenID")
        schanzen.append(["SchanzenID", "Wisla HS134"])
        schanzen.append(["K", 120])
        schanzen.append(["Pts./m", 1.8])

        process_fol_calc_worksheet(worksheet, schanzen)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        distance_cell = worksheet.cell(2, columns["Dist_m"])
        dpts_cell = worksheet.cell(2, columns["Dpts"])

        self.assertIsNone(distance_cell.value)
        self.assertEqual(
            dpts_cell.value,
            '=IF(G2="","",ROUND(IF(\'SchanzenID\'!$B$2<=160,60,120)'
            "+(G2-'SchanzenID'!$B$2)*'SchanzenID'!$B$3,1))",
        )
        self.assertEqual(
            dpts_cell.fill.fgColor.rgb[-6:],
            COLOR_CALCULATED[-6:],
        )

    def test_second_w_round_sums_current_and_previous_rpts_for_same_bib(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(HEADERS)
        worksheet.append(row(DG=1, Typ="W", BIB=11, Rpts=40, Total=40))
        worksheet.append(row(BIB=12, Rpts=50, Total=50))
        worksheet.append(row(BIB=13, Rpts=60, Total=60))
        worksheet.append(row(DG=2, Typ=" w ", BIB=11, Rpts=10, Total=50))
        worksheet.append(row(BIB=12, Rpts=20, Total=70))
        worksheet.append(row(BIB=13, Rpts=30, Total=30))
        worksheet.append(row(BIB=12, Rpts=25))
        worksheet.append(row(BIB=99, Rpts=15))

        process_fol_calc_worksheet(worksheet)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        total_column = columns["Total"]
        rpts_letter = worksheet.cell(1, columns["Rpts"]).column_letter

        self.assertEqual(
            worksheet.cell(8, total_column).value,
            f"={rpts_letter}8+{rpts_letter}3",
        )
        self.assertEqual(
            worksheet.cell(9, total_column).value,
            f"={rpts_letter}9",
        )
        self.assertEqual(worksheet.cell(5, total_column).value, 50)
        self.assertEqual(
            worksheet.cell(8, total_column).fill.fgColor.rgb[-6:],
            COLOR_CALCULATED[-6:],
        )

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        reloaded = load_workbook(output, data_only=False)
        reloaded_total = reloaded["fol_calc"].cell(8, total_column)
        self.assertEqual(
            reloaded_total.value,
            f"={rpts_letter}8+{rpts_letter}3",
        )
        self.assertEqual(
            reloaded_total.fill.fgColor.rgb[-6:],
            COLOR_CALCULATED[-6:],
        )

    def test_second_w_round_requires_a_strict_majority_of_differences(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(HEADERS)
        worksheet.append(row(DG=1, Typ="W", BIB=1, Rpts=40, Total=40))
        worksheet.append(row(BIB=2, Rpts=50, Total=50))
        worksheet.append(row(DG=2, Typ="W", BIB=1, Rpts=10, Total=50))
        worksheet.append(row(BIB=2, Rpts=20, Total=20))
        worksheet.append(row(BIB=3, Rpts=30, Total=None))
        worksheet.append(row(BIB=1, Rpts=15))

        process_fol_calc_worksheet(worksheet)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        total_column = columns["Total"]
        rpts_letter = worksheet.cell(1, columns["Rpts"]).column_letter

        # One differing and one equal pair is a tie, while blank Total values
        # are excluded from the decision.
        self.assertEqual(
            worksheet.cell(7, total_column).value,
            f"={rpts_letter}7",
        )

    def test_adds_codex_hyperlinks_with_season_and_type_mapping(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(HEADERS)
        worksheet.append(
            row(DG=1, Typ="W", Codex=3204, Date=date(2026, 3, 15))
        )
        worksheet.append(row(Codex="3205", Date="15.07.2025"))
        worksheet.append(
            row(DG=2, Typ="T", Codex="3206.0", Date="30.06.2025")
        )
        worksheet.append(
            row(DG=3, Typ=" q ", Codex=3207, Date="2026-01-10")
        )
        worksheet.append(
            row(DG=4, Typ="P", Codex=3208, Date="2026/02/10")
        )
        worksheet.append(row(DG=5, Typ="X", Codex=3209, Date="2026-03-10"))

        process_fol_calc_worksheet(worksheet)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        codex_column = columns["Codex"]
        typ_column = columns["Typ"]
        expected_links = {
            2: "https://medias1.fis-ski.com/pdf/2026/JP/3204/2026JP3204RL.pdf",
            3: "https://medias1.fis-ski.com/pdf/2026/JP/3205/2026JP3205RL.pdf",
            4: "https://medias1.fis-ski.com/pdf/2026/JP/3206/2026JP3206TNG.pdf",
            5: "https://medias1.fis-ski.com/pdf/2026/JP/3207/2026JP3207QUAL.pdf",
            6: "https://medias1.fis-ski.com/pdf/2026/JP/3208/2026JP3208TRIAL.pdf",
        }
        for row_number, expected_link in expected_links.items():
            cell = worksheet.cell(row_number, typ_column)
            self.assertEqual(cell.hyperlink.target, expected_link)
            self.assertEqual(cell.font.color.rgb, "FF0563C1")
            self.assertEqual(cell.font.underline, "single")

        self.assertEqual(worksheet.cell(2, codex_column).value, 3204)
        self.assertIsNone(worksheet.cell(2, codex_column).hyperlink)
        self.assertIsNone(worksheet.cell(7, typ_column).hyperlink)

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        reloaded = load_workbook(output, data_only=False)
        self.assertEqual(
            reloaded["fol_calc"].cell(2, typ_column).hyperlink.target,
            expected_links[2],
        )

    def test_standard_formats_are_applied_to_data_and_headers(self):
        process_fol_calc_worksheet(self.worksheet, self.schanzen)

        columns = {cell.value: cell.column for cell in self.worksheet[1]}
        cell = lambda row_number, header: self.worksheet.cell(
            row_number, columns[header]
        )

        self.assertEqual(cell(2, "Name").alignment.horizontal, "left")
        self.assertEqual(cell(2, "Vorname").alignment.horizontal, "left")
        self.assertEqual(cell(2, "MSG").alignment.horizontal, "left")
        self.assertEqual(cell(2, "Jumper_ID").alignment.horizontal, "left")
        self.assertEqual(cell(2, "Jump_ID").alignment.horizontal, "left")
        self.assertEqual(cell(2, "NSA").alignment.horizontal, "center")
        self.assertEqual(cell(1, "Name").alignment.horizontal, "center")
        self.assertEqual(cell(1, "Vorname").alignment.horizontal, "center")
        self.assertEqual(cell(1, "NSA").alignment.horizontal, "center")
        self.assertTrue(cell(1, "Name").font.bold)
        self.assertEqual(
            cell(1, "Name").fill.fgColor.rgb[-6:],
            COLOR_CALCULATED[-6:],
        )
        self.assertEqual(cell(2, "Dist_m").number_format, "0.0")
        self.assertEqual(cell(2, "Wind").number_format, "0.00")
        self.assertEqual(cell(1, "Wind").number_format, "General")
        for header in ("BIB", "Dpts", "Gpts", "Wpts", "Jpts", "Rpts"):
            self.assertTrue(cell(2, header).font.bold)

        rules_by_range = {
            str(conditional_format.sqref): rules
            for conditional_format, rules
            in self.worksheet.conditional_formatting._cf_rules.items()
        }
        self.assertIn("B2:B5", rules_by_range)
        self.assertIn("C2:C5", rules_by_range)
        self.assertIn("A2:A5", rules_by_range)
        self.assertIn("D2:D5", rules_by_range)
        self.assertIn("E2:E5", rules_by_range)

        last_column_letter = self.worksheet.cell(
            1, self.worksheet.max_column
        ).column_letter
        formula_range = f"G2:{last_column_letter}5"
        formula_rules = [
            rule
            for rule in rules_by_range[formula_range]
            if any(
                "_xlfn.ISFORMULA(G2)" in formula
                for formula in rule.formula
            )
        ]
        self.assertEqual(len(formula_rules), 1)
        self.assertEqual(
            formula_rules[0].dxf.font.color.rgb,
            "FF00B050",
        )

        dg_rules = rules_by_range["B2:B5"]
        self.assertEqual(
            {
                rule.dxf.fill.fgColor.rgb
                for rule in dg_rules
                if rule.dxf.fill is not None
            },
            {"FFD9D9D9", "FFBFBFBF"},
        )
        self.assertTrue(
            any("MOD($B2,2)=0" in rule.formula[0] for rule in dg_rules)
        )
        self.assertTrue(
            any("MOD($B2,2)=1" in rule.formula[0] for rule in dg_rules)
        )

        for target_range in ("A2:A5", "D2:D5", "E2:E5"):
            sui_rules = [
                rule
                for rule in rules_by_range[target_range]
                if 'UPPER(TRIM($E2&""))="SUI"' in rule.formula
            ]
            self.assertEqual(len(sui_rules), 1)
            self.assertEqual(sui_rules[0].dxf.font.color.rgb, "FF0000FF")

        for header in ("A", "B", "C", "D", "E", "Jpts"):
            column = columns[header]
            letter = self.worksheet.cell(1, column).column_letter
            target_range = f"{letter}2:{letter}5"
            judge_rules = [
                rule
                for rule in rules_by_range[target_range]
                if rule.dxf is not None
                and rule.dxf.fill is not None
                and rule.dxf.fill.fgColor.rgb == "FFF4CCCC"
            ]
            self.assertEqual(len(judge_rules), 2)
            self.assertTrue(
                any(
                    f'UPPER(TRIM($C2&""))="T"' in rule.formula[0]
                    and f"{letter}2>0" in rule.formula[0]
                    for rule in judge_rules
                )
            )
            self.assertTrue(
                any(
                    f'UPPER(TRIM($C2&""))="Q"' in rule.formula[0]
                    and f'{letter}2=""' in rule.formula[0]
                    and f"{letter}2<3" in rule.formula[0]
                    for rule in judge_rules
                )
            )

        for header in (
            "BIB",
            "Name",
            "Vorname",
            "NSA",
            "km/h",
            "Dist_m",
            "Dpts",
            "Gate",
            "Gpts",
            "Wind",
            "Wpts",
            "Rpts",
            "Total",
            "Dist_oben",
        ):
            column = columns[header]
            letter = self.worksheet.cell(1, column).column_letter
            target_range = f"{letter}2:{letter}5"
            missing_rules = [
                rule
                for rule in rules_by_range[target_range]
                if rule.formula == [f'{letter}2=""']
            ]
            self.assertEqual(len(missing_rules), 1)
            self.assertEqual(
                missing_rules[0].dxf.fill.fgColor.rgb,
                "FFF4CCCC",
            )

        distance_column = columns["Dist_m"]
        distance_letter = self.worksheet.cell(1, distance_column).column_letter
        distance_range = f"{distance_letter}2:{distance_letter}5"
        invalid_distance_rules = [
            rule
            for rule in rules_by_range[distance_range]
            if rule.formula == [
                f'AND({distance_letter}2<>"",'
                f"OR(NOT(ISNUMBER({distance_letter}2)),"
                f"MOD({distance_letter}2*2,1)<>0))"
            ]
        ]
        self.assertEqual(len(invalid_distance_rules), 1)
        self.assertEqual(
            invalid_distance_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )

        springer_id_column = columns["Jumper_ID"]
        springer_id_letter = self.worksheet.cell(
            1, springer_id_column
        ).column_letter
        springer_id_range = f"{springer_id_letter}2:{springer_id_letter}5"
        springer_id_rules = [
            rule
            for rule in rules_by_range[springer_id_range]
            if "COUNTIF" in rule.formula[0]
        ]
        self.assertEqual(len(springer_id_rules), 1)
        self.assertEqual(
            springer_id_rules[0].formula,
            [
                f'OR({springer_id_letter}2="",'
                f'COUNTIF(INDIRECT("\'SpringerID\'!A:A"),'
                f'{springer_id_letter}2)=0)'
            ],
        )
        self.assertEqual(
            springer_id_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )

        # Verify that formulas and standard styles survive xlsx I/O.
        output = BytesIO()
        self.workbook.save(output)
        output.seek(0)
        reloaded = load_workbook(output, data_only=False)
        reloaded_sheet = reloaded["fol_calc"]
        self.assertEqual(
            reloaded_sheet["H2"].value,
            '=IF(G2="","",ROUND(IF(\'SchanzenID\'!$B$2<=160,60,120)'
            "+(G2-'SchanzenID'!$B$2)*'SchanzenID'!$B$3,1))",
        )
        self.assertGreater(len(reloaded_sheet.conditional_formatting), 0)

    def test_primary_missing_rule_spans_bib_through_distance_points(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "fol_calc"
        worksheet.append(
            [
                "DG", "Typ", "BIB", "Name", "Vorname", "NSA",
                "km/h", "Dist_m", "Dpts",
            ]
        )
        worksheet.append([1, "Q", 21, None, None, None, 91.8, 100.2, None])
        worksheet.append([1, "Q", 22, "Maurer", "Nicole", "CAN", 91.8, 95.5, 15.9])

        schanzen = workbook.create_sheet("SchanzenID")
        schanzen.append(["SchanzenID", "Wisla HS134"])
        schanzen.append(["K", 120])
        schanzen.append(["Pts./m", 1.8])

        process_fol_calc_worksheet(worksheet, schanzen)

        rules_by_range = {
            str(conditional_format.sqref): rules
            for conditional_format, rules
            in worksheet.conditional_formatting._cf_rules.items()
        }
        primary_rules = [
            rule
            for rule in rules_by_range["C2:I3"]
            if rule.formula == ['C2=""']
        ]
        self.assertEqual(len(primary_rules), 1)
        self.assertEqual(
            primary_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )

        invalid_distance_rules = [
            rule
            for rule in rules_by_range["H2:H3"]
            if rule.formula
            == ['AND(H2<>"",OR(NOT(ISNUMBER(H2)),MOD(H2*2,1)<>0))']
        ]
        self.assertEqual(len(invalid_distance_rules), 1)
        self.assertEqual(
            invalid_distance_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )

    def test_duplicate_bib_rule_is_scoped_to_the_same_dg_and_survives_xlsx_io(self):
        worksheet = self.workbook.create_sheet("duplicate_bibs")
        worksheet.append(HEADERS)
        worksheet.append(row(DG=1, Typ="Q", BIB=32))
        worksheet.append(row(BIB=32))
        worksheet.append(row(DG=2, Typ="Q", BIB=32))
        worksheet.append(row(BIB=None))

        process_fol_calc_worksheet(worksheet)

        columns = {cell.value: cell.column for cell in worksheet[1]}
        dg_letter = worksheet.cell(1, columns["DG"]).column_letter
        bib_letter = worksheet.cell(1, columns["BIB"]).column_letter
        bib_range = f"{bib_letter}2:{bib_letter}5"
        rules_by_range = {
            str(conditional_format.sqref): rules
            for conditional_format, rules
            in worksheet.conditional_formatting._cf_rules.items()
        }
        duplicate_rules = [
            rule
            for rule in rules_by_range[bib_range]
            if "_xlfn.COUNTIFS" in rule.formula[0]
        ]

        self.assertEqual(len(duplicate_rules), 1)
        self.assertEqual(
            duplicate_rules[0].formula,
            [
                f'AND(${dg_letter}2<>"",${bib_letter}2<>"",'
                f'_xlfn.COUNTIFS(${dg_letter}$2:${dg_letter}$5,${dg_letter}2,'
                f'${bib_letter}$2:${bib_letter}$5,${bib_letter}2)>1)'
            ],
        )
        self.assertEqual(
            duplicate_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )

        output = BytesIO()
        self.workbook.save(output)
        output.seek(0)
        reloaded = load_workbook(output, data_only=False)
        reloaded_sheet = reloaded["duplicate_bibs"]
        reloaded_rules = [
            rule
            for conditional_format, rules
            in reloaded_sheet.conditional_formatting._cf_rules.items()
            if str(conditional_format.sqref) == bib_range
            for rule in rules
            if "_xlfn.COUNTIFS" in rule.formula[0]
        ]
        self.assertEqual(len(reloaded_rules), 1)
        self.assertEqual(reloaded_rules[0].formula, duplicate_rules[0].formula)
        self.assertEqual(
            reloaded_rules[0].dxf.fill.fgColor.rgb,
            "FFF4CCCC",
        )


if __name__ == "__main__":
    unittest.main()
