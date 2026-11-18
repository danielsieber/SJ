from datetime import date
from pathlib import Path
import sys
import tempfile
import unittest

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fis_season_check as season  # noqa: E402
from pdf_to_xlsx.pdf2xlsx_Listen import Jumper, SjListen  # noqa: E402


RESULTS_HTML = """
<a class="table-row" href="https://www.fis-ski.com/DB/general/athlete-biography.html?sectorcode=jp&competitorid=131300&type=result">
  <div class="g-row container"><div class="g-row justify-sb">
    <div class="g-lg-1 bold">1</div><div class="g-lg-1 gray">29</div>
    <div class="g-lg-2 gray">5649</div><div class="g-lg bold">ROCHAS Caroline</div>
    <div class="g-lg-1">1992</div>
    <div class="g-lg-2"><div class="country country_flag"><span>FRA</span></div></div>
    <div class="g-lg-8">91.5</div>
  </div></div>
</a>
"""

START_LIST = """
  BIB NAME                                                     NSA   SKI BRAND  SPEED
       CLUB                                                            DATE OF BIRTH
     1 ROCHAS Caroline                                 FRA
       UNION SPORTIVE AUTRANAISE                                              29 SEP 1992
     5 MACUGA Samantha                             USA          Fischer
                                                                              17 FEB 2001
"""


class Attachment:
    def __init__(self, label, url="u"):
        self.label, self.url = label, url


class ParserTests(unittest.TestCase):
    def test_results_page_gives_fis_code_name_year_and_nation(self):
        self.assertEqual(
            season.results_athletes(RESULTS_HTML),
            [{"fis_code": "5649", "name": "ROCHAS Caroline", "birth_year": 1992,
              "nsa": "FRA", "competitorid": "131300"}],
        )

    def test_start_list_gives_date_of_birth_and_club(self):
        people = season.pdf_people(START_LIST)
        self.assertEqual(people[("rochas caroline", "fra")], (date(1992, 9, 29), "UNION SPORTIVE AUTRANAISE"))
        self.assertEqual(people[("macuga samantha", "usa")], (date(2001, 2, 17), ""))

    def test_biography_gives_birthdate_and_gender(self):
        html = "<div>Birthdate</div><div>15-03-2005</div><div>Gender</div><div>Female</div>"
        self.assertEqual(season.biography(html), (date(2005, 3, 15), "F"))

    def test_start_list_is_preferred_and_final_round_is_skipped(self):
        attachments = [
            Attachment("Official Results (389 kb) download"),
            Attachment("Start List Final Round (355 kb) download"),
            Attachment("Start List 1st Round (362 kb) download"),
        ]
        self.assertEqual(season.choose_pdf(attachments).label, attachments[2].label)

    def test_team_rows_are_skipped_and_give_the_nation(self):
        html = """
        <a class="table-row" href="x?competitorid=217884"><div class="g-row justify-sb">
          <div>1</div><div>11</div><div>7985</div><div>JAPAN</div><div></div>
          <div><span>JPN</span></div></div></a>
        <a class="table-row" href="x?competitorid=174083"><div class="g-row justify-sb">
          <div></div><div>11-1</div><div>6681</div><div>MARUYAMA Nozomi</div><div>1998</div>
          <div></div></div></a>
        """
        self.assertEqual(
            season.results_athletes(html),
            [{"fis_code": "6681", "name": "MARUYAMA Nozomi", "birth_year": 1998,
              "nsa": "JPN", "competitorid": "174083"}],
        )

    def test_hill_values_of_local_formats(self):
        self.assertEqual(
            season._local_hill_values("Hill Size / (95%)   106m (100,5m)\n K-Point   95m"),
            {"Hill Size": 106, "HS 95%": 100.5, "K-Punkt": 95},
        )
        self.assertEqual(
            season._local_hill_values("HILL-SIZE / (95%)  117m (111m)  GATE FACTOR  6.32 points\n"
                                      "WIND FACTORS - HEAD / TAIL   7.29 / 10.94"),
            {"Hill Size": 117, "HS 95%": 111, "Gate Factor": 6.32, "HWC": 7.29, "TWC": 10.94},
        )

    def test_hill_of_a_local_pdf(self):
        hill = season.pdf_hill("Youth\n Hill Size:   66\n K-Point:   60\n", "Klingenthal")
        self.assertEqual((hill["SchanzenID"], hill["K-Punkt"]), ("Klingenthal HS66", 60))


class CompareTests(unittest.TestCase):
    def setUp(self):
        self.listen = SjListen(
            jumper_ids={("toth daniela vasilica", "rou"): "ROU TOTH DANIE",
                        ("rochas caroline", "fra"): "FRA ROCHA CAROL"},
            jumpers=[Jumper("ROU TOTH DANIE", "TOTH Daniela Vasilica", "ROU", 6563),
                     Jumper("FRA ROCHA CAROL", "ROCHAS Caroline", "FRA", 5649)],
            hills={"courchevel hs132": "Courchevel HS132"},
        )
        self.data = season.SeasonData()
        for code, name, nsa in (("5649", "ROCHAS Caroline", "FRA"),
                                ("6563", "HARALAMBIE Daniela", "ROU"),
                                ("9999", "MUSTER Anna", "SUI"),
                                ("9998", "MUSTER Annina", "SUI")):
            self.data.athletes[code] = season.Athlete(code, name, nsa, 2008, None, {"F"})
        self.data.athletes["9998"].genders = set()  # only mixed teams: gender unknown
        for hill in ("Courchevel HS132", "Klingenthal HS66"):
            self.data.hills[hill.casefold()] = season.Hill(hill, {"SchanzenID": hill, "Hill Size": 66})

    def test_missing_jumpers_name_changes_and_hills(self):
        result = season.compare(self.data, self.listen)
        self.assertEqual(sorted(a.name for a in result.missing), ["MUSTER Anna", "MUSTER Annina"])
        self.assertEqual([(a.name, l.jumper_id) for a, l in result.check],
                         [("HARALAMBIE Daniela", "ROU TOTH DANIE")])
        self.assertEqual([h.hill_id for h in result.missing_hills], ["Klingenthal HS66"])

    def test_only_birth_year_gives_first_of_january(self):
        athlete = season.Athlete("9999", "MUSTER Anna", "SUI", 2011, None, {"F"})
        season.complete_from_biography([athlete], fetcher=None, log=lambda message: None)
        self.assertEqual((athlete.dob, athlete.dob_source), (date(2011, 1, 1), season.YEAR_ONLY))

    def test_workbook_rows_can_be_copied_into_the_list(self):
        self.data.athletes["9999"].dob = date(2008, 5, 1)
        result = season.compare(self.data, self.listen)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Fehlende_JumperID.xlsx"
            season.write_workbook(result, self.listen, 2026, self.data, path)
            workbook = load_workbook(path)
            sheet = workbook["Fehlende_JumperID"]
            rows = list(sheet.iter_rows(values_only=True))
            self.assertEqual(rows[0][:7], season.JUMPER_HEADERS)
            self.assertEqual([row[0] for row in rows[1:]], ["SUI MUSTE ANNA", "SUI MUSTE ANNIN"])
            anna = rows[1]
            self.assertEqual(anna[1:7], ("MUSTER Anna", "SUI", "F", anna[4], 9999, 2008))
            self.assertEqual(anna[4].date(), date(2008, 5, 1))
            self.assertIsNone(rows[2][3])  # gender unknown
            self.assertIn("Gender unbekannt", rows[2][11])
            self.assertEqual(workbook.sheetnames, ["Übersicht", "Fehlende_JumperID", "Prüfen", "Fehlende_SchanzenID"])
            self.assertEqual(workbook["Fehlende_SchanzenID"]["A2"].value, "Klingenthal HS66")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
