from email.message import Message
from io import BytesIO
from datetime import date
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, patch
import warnings
import zipfile

from urllib.error import HTTPError

from openpyxl import load_workbook

from fis_raceid_crawler import (
    DEFAULT_OUTPUT_PATH,
    HEADERS,
    SHEET_NAME,
    WorkbookOpenError,
    _calendar_events,
    _competition_rows,
    _download,
    event_numbers,
    second_category,
    sheet_rows,
    write_event_sheet,
)


def record(codex, day, gender="M", category="GP", place="Klingenthal", hill="LH"):
    return {
        "season": 2027,
        "date": day,
        "eventid": "62788",
        "raceid": str(7000 + int(codex)),
        "codex": codex,
        "place": place,
        "category": category,
        "gender": gender,
        "hill": hill,
    }


class FisCrawlerTests(unittest.TestCase):
    def test_default_output_is_in_shared_data_directory(self):
        from sj_paths import DATA_DIRECTORY

        self.assertEqual(DEFAULT_OUTPUT_PATH, DATA_DIRECTORY / "SJ-Listen.xlsm")
        self.assertEqual(SHEET_NAME, "Event_ID")

    def test_calendar_events_extract_event_id_and_place(self):
        html = """
        <div id="calendardata">
          <div class="table-row reset-padding" id="63279">
            <div class="clip bold font_md_large font_lg_large">Predazzo</div>
          </div>
        </div>
        """

        self.assertEqual(_calendar_events(html), [("63279", "Predazzo")])

    def test_competitions_extract_fis_race_id_and_preserve_codex(self):
        html = """
        <div class="table-row reset-padding">
          <div class="timezone-date" data-date="2026-10-16">16 Oct</div>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">D</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">16 Oct</a>
          <a href="https://live.fis-ski.com/jp-0381/results-pda.htm">
            0381 review
          </a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">
            Individual Normal Hill
          </a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">FESA</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">M</a>
        </div>
        """

        records = _competition_rows(html, "63279", "Predazzo", 2027)

        self.assertEqual(
            records,
            [
                {
                    "season": 2027,
                    "date": date(2026, 10, 16),
                    "eventid": "63279",
                    "raceid": "7934",
                    "codex": "0381",
                    "place": "Predazzo",
                    "category": "FESA",
                    "gender": "M",
                    "hill": "NH",
                }
            ],
        )
        self.assertEqual(
            HEADERS,
            (
                "season",
                "date",
                "eventid",
                "raceid",
                "codex",
                "place",
                "hill",
                "cat1",
                "cat2",
                "gender",
                "EventNr",
            ),
        )

    def test_sheet_rows_are_sorted_with_header_and_excel_dates(self):
        rows = sheet_rows(
            [
                record("0279", date(2026, 10, 25), "A"),
                record("0278", date(2026, 10, 24)),
                record("0275", date(2026, 10, 24), "W", "QUA"),
            ]
        )
        self.assertEqual(rows[0], list(HEADERS))
        self.assertEqual([row[4] for row in rows[1:]], ["0275", "0278", "0279"])
        # Excel serial number of 2026-10-24
        self.assertEqual(rows[1][1], (date(2026, 10, 24) - date(1899, 12, 30)).days)
        self.assertEqual(rows[1][1], 46319)
        self.assertEqual([row[-1] for row in rows[1:]], [1, 2, 3])

    def test_rows_are_sorted_by_event_number_date_and_codex(self):
        rows = sheet_rows(
            [
                record("0010", date(2026, 12, 12), category="WC", place="Titisee"),
                record("0011", date(2026, 12, 13), category="WC", place="Titisee"),
                record("0012", date(2026, 12, 12), "W", "WC", place="Titisee"),
                record("0321", date(2026, 12, 11), category="WC", place="Titisee"),
            ]
        )
        self.assertEqual(
            [(row[-1], row[4]) for row in rows[1:]],
            # women end on 12 Dec, men on 13 Dec
            [(1, "0012"), (2, "0321"), (2, "0010"), (2, "0011")],
        )

    def test_cat2_turns_qualifications_into_gp_or_wc(self):
        self.assertEqual(second_category("QUA", date(2026, 5, 1)), "GP")
        self.assertEqual(second_category("QUA", date(2026, 10, 31)), "GP")
        self.assertEqual(second_category("QUA", date(2026, 11, 1)), "WC")
        self.assertEqual(second_category("QUA", date(2027, 4, 30)), "WC")
        self.assertEqual(second_category("COC", date(2026, 8, 1)), "COC")
        rows = sheet_rows([record("0275", date(2026, 10, 24), "W", "QUA")])
        self.assertEqual(rows[1][7:9], ["QUA", "GP"])

    def test_championship_qualification_joins_the_championship(self):
        rows = sheet_rows(
            [
                record("0064", date(2027, 2, 26), "W", "WSC", place="Falun", hill="NH"),
                record("0065", date(2027, 2, 26), "W", "QUA", place="Falun", hill="NH"),
                record("0071", date(2027, 3, 5), "W", "QUA", place="Falun", hill="LH"),
                record("0072", date(2027, 3, 5), "W", "WSC", place="Falun", hill="LH"),
            ]
        )
        self.assertEqual(
            [(row[4], row[8], row[-1]) for row in rows[1:]],
            [("0064", "WSC", 1), ("0065", "WSC", 1), ("0071", "WSC", 2), ("0072", "WSC", 2)],
        )

    def test_event_numbers_group_competitions_of_one_event(self):
        records = [
            record("0001", date(2026, 11, 20), category="QUA", place="Ruka"),
            record("0002", date(2026, 11, 21), place="Ruka", category="WC"),
            record("0003", date(2026, 11, 22), place="Ruka", category="WC"),
            record("0004", date(2026, 11, 22), "W", "WC", place="Ruka"),
            # same place and category, but a different hill
            record("0005", date(2026, 11, 23), place="Ruka", category="WC", hill="NH"),
            # same place, category and gender, more than 5 days later
            record("0006", date(2026, 11, 26), place="Ruka", category="WC"),
            record("0007", date(2026, 11, 21), place="Lillehammer", category="WC"),
        ]
        numbers = dict(
            zip(
                [item["codex"] for item in records],
                event_numbers(records),
            )
        )
        # the qualification (cat2 WC) belongs to the World Cup event
        self.assertEqual(numbers["0001"], numbers["0002"])
        self.assertEqual(numbers["0002"], numbers["0003"])
        # numbered by the last day: Lillehammer (21 Nov) before Ruka (22 Nov)
        self.assertEqual(list(numbers.values()), [2, 2, 2, 3, 4, 5, 1])

    def test_no_records_leave_the_workbook_untouched(self):
        excel = MagicMock()
        with TemporaryDirectory() as directory:
            path = Path(directory) / "SJ-Listen.xlsm"
            path.touch()
            with self.assertRaises(ValueError):
                write_event_sheet([], path, excel_factory=lambda: excel)
        excel.Workbooks.Open.assert_not_called()

    def test_open_workbook_is_reported(self):
        excel = MagicMock()
        with TemporaryDirectory() as directory:
            path = Path(directory) / "SJ-Listen.xlsm"
            path.touch()
            (Path(directory) / "~$SJ-Listen.xlsm").touch()
            with self.assertRaises(WorkbookOpenError):
                write_event_sheet(
                    [record("0278", date(2026, 10, 24))], path, excel_factory=lambda: excel
                )
        excel.Workbooks.Open.assert_not_called()

    def test_read_only_workbook_is_not_saved_and_excel_quits(self):
        excel = MagicMock()
        workbook = excel.Workbooks.Open.return_value
        workbook.ReadOnly = True
        with TemporaryDirectory() as directory:
            path = Path(directory) / "SJ-Listen.xlsm"
            path.touch()
            with self.assertRaises(WorkbookOpenError):
                write_event_sheet(
                    [record("0278", date(2026, 10, 24))], path, excel_factory=lambda: excel
                )
        workbook.Save.assert_not_called()
        workbook.Close.assert_called_once_with(False)
        excel.Quit.assert_called_once()

    def test_excel_writes_only_event_sheet(self):
        source = DEFAULT_OUTPUT_PATH
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.skipTest("pywin32 not installed")
        if not source.is_file():
            self.skipTest("SJ-Listen.xlsm not in the data folder")
        from openpyxl import load_workbook

        records = [
            record("0279", date(2026, 10, 25), "A"),
            record("0278", date(2026, 10, 24)),
        ]
        with TemporaryDirectory() as directory:
            path = Path(directory) / source.name
            shutil.copy2(source, path)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                before = load_workbook(path, keep_vba=True)
            other_sheets = {
                sheet.title: [row for row in sheet.iter_rows(values_only=True)]
                for sheet in before.worksheets
                if sheet.title != SHEET_NAME
            }
            parts = set(zipfile.ZipFile(path).namelist())
            before.close()

            write_event_sheet(records, path)

            self.assertEqual(set(zipfile.ZipFile(path).namelist()), parts)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                after = load_workbook(path, keep_vba=True)
            sheet = after[SHEET_NAME]
            # empty rows that keep only a row format (e.g. height) are ignored
            rows = [row for row in sheet.iter_rows(values_only=True) if any(row)]
            self.assertEqual(rows[0], HEADERS)
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[1][1].date(), date(2026, 10, 24))
            self.assertEqual(rows[1][2:7], ("62788", "7278", "0278", "Klingenthal", "LH"))
            self.assertEqual(sheet.auto_filter.ref, "A1:K3")
            self.assertEqual([row[-1] for row in rows[1:]], [1, 2])
            for title, values in other_sheets.items():
                self.assertEqual(
                    [row for row in after[title].iter_rows(values_only=True)], values
                )
            after.close()

    def test_calendar_without_event_list_reports_changed_page(self):
        with self.assertRaisesRegex(ValueError, "calendardata"):
            _calendar_events("<html><body>Unexpected page</body></html>")

    @patch("fis_web.time.sleep")
    @patch("fis_web.urlopen")
    def test_download_retries_rate_limit_using_retry_after(self, urlopen, sleep):
        error = HTTPError(
            "https://example.test",
            429,
            "Too Many Requests",
            {"Retry-After": "3"},
            None,
        )
        response = BytesIO(b"<html>ok</html>")
        response.headers = Message()
        urlopen.side_effect = [error, response]

        self.assertEqual(_download("https://example.test"), "<html>ok</html>")
        sleep.assert_called_once_with(3)

    def test_ambiguous_date_skips_only_that_competition(self):
        html = """
        <div class="table-row">
          <div class="timezone-date" data-date="2026-10-16">16 Oct</div>
          <div class="timezone-date" data-date="2026-10-17">17 Oct</div>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">0381 review</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">FESA</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7934">M</a>
        </div>
        <div class="table-row">
          <div class="timezone-date" data-date="2026-10-18">18 Oct</div>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7935">0382 review</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7935">FESA</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7935">W</a>
        </div>
        """
        warnings = []

        records = _competition_rows(
            html, "63279", "Predazzo", 2027, warn=warnings.append
        )

        self.assertEqual([record["codex"] for record in records], ["0382"])
        self.assertEqual(len(warnings), 1)
        self.assertIn("0381", warnings[0])


if __name__ == "__main__":
    unittest.main()
