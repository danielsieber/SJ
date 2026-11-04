from email.message import Message
from io import BytesIO
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from urllib.error import HTTPError

from openpyxl import load_workbook

from fis_raceid_crawler import (
    DEFAULT_OUTPUT_PATH,
    HEADERS,
    _calendar_events,
    _competition_rows,
    _download,
    write_workbook,
)


class FisCrawlerTests(unittest.TestCase):
    def test_default_output_is_in_sjpdf_project_data(self):
        self.assertEqual(
            DEFAULT_OUTPUT_PATH,
            Path(__file__).resolve().parents[2] / "data" / "raceid" / "raceid.xlsx",
        )

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
                "category",
                "gender",
            ),
        )

    def test_workbook_sorts_rows_and_puts_mixed_events_in_mixed_sheet(self):
        records = [
            {
                "season": 2027,
                "date": date(2026, 10, 25),
                "eventid": "62788",
                "raceid": "7757",
                "codex": "0279",
                "place": "Klingenthal",
                "category": "GP",
                "gender": "A",
            },
            {
                "season": 2027,
                "date": date(2026, 10, 24),
                "eventid": "62788",
                "raceid": "7756",
                "codex": "0278",
                "place": "Klingenthal",
                "category": "GP",
                "gender": "M",
            },
            {
                "season": 2027,
                "date": date(2026, 10, 24),
                "eventid": "62788",
                "raceid": "7755",
                "codex": "0277",
                "place": "Klingenthal",
                "category": "GP",
                "gender": "W",
            },
            {
                "season": 2027,
                "date": date(2026, 10, 24),
                "eventid": "62788",
                "raceid": "7753",
                "codex": "0275",
                "place": "Klingenthal",
                "category": "QUA",
                "gender": "W",
            },
        ]
        with TemporaryDirectory() as directory:
            path = Path(directory) / "raceid.xlsx"
            write_workbook(records, path)
            workbook = load_workbook(path, read_only=True, data_only=True)
            sheetnames = workbook.sheetnames
            mixed_rows = list(workbook["mixed"].iter_rows(values_only=True))
            all_codexes = [
                row[4] for row in workbook["all"].iter_rows(min_row=2, values_only=True)
            ]
            women_codexes = [
                row[4]
                for row in workbook["women"].iter_rows(min_row=2, values_only=True)
            ]
            workbook.close()

        self.assertEqual(sheetnames, ["men", "women", "mixed", "all"])
        self.assertEqual(len(mixed_rows), 2)
        self.assertEqual(mixed_rows[1][4:], ("0279", "Klingenthal", "GP", "A"))
        self.assertEqual(all_codexes, ["0275", "0277", "0278", "0279"])
        self.assertEqual(women_codexes, ["0275", "0277"])

    def test_calendar_without_event_list_reports_changed_page(self):
        with self.assertRaisesRegex(ValueError, "calendardata"):
            _calendar_events("<html><body>Unexpected page</body></html>")

    @patch("fis_raceid_crawler.time.sleep")
    @patch("fis_raceid_crawler.urlopen")
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


if __name__ == "__main__":
    unittest.main()
