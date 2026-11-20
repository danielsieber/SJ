from datetime import date
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fis_pdf_archive as archive  # noqa: E402


class Attachment:
    def __init__(self, url, label="Official Results (300 kb) download"):
        self.url, self.label = url, label


def record(codex, day, category="COC", gender="M", place="Zakopane", hill="LH"):
    return {"season": 2027, "date": day, "eventid": "1", "raceid": codex, "codex": codex,
            "place": place, "category": category, "gender": gender, "hill": hill}


class FakeDownloader:
    def __init__(self):
        self.urls = []

    def download(self, url):
        self.urls.append(url)
        return f"2027JP{url[-4:]}RL.pdf", b"%PDF-1.4 test"


class FolderNameTests(unittest.TestCase):
    def test_like_list_event_without_number_and_codex_with_last_day(self):
        event = archive.Event(2027, 133, [record("0471", date(2027, 3, 13)),
                                          record("0472", date(2027, 3, 14))], "COC")
        self.assertEqual(archive.folder_name(event), "2027-03-14 COCM Zakopane LH")

    def test_unsafe_characters_and_missing_hill(self):
        event = archive.Event(2027, 5, [record("0286", date(2026, 8, 6), "FESA", "W",
                                               "Pöhla/Erzgebirge", None)], "AC")
        self.assertEqual(archive.folder_name(event), "2026-08-06 ACW Pöhla_Erzgebirge")


class SeasonEventTests(unittest.TestCase):
    def test_qualification_joins_its_world_cup_event(self):
        records = [
            record("0359", date(2027, 3, 18), "QUA", place="Planica", hill="FH"),
            record("0360", date(2027, 3, 19), "WC", place="Planica", hill="FH"),
            record("0059", date(2027, 3, 20), "WC", "W", "Planica", "FH"),
        ]

        class Fetcher:
            def text(self, url):
                return url

        originals = (archive.calendar_events, archive._competition_rows, archive._downloader)
        try:
            archive.calendar_events = lambda html: [("1", "Planica")]
            archive._competition_rows = lambda html, eventid, place, season, warn: records

            class Downloader:
                @staticmethod
                def competition_rows(html):
                    class Competition:
                        def __init__(self, codex):
                            self.codex = codex
                            self.attachments = (Attachment(f"u{codex}"),)
                    return [Competition(r["codex"]) for r in records]
            archive._downloader = Downloader
            events = archive.season_events(2027, Fetcher(), log=lambda message: None)
        finally:
            archive.calendar_events, archive._competition_rows, archive._downloader = originals
        names = {archive.folder_name(event): sorted(event.attachments) for event in events}
        self.assertEqual(names, {
            "2027-03-20 WCW Planica FH": ["u0059"],
            "2027-03-19 WCM Planica FH": ["u0359", "u0360"],
        })


class ArchiveEventTests(unittest.TestCase):
    def test_only_missing_pdfs_are_downloaded(self):
        event = archive.Event(2027, 1, [record("0471", date(2027, 3, 14))], "COC",
                              {"https://x/0471": Attachment("https://x/0471"),
                               "https://x/0472": Attachment("https://x/0472")})
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            downloader = FakeDownloader()
            folder, added, problems = archive.archive_event(event, output, downloader, log=lambda m: None)
            self.assertEqual((added, problems), (2, []))
            self.assertEqual(sorted(path.name for path in folder.iterdir()),
                             ["2027JP0471RL.pdf", "2027JP0472RL.pdf", "pdf_links.csv", "pdf_links.xlsx"])
            self.assertEqual(archive.read_links(folder)["https://x/0472"], "2027JP0472RL.pdf")
            # second run: nothing new
            again = FakeDownloader()
            _folder, added, _problems = archive.archive_event(event, output, again, log=lambda m: None)
            self.assertEqual((added, again.urls), (0, []))

    def test_renamed_folder_is_found_by_codex(self):
        # archive_xlsx renamed the folder like its workbook; a new PDF of the
        # event joins it instead of a new folder with the calendar name.
        event = archive.Event(2027, 1, [record("0471", date(2027, 3, 14))], "COC",
                              {"https://x/0471": Attachment("https://x/0471"),
                               "https://x/0472": Attachment("https://x/0472")})
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            renamed = output / "2027-03-14 COCM Zakopane HS140"
            renamed.mkdir()
            (renamed / "2027JP0471RL.pdf").write_bytes(b"%PDF")
            (renamed / "pdf_links.csv").write_text("pdf,url\n2027JP0471RL.pdf,https://x/0471\n", encoding="utf-8-sig")
            downloader = FakeDownloader()
            folder, added, _problems = archive.archive_event(event, output, downloader, log=lambda m: None)
            self.assertEqual((folder, added, downloader.urls), (renamed, 1, ["https://x/0472"]))
            self.assertEqual([path.name for path in output.iterdir()], [renamed.name])
            self.assertEqual(sorted(archive.read_links(renamed).values()), ["2027JP0471RL.pdf", "2027JP0472RL.pdf"])

            from openpyxl import load_workbook
            sheet = load_workbook(renamed / "pdf_links.xlsx").active
            self.assertEqual([cell.value for cell in sheet[3]], ["2027JP0472RL.pdf", "https://x/0472"])
            self.assertEqual(sheet["B3"].hyperlink.target, "https://x/0472")
            self.assertEqual(sheet["A3"].hyperlink.target, "2027JP0472RL.pdf")


if __name__ == "__main__":
    unittest.main()
