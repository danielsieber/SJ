import importlib.util
import sys
import tempfile
import unittest
from datetime import date
from email.message import Message
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


MODULE_PATH = Path(__file__).with_name("FIS_pdf-download.py")
SPEC = importlib.util.spec_from_file_location("fis_pdf_download", MODULE_PATH)
downloader = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = downloader
SPEC.loader.exec_module(downloader)


class PdfDownloadTests(unittest.TestCase):
    def test_codex_is_padded_to_four_digits(self):
        self.assertEqual(downloader.normalize_codex("29"), "0029")
        self.assertEqual(downloader.normalize_codex("279"), "0279")
        self.assertEqual(downloader.normalize_codex("0295"), "0295")

    def test_non_numeric_or_long_codex_is_rejected(self):
        for value in ("", "2", "2A", "12345"):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    downloader.normalize_codex(value)

    def test_event_parser_collects_all_pdf_types_and_gender(self):
        html = """
        <div class="table-row reset-padding">
          <div class="timezone-date" data-date="2026-10-25">25 Oct</div>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7757">25 Oct</a>
          <a href="https://live.fis-ski.com/jp-0279/results-pda.htm">0279 review</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7757">
            Team Large Hill
          </a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7757">GP</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7757">A</a>
          <div class="drop-btn__item">
            <span>Official Results (250 kb)</span>
            <a href="https://www.fis-ski.com/DB/v2/download/competition-attachment/11111111-1111-1111-1111-111111111111.pdf">
              Download
            </a>
          </div>
          <div class="drop-btn__item">
            <span>Start List (200 kb)</span>
            <a href="https://www.fis-ski.com/DB/v2/download/competition-attachment/22222222-2222-2222-2222-222222222222.pdf">
              Download
            </a>
          </div>
        </div>
        """

        competitions = downloader.competition_rows(html)

        self.assertEqual(len(competitions), 1)
        self.assertEqual(competitions[0].codex, "0279")
        self.assertEqual(competitions[0].gender, "A")
        self.assertEqual(competitions[0].category, "GP")
        self.assertEqual(len(competitions[0].attachments), 2)
        self.assertIn("Start List", competitions[0].attachments[1].label)

    def test_folder_name_uses_newest_same_gender_date_and_category(self):
        candidate = downloader.Candidate(
            season=2027,
            eventid="62754",
            codex="0281",
            gender="M",
            place="Einsiedeln",
        )
        competitions = [
            downloader.Competition("0281", "7759", "M", "2026-09-12", "FC", ()),
            downloader.Competition("0280", "7758", "W", "2026-09-12", "ICOC", ()),
            downloader.Competition("0283", "7761", "M", "2026-09-13", "FC", ()),
        ]

        self.assertEqual(
            downloader.folder_name_for_group(candidate, competitions),
            "2026-09-13 Einsiedeln FCM",
        )

    def test_archive_existing_files_uses_metadata_and_leaves_directories(self):
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            (output_dir / "old.pdf").write_bytes(b"%PDF-old")
            (output_dir / "myFolderName.txt").write_text(
                "2026-09-13 Einsiedeln FCM", encoding="utf-8"
            )
            existing_subfolder = output_dir / "existing-subfolder"
            existing_subfolder.mkdir()
            (existing_subfolder / "keep.pdf").write_bytes(b"%PDF-keep")

            archive_dir = downloader.archive_existing_files(output_dir)

            self.assertEqual(
                archive_dir,
                output_dir / "archiv" / "2026-09-13 Einsiedeln FCM",
            )
            self.assertEqual(
                (archive_dir / "old.pdf").read_bytes(),
                b"%PDF-old",
            )
            self.assertEqual(
                (archive_dir / "myFolderName.txt").read_text(encoding="utf-8"),
                "2026-09-13 Einsiedeln FCM",
            )
            self.assertTrue(existing_subfolder.is_dir())
            self.assertTrue((existing_subfolder / "keep.pdf").is_file())
            self.assertFalse((output_dir / "old.pdf").exists())
            self.assertFalse((output_dir / "myFolderName.txt").exists())

    def test_archive_existing_files_requires_folder_name_before_moving(self):
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            pdf_path = output_dir / "old.pdf"
            pdf_path.write_bytes(b"%PDF-old")

            with self.assertRaisesRegex(FileNotFoundError, "myFolderName.txt"):
                downloader.archive_existing_files(output_dir)

            self.assertTrue(pdf_path.is_file())

    def test_archive_existing_files_uses_unique_names_on_collision(self):
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            archive_dir = output_dir / "archiv" / "2026-09-13 Einsiedeln FCM"
            archive_dir.mkdir(parents=True)
            (archive_dir / "old.pdf").write_bytes(b"%PDF-archived")
            (output_dir / "old.pdf").write_bytes(b"%PDF-new")
            (output_dir / "myFolderName.txt").write_text(
                "2026-09-13 Einsiedeln FCM", encoding="utf-8"
            )

            downloader.archive_existing_files(output_dir)

            self.assertEqual((archive_dir / "old.pdf").read_bytes(), b"%PDF-archived")
            self.assertEqual((archive_dir / "old (2).pdf").read_bytes(), b"%PDF-new")

    @patch("fis_pdf_download.time.sleep")
    @patch("fis_pdf_download._request")
    def test_download_saves_folder_name_next_to_pdfs(self, request, _sleep):
        event_html = """
        <div class="table-row">
          <div class="timezone-date" data-date="2026-09-13">13 Sep</div>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7761">13 Sep</a>
          <a href="https://live.fis-ski.com/jp-0281/results-pda.htm">0281 review</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7761">Individual Hill</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7761">FC</a>
          <a href="/DB/general/results.html?sectorcode=JP&amp;raceid=7761">M</a>
          <div class="drop-btn__item">
            <span>Official Results</span>
            <a href="https://www.fis-ski.com/DB/v2/download/competition-attachment/11111111-1111-1111-1111-111111111111.pdf">Download</a>
          </div>
        </div>
        """

        class FakeResponse:
            def __init__(self, content, headers):
                self.content = content
                self.headers = headers

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return self.content

        html_headers = Message()
        html_headers["Content-Type"] = "text/html; charset=utf-8"
        pdf_headers = Message()
        pdf_headers["Content-Type"] = "application/pdf"
        pdf_headers["Content-Disposition"] = "inline; filename=2027JP0281RL.pdf"
        request.side_effect = [
            FakeResponse(event_html.encode("utf-8"), html_headers),
            FakeResponse(b"%PDF-1.4 test", pdf_headers),
        ]
        candidate = downloader.Candidate(
            season=2027,
            eventid="62754",
            codex="0281",
            gender="M",
            place="Einsiedeln",
        )

        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            (output_dir / "previous.pdf").write_bytes(b"%PDF-previous")
            (output_dir / "myFolderName.txt").write_text(
                "previous folder", encoding="utf-8"
            )
            downloaded = downloader.download_candidate(candidate, output_dir)
            archived_dir = output_dir / "archiv" / "previous folder"

            self.assertEqual([path.name for path in downloaded], ["2027JP0281RL.pdf"])
            self.assertEqual(
                (output_dir / "myFolderName.txt").read_text(encoding="utf-8"),
                "2026-09-13 Einsiedeln FCM",
            )
            self.assertEqual(downloaded[0].parent, output_dir)
            self.assertEqual(
                (archived_dir / "previous.pdf").read_bytes(),
                b"%PDF-previous",
            )
            self.assertEqual(
                (archived_dir / "myFolderName.txt").read_text(encoding="utf-8"),
                "previous folder",
            )

    def test_workbook_lookup_preserves_padded_codex_and_event(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "raceid.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "all"
            sheet.append(
                [
                    "season",
                    "date",
                    "eventid",
                    "raceid",
                    "codex",
                    "place",
                    "category",
                    "gender",
                ]
            )
            sheet.append(
                [2027, date(2026, 10, 25), "62788", "7757", "0279", "Klingenthal", "GP", "A"]
            )
            workbook.save(path)
            workbook.close()

            candidates = downloader.read_workbook_candidates(path, "0279")

        self.assertEqual(
            candidates,
            [
                downloader.Candidate(
                    season=2027,
                    eventid="62788",
                    codex="0279",
                    gender="A",
                    place="Klingenthal",
                    event_date=date(2026, 10, 25),
                    category="GP",
                )
            ],
        )

    def test_filename_comes_from_fis_content_disposition(self):
        headers = Message()
        headers["Content-Disposition"] = "inline; filename=2027JP0295RL.pdf"
        self.assertEqual(
            downloader._pdf_filename(
                headers,
                "https://www.fis-ski.com/DB/v2/download/competition-attachment/uuid.pdf",
            ),
            "2027JP0295RL.pdf",
        )


if __name__ == "__main__":
    unittest.main()
