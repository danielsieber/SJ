import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from src.utils.fis_pdf import (
    FIS_PDF_DOCUMENT_TYPES,
    _download_pdf_batch,
    build_fis_pdf_url,
    fis_season,
)


class FisPdfTests(unittest.TestCase):
    def test_season_changes_in_june(self):
        self.assertEqual(fis_season(date(2025, 5, 31)), 2025)
        self.assertEqual(fis_season(date(2025, 6, 1)), 2026)

    def test_builds_official_url(self):
        self.assertEqual(
            build_fis_pdf_url(2026, "3204", "RL1"),
            "https://medias1.fis-ski.com/pdf/2026/JP/3204/"
            "2026JP3204RL1.pdf",
        )

    def test_batch_skips_existing_files_and_reports_new_names(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "Data_pdf"
            destination.mkdir()
            (destination / "2026-3204-RL.pdf").write_bytes(b"%PDF existing")

            def successful_download(_url, target):
                target.write_bytes(b"%PDF downloaded")
                return True, None

            with patch(
                "src.utils.fis_pdf.download_fis_pdf",
                side_effect=successful_download,
            ) as download, patch(
                "src.utils.fis_pdf._show_download_popup"
            ) as popup:
                _download_pdf_batch(((2026, "3204"),), destination)

            self.assertEqual(download.call_count, len(FIS_PDF_DOCUMENT_TYPES) - 1)
            requested_names = {call.args[1].name for call in download.call_args_list}
            self.assertNotIn("2026-3204-RL.pdf", requested_names)
            popup.assert_called_once()
            downloaded_names, popup_destination = popup.call_args.args
            self.assertEqual(popup_destination, destination)
            self.assertEqual(set(downloaded_names), requested_names)


if __name__ == "__main__":
    unittest.main()
