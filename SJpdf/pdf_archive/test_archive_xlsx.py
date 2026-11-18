from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive_xlsx  # noqa: E402


class NameTests(unittest.TestCase):
    def test_workbook_is_named_after_the_hill_of_the_pdfs(self):
        folder = Path("2026-09-13 FCM Einsiedeln LH")
        values = [{"Place": "Einsiedeln (SUI)", "Hill Size": 117.0},
                  {"Place": "Einsiedeln (SUI)", "Hill Size": 117}]
        self.assertEqual(archive_xlsx.workbook_name(folder, values), "2026-09-13 FCM Einsiedeln HS117")

    def test_place_of_the_calendar_without_country_code(self):
        folder = Path("2023-06-27 FCW Zakopane NH")
        values = [{"Place": "ZAKOPANE JUMPING HILLS SKI JUMPING", "Hill Size": 105}]
        self.assertEqual(archive_xlsx.workbook_name(folder, values), "2023-06-27 FCW Zakopane HS105")
        self.assertEqual(archive_xlsx.calendar_place(Path("2026-08-06 ACW Pöhla_Erzgebirge SH")),
                         "Pöhla/Erzgebirge")

    def test_folder_name_without_hill(self):
        folder = Path("2024-01-06 WCM FIS Four Hills Tournament")
        self.assertEqual(archive_xlsx.workbook_name(folder, []), folder.name)


class FolderTests(unittest.TestCase):
    def test_seasons_from_the_download_log_and_folder_without_results(self):
        with tempfile.TemporaryDirectory() as directory:
            archive = Path(directory)
            (archive / "archiv_log.csv").write_text(
                "Zeit;Saison;Ordner;Codex;PDFs neu;PDFs gesamt;Meldung\n"
                "2026-11-18 12:00:00;2024;2024-01-06 WCM FIS Four Hills Tournament;3093;1;1;\n",
                encoding="utf-8-sig",
            )
            folder = archive / "2024-01-06 WCM FIS Four Hills Tournament"
            folder.mkdir()
            (folder / "2024JP30934H.pdf").write_bytes(b"%PDF")
            self.assertEqual(archive_xlsx.folder_seasons(archive),
                             {"2024-01-06 WCM FIS Four Hills Tournament": 2024})
            results = archive_xlsx.build_archive(archive, [2024], listen_file=None, log=lambda m: None)
            self.assertEqual([(r.folder.name, r.status) for r in results],
                             [(folder.name, "keine Resultate")])
            self.assertIn("keine Resultate", (archive / "xlsx_log.csv").read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
