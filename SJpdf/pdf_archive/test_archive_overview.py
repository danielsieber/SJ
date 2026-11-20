from pathlib import Path
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive_overview  # noqa: E402


def make_folder(archive: Path, name: str, pdfs=(), notes=None) -> Path:
    folder = archive / name
    folder.mkdir()
    for pdf in pdfs:
        (folder / f"{pdf}.pdf").write_bytes(b"%PDF")
    if notes is not None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Notes"
        for label, value in notes.items():
            sheet.append((label, value))
        workbook.save(folder / f"{name}.xlsx")
    return folder


class OverviewTests(unittest.TestCase):
    def test_rows_and_links_of_unread_pdfs(self):
        with tempfile.TemporaryDirectory() as directory:
            archive = Path(directory)
            partly = make_folder(archive, "2017-02-04 COCM Erzurum HS140",
                                 ["2017JP3951RL", "2017JP3951RL1", "2017JP3951SLR1"],
                                 {"Anzahl PDFs": 3, "Gelesen": "2: 2017JP3951RL1, 2017JP3951SLR1",
                                  "Nicht gelesen": "1: 2017JP3951RL"})
            make_folder(archive, "2027-03-21 WCM Planica FH", ["2027JP0001PROG"])
            output = archive_overview.build_overview(archive, log=lambda m: None)

            rows = list(load_workbook(output).active.iter_rows(values_only=True))
            self.assertEqual(rows[0][:10], archive_overview.HEADERS)
            erzurum = rows[1]
            self.assertEqual(erzurum[1:9], ("COC", "M", "Erzurum", 140, 3, 2, 1, "teilweise"))
            self.assertEqual(erzurum[10], "2017JP3951RL.pdf")
            self.assertEqual(rows[2][1:9], ("WC", "M", "Planica", None, 0, 0, 0, "keine Rangliste-PDFs"))
            sheet = load_workbook(output).active
            self.assertEqual(sheet["K2"].hyperlink.target, str(partly / "2017JP3951RL.pdf"))


if __name__ == "__main__":
    unittest.main()
