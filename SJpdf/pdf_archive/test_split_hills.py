from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive_xlsx  # noqa: E402
import split_hills  # noqa: E402

EVENT = "2017-02-05 COCM Erzurum - Kiremitliktepe LH"
# Hill size of each PDF (RL1 of 3952 cannot be read: the codex keeps its RL value).
HILL_SIZES = {"2017JP3951RL": 109, "2017JP3951RLT": 109, "2017JP3952RL": 140, "2017JP3952RL1": None}


def fake_extract(path):
    size = HILL_SIZES.get(path.stem)
    if size is None:
        raise split_hills.PdfParseError(f"{path.name}: no text")
    return {"Hill Size": size}


def make_event(archive: Path, names) -> Path:
    folder = archive / EVENT
    folder.mkdir()
    for name in names:
        (folder / f"{name}.pdf").write_bytes(b"%PDF")
    (folder / "pdf_links.csv").write_text(
        "pdf,url\n" + "".join(f"{name}.pdf,https://x/{name}\n" for name in names), encoding="utf-8-sig"
    )
    (folder / f"{EVENT[:10]} COCM Erzurum HS140.xlsx").write_bytes(b"old")
    return folder


class SplitTests(unittest.TestCase):
    def test_codexes_with_different_hill_sizes_get_own_folders(self):
        names = ["2017JP3951RL", "2017JP3951RLT", "2017JP3951SLR1", "2017JP3952RL", "2017JP3952RL1", "2017JP3952STC"]
        with tempfile.TemporaryDirectory() as directory, patch.object(split_hills, "extract_data", fake_extract):
            archive = Path(directory)
            folder = make_event(archive, names)
            result = split_hills.split_by_hill_size(folder, log=lambda m: None)

            self.assertEqual([path.name for path in result], [f"{EVENT} HS109", f"{EVENT} HS140"])
            self.assertFalse(folder.exists())
            self.assertEqual(sorted(p.name for p in result[0].glob("*.pdf")),
                             ["2017JP3951RL.pdf", "2017JP3951RLT.pdf", "2017JP3951SLR1.pdf"])
            self.assertEqual(sorted(p.name for p in result[1].glob("*.pdf")),
                             ["2017JP3952RL.pdf", "2017JP3952RL1.pdf", "2017JP3952STC.pdf"])
            links = split_hills._read_links(result[1])
            self.assertEqual({row["pdf"]: row["url"] for row in links}["2017JP3952RL1.pdf"], "https://x/2017JP3952RL1")

            # A PDF downloaded later into the event folder joins its codex.
            folder.mkdir()
            (folder / "2017JP3951RTRIA.pdf").write_bytes(b"%PDF")
            split_hills.split_by_hill_size(folder, log=lambda m: None)
            self.assertTrue((result[0] / "2017JP3951RTRIA.pdf").exists())
            self.assertFalse(folder.exists())

    def test_leftover_links_of_an_interrupted_split_are_moved(self):
        with tempfile.TemporaryDirectory() as directory:
            archive = Path(directory)
            folder = archive / EVENT
            split = archive / f"{EVENT} HS109"
            folder.mkdir()
            split.mkdir()
            (split / "2017JP3951RL.pdf").write_bytes(b"%PDF")
            (folder / "pdf_links.csv").write_text("pdf,url\n2017JP3951RL.pdf,https://x/a\n", encoding="utf-8-sig")
            self.assertEqual(split_hills.split_by_hill_size(folder, log=lambda m: None), [split])
            self.assertFalse(folder.exists())
            self.assertEqual(split_hills._read_links(split)[0]["url"], "https://x/a")

    def test_one_hill_size_stays(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(split_hills, "extract_data", fake_extract):
            folder = make_event(Path(directory), ["2017JP3951RL", "2017JP3951RLT"])
            self.assertEqual(split_hills.split_by_hill_size(folder, log=lambda m: None), [folder])
            self.assertEqual(len(list(folder.glob("*.xlsx"))), 1)

    def test_names_of_split_folders(self):
        folder = Path(f"{EVENT} HS109")
        self.assertEqual(archive_xlsx.calendar_place(folder), "Erzurum - Kiremitliktepe")
        from datetime import datetime
        values = [{"Place": "Erzurum (TUR)", "Hill Size": 109, "Datum": datetime(2017, 2, 3)},
                  {"Place": "Erzurum (TUR)", "Hill Size": 109, "Datum": datetime(2017, 2, 4)}]
        self.assertEqual(archive_xlsx.workbook_name(folder, values), "2017-02-04 COCM Erzurum HS109")


class RenameTests(unittest.TestCase):
    def test_folder_gets_the_workbook_name_with_codex_when_taken(self):
        with tempfile.TemporaryDirectory() as directory:
            archive = Path(directory)
            first = archive / "2017-04-02 FCM Rasnov NH"
            first.mkdir()
            (first / "2017JP3850RL.pdf").write_bytes(b"%PDF")
            second = archive / "2017-04-02 FCM Rasnov NH 2"
            second.mkdir()
            (second / "2017JP3851RL.pdf").write_bytes(b"%PDF")
            name = "2017-04-02 FCM Rasnov HS100"
            renamed = archive_xlsx.rename_folder(first, name, first.glob("*.pdf"))
            self.assertEqual(renamed.name, name)
            other = archive_xlsx.rename_folder(second, name, second.glob("*.pdf"))
            self.assertEqual(other.name, f"{name} 3851")
            # Next run: both keep their names.
            self.assertEqual(archive_xlsx.rename_folder(renamed, name, renamed.glob("*.pdf")), renamed)
            self.assertEqual(archive_xlsx.rename_folder(other, name, other.glob("*.pdf")), other)
            self.assertEqual(archive_xlsx.calendar_place(other), "Rasnov")
            self.assertEqual(archive_xlsx.pdf_season(other), 2017)


if __name__ == "__main__":
    unittest.main()
