import tempfile
import unittest
from pathlib import Path

from src.config import WEB_ROOT_RELATIVE_PATH, find_web_root


class FindWebRootTests(unittest.TestCase):
    def setUp(self):
        self._temporary = tempfile.TemporaryDirectory()
        root = Path(self._temporary.name)
        self.drive_root = root / "drive"
        self.home = root / "home"
        self.drive_root.mkdir()
        self.home.mkdir()

    def tearDown(self):
        self._temporary.cleanup()

    def find(self, environ=None):
        return find_web_root(environ or {}, self.home, self.drive_root)

    def test_desktop_folder_on_drive_root_wins(self):
        desktop = self.drive_root / WEB_ROOT_RELATIVE_PATH
        desktop.mkdir(parents=True)
        (self.home / WEB_ROOT_RELATIVE_PATH).mkdir(parents=True)
        self.assertEqual(self.find()[0], desktop)

    def test_laptop_folder_in_home_directory(self):
        laptop = self.home / WEB_ROOT_RELATIVE_PATH
        laptop.mkdir(parents=True)
        self.assertEqual(self.find()[0], laptop)

    def test_existing_environment_folder_wins(self):
        (self.drive_root / WEB_ROOT_RELATIVE_PATH).mkdir(parents=True)
        custom = self.home / "custom"
        custom.mkdir()
        self.assertEqual(self.find({"SJA_WEB_ROOT": str(custom)}), (custom, "SJA_WEB_ROOT"))

    def test_nothing_found(self):
        self.assertEqual(self.find(), (None, "None"))


if __name__ == "__main__":
    unittest.main()
