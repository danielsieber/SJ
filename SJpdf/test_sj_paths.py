import tempfile
import unittest
from pathlib import Path

from sj_paths import ENV_VARIABLE, SYNCED_RELATIVE_PATH, find_data_directory


class FindDataDirectoryTests(unittest.TestCase):
    def setUp(self):
        self._temporary = tempfile.TemporaryDirectory()
        root = Path(self._temporary.name)
        self.drive_root = root / "drive"
        self.home = root / "home"
        self.project = root / "checkout" / "SJpdf"
        self.drive_root.mkdir()
        self.home.mkdir()
        self.project.mkdir(parents=True)

    def tearDown(self):
        self._temporary.cleanup()

    def find(self, environ=None):
        return find_data_directory(
            self.project, self.home, self.drive_root, environ or {}
        )

    def test_desktop_folder_on_drive_root_wins(self):
        desktop = self.drive_root / SYNCED_RELATIVE_PATH
        desktop.mkdir(parents=True)
        (self.home / SYNCED_RELATIVE_PATH).mkdir(parents=True)
        self.assertEqual(self.find(), desktop)

    def test_laptop_folder_in_home_directory(self):
        laptop = self.home / SYNCED_RELATIVE_PATH
        laptop.mkdir(parents=True)
        self.assertEqual(self.find(), laptop)

    def test_falls_back_to_checkout_data(self):
        self.assertEqual(self.find(), self.project / "data")

    def test_environment_variable_overrides_everything(self):
        (self.drive_root / SYNCED_RELATIVE_PATH).mkdir(parents=True)
        custom = self.home / "custom"
        self.assertEqual(self.find({ENV_VARIABLE: str(custom)}), custom)


if __name__ == "__main__":
    unittest.main()
