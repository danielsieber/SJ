"""Locate the shared SJpdf data folder on every machine.

Search order:
1. Environment variable ``SJPDF_DATA_DIR`` (if set, always used)
2. ``C:\\CLOUD-DANIEL\\SJA\\SJpdf\\data`` (Desktop)
3. ``%USERPROFILE%\\CLOUD-DANIEL\\SJA\\SJpdf\\data`` (Laptop)
4. ``data`` inside this checkout (fallback)

Running ``python sj_paths.py`` prints the folder that is used.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Mapping

ENV_VARIABLE = "SJPDF_DATA_DIR"
PROJECT_ROOT = Path(__file__).resolve().parent
SYNCED_RELATIVE_PATH = Path("CLOUD-DANIEL", "SJA", "SJpdf", "data")


def candidate_data_directories(
    project_root: Path = PROJECT_ROOT,
    home_directory: Path | None = None,
    drive_root: Path | None = None,
) -> list[Path]:
    """Return the synced and checkout data folders in search order."""
    home_directory = Path.home() if home_directory is None else home_directory
    drive_root = Path(home_directory.anchor) if drive_root is None else drive_root
    return [
        drive_root / SYNCED_RELATIVE_PATH,
        home_directory / SYNCED_RELATIVE_PATH,
        project_root / "data",
    ]


def find_data_directory(
    project_root: Path = PROJECT_ROOT,
    home_directory: Path | None = None,
    drive_root: Path | None = None,
    environ: Mapping[str, str] | None = None,
) -> Path:
    """Return the first existing data folder, or the checkout folder."""
    environ = os.environ if environ is None else environ
    configured = environ.get(ENV_VARIABLE, "").strip()
    if configured:
        return Path(configured).expanduser()

    candidates = candidate_data_directories(project_root, home_directory, drive_root)
    for directory in candidates:
        if directory.is_dir():
            return directory
    return candidates[-1]


DATA_DIRECTORY = find_data_directory()


if __name__ == "__main__":
    print(DATA_DIRECTORY)
