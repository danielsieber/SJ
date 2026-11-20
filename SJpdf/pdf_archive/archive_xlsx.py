"""Write the Excel file of pdf2xlsx into every event folder of the PDF archive.

Each folder of C:\\downloaded_FIS_pdf (see fis_pdf_archive.py) is one event,
the same PDFs as in data\\pdf for latest event.xlsx. The workbook is saved in
the folder under the name of the competition, e.g.

    C:\\downloaded_FIS_pdf\\2026-09-13 FCM Einsiedeln LH\\2026-09-13 FCM Einsiedeln HS117.xlsx

(date of the last competition, category and gender of the folder, SchanzenID
of Liste_Schanzen_ID). The folder gets the same name:

    C:\downloaded_FIS_pdf\2026-09-13 FCM Einsiedeln HS117\2026-09-13 FCM Einsiedeln HS117.xlsx

When another folder has that name already, the codex of the first
competition is added ('2017-04-02 FCM Rasnov HS100 3851'). Folders without
result PDFs keep the name of the FIS calendar. A folder whose workbook is newer than its PDFs and SJ-Listen.xlsm is skipped;
--force writes all again. Folders without result PDFs (tournament standings,
cancelled events) are only listed in the log xlsx_log.csv.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import warnings as python_warnings
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
from pdf_to_xlsx.pdf2xlsx import create_workbook, find_result_pdfs  # noqa: E402
from pdf_to_xlsx.pdf2xlsx_Daten import PdfParseError, extract_data  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
from split_hills import CODEX_PATTERN, codex_of, split_by_hill_size  # noqa: E402
from pdf_to_xlsx.pdf2xlsx_Listen import (  # noqa: E402
    SJ_LISTEN_FILE,
    SjListen,
    hill_name,
    read_sj_listen,
)

OUTPUT_DIRECTORY = Path(r"C:\downloaded_FIS_pdf")
ARCHIVE_LOG = "archiv_log.csv"  # written by fis_pdf_archive.py: season of each folder
LOG_FILENAME = "xlsx_log.csv"
MAX_SKIPPED_HILL_SIZE = 30  # events on hills up to HS30 are not evaluated
LOG_HEADERS = ("Zeit", "Saison", "Ordner", "Datei", "Status", "Fehlende JumperID",
               "Fehlende SchanzenID", "Warnungen", "Meldung")
WORKBOOK_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2} .+\.xlsx$")  # written here
FOLDER_PATTERN = re.compile(
    r"^(?P<date>\d{4}-\d{2}-\d{2}) (?P<category>\S+) (?P<place>.+?)(?: (?:SH|MH|NH|LH|FH))?(?P<split> HS\d+)?(?: \d{4})?$"
)
JUMPER_WARNING = re.compile(r"^Ranglisten: JumperID ")
HILL_WARNING = re.compile(r"^Daten: SchanzenID ")


@dataclass
class FolderResult:
    folder: Path
    status: str  # neu, aktuell, keine Resultate, Fehler
    workbook: Path | None = None
    warnings: list[str] = field(default_factory=list)
    message: str = ""

    @property
    def missing_jumpers(self) -> int:
        return sum(1 for warning in self.warnings if JUMPER_WARNING.match(warning))

    @property
    def missing_hills(self) -> int:
        return sum(1 for warning in self.warnings if HILL_WARNING.match(warning))


def folder_seasons(archive_directory: Path) -> dict[str, int]:
    """{folder name: season}: from the PDF names ('2017JP3951RL.pdf' -> 2017),
    else from the download log of fis_pdf_archive.py."""
    seasons: dict[str, int] = {}
    log_file = archive_directory / ARCHIVE_LOG
    if log_file.is_file():
        with log_file.open(encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle, delimiter=";"):
                if row.get("Ordner") and row.get("Saison"):
                    seasons.setdefault(row["Ordner"], int(row["Saison"]))
    if archive_directory.is_dir():
        for folder in archive_directory.iterdir():
            if folder.is_dir():
                season = pdf_season(folder)
                if season is not None:
                    seasons[folder.name] = season
    return seasons


def pdf_season(folder: Path) -> int | None:
    """Season of the PDFs of a folder ('2022JP3813RL.pdf' -> 2022), the most frequent."""
    years = Counter(int(pdf.name[:4]) for pdf in folder.glob("*.pdf") if CODEX_PATTERN.match(pdf.stem))
    return years.most_common(1)[0][0] if years else None


def rename_folder(folder: Path, name: str, pdfs: Iterable[Path]) -> Path:
    """Rename the folder to the workbook name; with the codex of its first
    competition when another folder has that name. Keep the folder when it
    cannot be renamed (a file is open)."""
    if folder.name == name:
        return folder
    target = folder.parent / name
    if target.exists():
        codexes = sorted(codex for pdf in pdfs if (codex := codex_of(pdf)))
        if not codexes:
            return folder
        target = folder.parent / f"{name} {codexes[0][-4:]}"
        if target == folder or target.exists():
            return folder
    try:
        folder.rename(target)
    except OSError:
        return folder
    return target


def calendar_place(folder: Path) -> str | None:
    """Place of the FIS calendar in the folder name ('/' was saved as '_')."""
    match = FOLDER_PATTERN.match(folder.name)
    return match["place"].replace("_", "/") if match else None


def pdf_place(values: dict, folder: Path) -> str | None:
    """Place of the PDF with country code ('Engelberg (SUI)'), else of the calendar."""
    place = str(values.get("Place") or "")
    return place if re.search(r"\([A-Z]{3}\)", place) else calendar_place(folder)


def workbook_name(folder: Path, values: Iterable[dict], listen: SjListen | None = None) -> str:
    """'2026-09-13 FCM Einsiedeln HS117': date, category and gender of the
    folder, SchanzenID of the PDFs (most frequent); the folder name when the
    PDFs give no hill."""
    match = FOLDER_PATTERN.match(folder.name)
    hills = Counter(
        hill for value in values
        if (hill := hill_name(pdf_place(value, folder), value.get("Hill Size"))) is not None
    )
    if listen is not None:  # SchanzenID of Liste_Schanzen_ID (column SchanzenID_pdf)
        mapped: Counter = Counter()
        for hill, count in hills.items():
            mapped[listen.hill_id(hill) or hill] += count
        hills = mapped
    if match is None or not hills:
        return folder.name
    day = match["date"]
    if match["split"]:
        # Folder of one hill of an event (split_hills.py): its last PDF date.
        dates = [value["Datum"] for value in values if value.get("Datum")]
        if dates:
            day = max(dates).strftime("%Y-%m-%d")
    name = f"{day} {match['category']} {hills.most_common(1)[0][0]}"
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")


def _up_to_date(workbook: Path, pdfs: Iterable[Path], listen_file: Path | None) -> bool:
    if not workbook.is_file():
        return False
    newest = max((pdf.stat().st_mtime for pdf in pdfs), default=0.0)
    if listen_file is not None and listen_file.is_file():
        newest = max(newest, listen_file.stat().st_mtime)
    return workbook.stat().st_mtime > newest


def build_folder(
    folder: Path, listen: SjListen | None, listen_file: Path | None = None, force: bool = False
) -> FolderResult:
    """Write (or keep) the workbook of one event folder."""
    pdfs = find_result_pdfs(folder)
    if not pdfs:
        return FolderResult(folder, "keine Resultate", message="kein RL-, RLQ-, RLT- oder RTRIA-PDF")
    values = []
    for pdf in pdfs:
        try:
            values.append(extract_data(pdf))
        except PdfParseError:
            continue  # reported by create_workbook
    old = [path for path in folder.glob("*.xlsx") if WORKBOOK_PATTERN.match(path.name)]
    hill_sizes = [value["Hill Size"] for value in values if value.get("Hill Size") is not None]
    if hill_sizes and max(hill_sizes) <= MAX_SKIPPED_HILL_SIZE:
        for path in old:
            path.unlink(missing_ok=True)
        return FolderResult(folder, "kleine Schanze",
                            message=f"HS{max(hill_sizes)}: Schanzen bis HS{MAX_SKIPPED_HILL_SIZE} werden nicht ausgewertet")
    name = workbook_name(folder, values, listen)
    if FOLDER_PATTERN.match(name):  # a hill was read: folder named like the workbook
        renamed = rename_folder(folder, name, pdfs)
        if renamed != folder:
            folder = renamed
            pdfs = find_result_pdfs(folder)
            old = [path for path in folder.glob("*.xlsx") if WORKBOOK_PATTERN.match(path.name)]
    file_name = folder.name if folder.name.startswith(f"{name} ") or folder.name == name else name
    workbook = folder / f"{file_name}.xlsx"
    if not force and _up_to_date(workbook, folder.glob("*.pdf"), listen_file) and old == [workbook]:
        return FolderResult(folder, "aktuell", workbook)
    try:
        with python_warnings.catch_warnings():
            python_warnings.simplefilter("ignore")
            warnings = create_workbook(
                pdfs, workbook, listen, competition=name, save_copy=False,
                place=calendar_place(folder),
            )
    except PermissionError as error:
        return FolderResult(folder, "Fehler", workbook, message=f"{workbook.name} ist offen ({error})")
    except (OSError, PdfParseError) as error:
        return FolderResult(folder, "Fehler", workbook, message=str(error))
    for path in old:
        if path != workbook:
            path.unlink(missing_ok=True)  # name changed (e.g. hill size corrected)
    return FolderResult(folder, "neu", workbook, warnings)


def _log_row(output_directory: Path, values) -> None:
    log_file = output_directory / LOG_FILENAME
    new_file = not log_file.is_file()
    with log_file.open("a", encoding="utf-8-sig" if new_file else "utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        if new_file:
            writer.writerow(LOG_HEADERS)
        writer.writerow(values)


def build_archive(
    archive_directory: Path,
    seasons: Iterable[int] | None = None,
    listen_file: Path | None = SJ_LISTEN_FILE,
    force: bool = False,
    log: Callable[[str], None] = print,
) -> list[FolderResult]:
    """Workbooks of all folders (of the given seasons) of the archive."""
    season_of = folder_seasons(archive_directory)
    wanted = set(seasons) if seasons else None
    # One folder per hill size when the codexes of an event differ (split_hills.py).
    for folder in sorted(archive_directory.iterdir()):
        if folder.is_dir() and (
            wanted is None or season_of.get(folder.name) in wanted
        ):
            for split in split_by_hill_size(folder, log):
                season_of.setdefault(split.name, season_of.get(folder.name))
    folders = sorted(
        path for path in archive_directory.iterdir()
        if path.is_dir() and (wanted is None or season_of.get(path.name) in wanted)
    )
    listen = read_sj_listen(listen_file) if listen_file is not None else None
    results = []
    for index, folder in enumerate(folders, start=1):
        result = build_folder(folder, listen, listen_file, force)
        results.append(result)
        season_of.setdefault(result.folder.name, season_of.get(folder.name))
        details = ""
        if result.status == "neu":
            details = (f"{len(result.warnings)} Warnungen, {result.missing_jumpers} JumperID und "
                       f"{result.missing_hills} SchanzenID fehlen")
        renamed = f" -> {result.folder.name}" if result.folder != folder else ""
        log(f"[{index}/{len(folders)}] {folder.name}{renamed}: {result.status} {details or result.message}".rstrip())
        if result.status != "aktuell":
            _log_row(archive_directory, (
                datetime.now().replace(microsecond=0).isoformat(sep=" "),
                season_of.get(result.folder.name, ""), result.folder.name,
                result.workbook.name if result.workbook else "", result.status,
                result.missing_jumpers, result.missing_hills, len(result.warnings),
                result.message or " | ".join(result.warnings)[:30000],
            ))
    counts = Counter(result.status for result in results)
    log("Fertig: " + ", ".join(f"{count} {status}" for status, count in sorted(counts.items())))
    return results


def main(argv: list[str] | None = None) -> int:
    sys.path.insert(0, str(PROJECT_ROOT / "season_check"))
    from fis_season_check import parse_seasons  # noqa: E402 - only for the arguments

    parser = argparse.ArgumentParser(description="xlsx von pdf2xlsx in jeden Event-Ordner des PDF-Archivs schreiben.")
    parser.add_argument("--season", nargs="*", help="Saisons, z. B. 2026 oder 2024-2026 (Standard: alle Ordner)")
    parser.add_argument("--archive", type=Path, default=OUTPUT_DIRECTORY)
    parser.add_argument("--listen", type=Path, default=SJ_LISTEN_FILE)
    parser.add_argument("--force", action="store_true", help="auch aktuelle xlsx neu schreiben")
    args = parser.parse_args(argv)
    seasons = parse_seasons(args.season) if args.season else None
    build_archive(args.archive, seasons, args.listen, args.force)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
