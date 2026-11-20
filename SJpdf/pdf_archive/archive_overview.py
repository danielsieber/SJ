"""Overview of all event folders of the PDF archive as one Excel list.

One row per folder of C:\\downloaded_FIS_pdf: date, type, gender, place and
hill size (from the folder name, see archive_xlsx.py), the number of PDFs of
the Rangliste (sheet Notes of the folder's workbook: 'Anzahl PDFs',
'Gelesen', 'Nicht gelesen') and each PDF that could not be read, as a link
to the PDF in the folder. Saved as Archiv_Uebersicht.xlsx in the archive.
"""

from __future__ import annotations

import argparse
import re
import sys
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from archive_xlsx import FOLDER_PATTERN, OUTPUT_DIRECTORY, WORKBOOK_PATTERN  # noqa: E402
from pdf_to_xlsx.pdf2xlsx import find_result_pdfs  # noqa: E402

OVERVIEW_FILENAME = "Archiv_Uebersicht.xlsx"
HEADERS = ("Datum", "Typ", "Gender", "Place", "HS", "PDFs Rangliste", "Verarbeitet",
           "Nicht verarbeitet", "Status", "Ordner")
GENDERS = {"M": "M", "W": "W", "A": "A"}  # last letter of the category: men, women, mixed
LINK_FONT = Font(color="FF0563C1", underline="single")
HEADER_FILL = PatternFill("solid", fgColor="FFD9E1F2")
RED_FILL = PatternFill("solid", fgColor="FFF8CBAD")
GREEN_FILL = PatternFill("solid", fgColor="FFC6EFCE")
COUNT_PATTERN = re.compile(r"^\s*(\d+)\s*(?::\s*(.*))?$")


@dataclass
class FolderRow:
    folder: Path
    date: str = ""
    type: str = ""
    gender: str = ""
    place: str = ""
    hill_size: int | None = None
    pdfs: int = 0
    read: int = 0
    unread: list[str] = field(default_factory=list)  # PDF names (stems) not read
    status: str = ""


def _names(value) -> tuple[int, list[str]]:
    """'2: 2017JP3951RL, 2017JP3951RL1' -> (2, [names]); '0' -> (0, [])."""
    match = COUNT_PATTERN.match(str(value or "0"))
    if not match:
        return 0, []
    names = [name.strip() for name in (match.group(2) or "").split(",") if name.strip()]
    return int(match.group(1)), names


def notes_values(workbook_path: Path) -> dict[str, object]:
    """Info rows of the sheet Notes ({label: value})."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        values = {}
        if "Notes" in workbook.sheetnames:
            for row in workbook["Notes"].iter_rows(max_row=20, max_col=2, values_only=True):
                if row and row[0]:
                    values[str(row[0])] = row[1]
        return values
    finally:
        workbook.close()


def folder_row(folder: Path) -> FolderRow:
    row = FolderRow(folder)
    match = FOLDER_PATTERN.match(folder.name)
    if match:
        row.date = match["date"]
        category = match["category"]
        if category[-1:] in GENDERS:
            row.type, row.gender = category[:-1], GENDERS[category[-1]]
        else:
            row.type = category
        row.place = match["place"].replace("_", "/")
        if match["split"]:
            row.hill_size = int(match["split"].strip()[2:])
    workbooks = [path for path in folder.glob("*.xlsx") if WORKBOOK_PATTERN.match(path.name)]
    result_pdfs = find_result_pdfs(folder)
    if not workbooks:
        row.pdfs = len(result_pdfs)
        # Small hills (up to HS30) are not evaluated: no xlsx, nothing failed.
        row.status = "keine Rangliste-PDFs" if not result_pdfs else "kein xlsx (kleine Schanze?)"
        return row
    try:
        notes = notes_values(workbooks[0])
    except OSError as error:
        row.status = f"xlsx nicht lesbar ({error})"
        return row
    row.pdfs = int(notes.get("Anzahl PDFs") or len(result_pdfs))
    row.read = _names(notes.get("Gelesen"))[0]
    row.unread = _names(notes.get("Nicht gelesen"))[1]
    if "Gelesen" not in notes:  # workbook before the rows Gelesen / Nicht gelesen
        row.read = row.pdfs
    row.status = "alle verarbeitet" if not row.unread else "teilweise" if row.read else "keine verarbeitet"
    return row


def write_overview(rows: list[FolderRow], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Archiv"
    most = max((len(row.unread) for row in rows), default=0)
    headers = (*HEADERS, *(f"PDF nicht verarbeitet {n}" for n in range(1, most + 1)))
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
    for row in rows:
        sheet.append((
            datetime.strptime(row.date, "%Y-%m-%d") if row.date else None,
            row.type or None, row.gender or None, row.place or None, row.hill_size,
            row.pdfs, row.read, len(row.unread), row.status, row.folder.name,
            *(f"{name}.pdf" for name in row.unread),
        ))
        cells = sheet[sheet.max_row]
        cells[0].number_format = "yyyy-mm-dd"
        if row.unread:
            cells[7].fill = RED_FILL
        elif row.pdfs:
            cells[7].fill = GREEN_FILL
        folder_cell = cells[9]
        folder_cell.hyperlink = str(row.folder)
        folder_cell.font = LINK_FONT
        for offset, name in enumerate(row.unread):
            cell = cells[len(HEADERS) + offset]
            pdf = row.folder / f"{name}.pdf"
            if pdf.is_file():
                cell.hyperlink = str(pdf)
                cell.font = LINK_FONT
    widths = {"A": 11, "B": 7, "C": 7, "D": 28, "E": 6, "F": 9, "G": 11, "H": 11, "I": 20, "J": 50}
    for letter, width in widths.items():
        sheet.column_dimensions[letter].width = width
    for column in range(len(HEADERS) + 1, len(headers) + 1):
        sheet.column_dimensions[sheet.cell(row=1, column=column).column_letter].width = 22
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)
    workbook.close()


def build_overview(archive_directory: Path, output: Path | None = None, log=print) -> Path:
    folders = sorted(path for path in archive_directory.iterdir() if path.is_dir())
    rows = []
    for index, folder in enumerate(folders, start=1):
        if index % 100 == 0:
            log(f"[{index}/{len(folders)}] {folder.name}")
        rows.append(folder_row(folder))
    output = output or archive_directory / OVERVIEW_FILENAME
    write_overview(rows, output)
    unread = sum(len(row.unread) for row in rows)
    log(f"{len(rows)} Ordner, {sum(row.read for row in rows)} PDFs verarbeitet, "
        f"{unread} nicht: {output}")
    return output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Übersicht aller Event-Ordner des PDF-Archivs als xlsx.")
    parser.add_argument("--archive", type=Path, default=OUTPUT_DIRECTORY)
    parser.add_argument("--output", type=Path, default=None, help=f"Standard: {OVERVIEW_FILENAME} im Archiv")
    args = parser.parse_args(argv)
    try:
        build_overview(args.archive, args.output)
    except PermissionError as error:
        print(f"Kann nicht speichern, Datei offen? {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
