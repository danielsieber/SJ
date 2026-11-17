"""Sheet 'Notes': where the files are and what the program noticed.

The sheet is rewritten on every run. It is the first and the active sheet
when the workbook is opened, so warnings are seen at once.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook import Workbook
from openpyxl.workbook.properties import CalcProperties


SHEET_NAME = "Notes"
LINKS_FILENAME = "pdf_links.csv"  # written by pdf_download/FIS_pdf-download.py
HEADER_FILL = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
WARNING_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
LINK_FONT = Font(color="FF0563C1", underline="single")
SOURCE_PATTERN = re.compile(r"^(\d{4}JP\d{4}[A-Z0-9]*|Ranglisten|Daten|SJ-Listen)\b", re.I)

# Folder and file name of the workbook itself, calculated by Excel.
FILENAME = 'CELL("filename",$A$1)'
FOLDER_FORMULA = f'=LEFT({FILENAME},FIND("[",{FILENAME})-2)'
FILE_FORMULA = (
    f'=MID({FILENAME},FIND("[",{FILENAME})+1,'
    f'FIND("]",{FILENAME})-FIND("[",{FILENAME})-1)'
)


@dataclass(frozen=True)
class Note:
    kind: str  # "Warnung", "Fehler" or "Info"
    message: str

    @property
    def source(self) -> str:
        match = SOURCE_PATTERN.match(self.message)
        return match.group(1) if match else ""


def read_pdf_links(folder: Path) -> dict[str, str] | None:
    """Return {pdf file name: FIS link} or None when the link file is missing."""
    links_file = folder / LINKS_FILENAME
    if not links_file.is_file():
        return None
    links = {}
    with links_file.open(encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            if row.get("pdf") and row.get("url"):
                links[row["pdf"]] = row["url"]
    return links


def write_notes_sheet(
    workbook: Workbook,
    created: datetime,
    pdfs: Sequence[tuple[Path, str | None]],
    notes: Sequence[Note],
    copy_name: str | None = None,
) -> None:
    """Add the Notes sheet (first) and make it the active sheet.

    copy_name is the file name of the competition copy (e.g.
    '2026-09-13 FCM Einsiedeln HS117.xlsx'), None when no copy was saved.
    """
    worksheet = workbook.create_sheet(SHEET_NAME, 0)
    bold = Font(bold=True)
    worksheet["A1"] = "Notes"
    worksheet["A1"].font = Font(bold=True, size=14)

    folders = sorted({str(path.parent) for path, _url in pdfs})
    info = (
        ("Erstellt", created),
        ("xlsx-Ordner", FOLDER_FORMULA),
        ("xlsx-Datei", FILE_FORMULA),
        ("Kopie gespeichert als", copy_name or "keine (siehe Meldungen)"),
        ("PDF-Ordner", "; ".join(folders) or None),
        ("Anzahl PDFs", len(pdfs)),
    )
    for row, (label, value) in enumerate(info, start=3):
        worksheet.cell(row=row, column=1, value=label).font = bold
        cell = worksheet.cell(row=row, column=2, value=value)
        cell.alignment = Alignment(horizontal="left")
        if isinstance(value, datetime):
            cell.number_format = "yyyy-mm-dd hh:mm:ss"
        if label == "Kopie gespeichert als" and copy_name is None:
            cell.fill = WARNING_FILL

    row = 3 + len(info) + 1
    row = _header(worksheet, row, ("Art", "Quelle", "Meldung"))
    for note in notes or (Note("Info", "Keine Anmerkungen."),):
        values = (note.kind, note.source, note.message)
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row, column=column, value=value)
            cell.alignment = Alignment(wrap_text=column == 3, vertical="top")
            if note.kind != "Info":
                cell.fill = WARNING_FILL
        row += 1

    row = _header(worksheet, row + 1, ("PDF", "Speicherort", "FIS-Link"))
    for path, url in pdfs:
        worksheet.cell(row=row, column=1, value=path.stem)
        worksheet.cell(row=row, column=2, value=str(path))
        if url:
            cell = worksheet.cell(row=row, column=3, value=url)
            cell.hyperlink = url
            cell.font = LINK_FONT
        row += 1

    worksheet.column_dimensions["A"].width = 21
    worksheet.column_dimensions["B"].width = 70
    worksheet.column_dimensions["C"].width = 100

    # Excel calculates the path formulas when the file is opened.
    workbook.calculation = CalcProperties(fullCalcOnLoad=True)
    workbook.active = workbook.sheetnames.index(SHEET_NAME)
    for sheet in workbook.worksheets:
        sheet.sheet_view.tabSelected = sheet.title == SHEET_NAME


def _header(worksheet, row: int, labels: Sequence[str]) -> int:
    for column, label in enumerate(labels, start=1):
        cell = worksheet.cell(row=row, column=column, value=label)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
    return row + 1
