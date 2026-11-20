"""pdf_links.csv and pdf_links.xlsx of an event folder of the PDF archive.

Each folder lists its PDFs with their FIS link in pdf_links.csv (read by
pdf2xlsx and fis_season_check) and the same list in pdf_links.xlsx, with
both columns as clickable links (PDF in the folder, link to FIS).
"""

from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

LINKS_FILENAME = "pdf_links.csv"
LINKS_XLSX_FILENAME = "pdf_links.xlsx"
HEADERS = ("pdf", "url")
LINK_FONT = Font(color="FF0563C1", underline="single")


def read_rows(folder: Path) -> list[dict[str, str]]:
    """Rows {pdf, url} of pdf_links.csv."""
    links_file = folder / LINKS_FILENAME
    if not links_file.is_file():
        return []
    with links_file.open(encoding="utf-8-sig", newline="") as handle:
        return [
            {"pdf": row.get("pdf") or "", "url": row.get("url") or ""}
            for row in csv.DictReader(handle) if row.get("pdf")
        ]


def write_xlsx(folder: Path, rows: list[dict[str, str]] | None = None) -> None:
    """pdf_links.xlsx from the rows (default: of pdf_links.csv)."""
    rows = read_rows(folder) if rows is None else rows
    xlsx_file = folder / LINKS_XLSX_FILENAME
    if not rows:
        xlsx_file.unlink(missing_ok=True)
        return
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "pdf_links"
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append((row["pdf"], row["url"]))
        pdf_cell, url_cell = sheet[sheet.max_row]
        if (folder / row["pdf"]).is_file():
            pdf_cell.hyperlink = row["pdf"]  # relative: works when the folder moves
            pdf_cell.font = LINK_FONT
        if row["url"]:
            url_cell.hyperlink = row["url"]
            url_cell.font = LINK_FONT
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 100
    sheet.freeze_panes = "A2"
    workbook.save(xlsx_file)


def write_rows(folder: Path, rows: list[dict[str, str]]) -> None:
    """Write pdf_links.csv and pdf_links.xlsx (both removed when there are no rows)."""
    links_file = folder / LINKS_FILENAME
    if not rows:
        links_file.unlink(missing_ok=True)
    else:
        with links_file.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=HEADERS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    write_xlsx(folder, rows)


def append_row(folder: Path, pdf_name: str, url: str) -> None:
    """Add one PDF to pdf_links.csv and pdf_links.xlsx."""
    rows = [row for row in read_rows(folder) if row["pdf"] != pdf_name]
    rows.append({"pdf": pdf_name, "url": url})
    write_rows(folder, rows)
