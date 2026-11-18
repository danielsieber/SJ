"""Read the start order from FIS start lists (SLR1 1st round, SLRF final round).

A start list is printed in start order. The order is usually the bib order
(round 1) or the reverse ranking of round 1 (final round), but World Cup
competitions around New Year use a zip order (50., 1., 49., 2., ...), and
Grand Prix / World Cup start lists begin with the highest bib. The start
lists therefore decide the jump order in the Ranglisten sheet.

Each bib is read from the BIB column (position of the header label); the
column START ORD. (Swiss Timing final round) is checked against the printed
order. A start list that cannot be read without doubt raises
``StartListError``.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pymupdf
from openpyxl.styles import Alignment, Font
from openpyxl.workbook import Workbook

if __package__:
    from .pdf2xlsx_Daten import PdfParseError
    from .pdf2xlsx_Resultate import GENDER_SUFFIX, HEADER_FILL, looks_like_athlete
else:
    from pdf2xlsx_Daten import PdfParseError
    from pdf2xlsx_Resultate import GENDER_SUFFIX, HEADER_FILL, looks_like_athlete


START_LIST_COLUMNS = ("Start Order", "Bib", "Name")
# Start list per round of a competition (RL): (first round, final round).
START_LIST_SUFFIXES = ("SLR1", "SLRF")
FOOTER_PATTERN = re.compile(r"^\d{1,2} [A-Z]{3} \d{4} / .+ / \d+$")
NUMBER_PATTERN = re.compile(r"^\*?\d{1,3}(?:-\d)?$")  # team events: '8-1'; KO lucky loser: '*35'
MAX_LABEL_DISTANCE = 12  # pt between a number and its column label centre
MAX_NAME_DISTANCE = 12  # pt between the bib and the athlete name (vertical)


class StartListError(PdfParseError):
    """Raised when the start order cannot be read without doubt."""


def start_list_paths(result_path: Path) -> tuple[Path | None, Path | None]:
    """Return the SLR1 and SLRF PDFs next to a competition PDF (...RL.pdf)."""
    match = re.fullmatch(r"(.*\d)RL", result_path.stem, re.I)
    if match is None:
        return None, None
    paths = [
        result_path.with_name(f"{match.group(1)}{suffix}{result_path.suffix}")
        for suffix in START_LIST_SUFFIXES
    ]
    return tuple(path if path.exists() else None for path in paths)  # type: ignore[return-value]


def _centre(box) -> float:
    return (box[0] + box[2]) / 2


def _middle(box) -> float:
    return (box[1] + box[3]) / 2


def _page_rows(page: pymupdf.Page, page_number: int, name: str) -> list[dict[str, Any]]:
    words = page.get_text("words")
    bib_labels = [word for word in words if word[4].upper() in ("BIB", "STNR")]  # StNr: WinSpringen
    if not bib_labels:
        return []
    if len(bib_labels) > 1:
        raise StartListError(f"{name}: page {page_number}: more than one BIB column.")
    bib_label = bib_labels[0]

    spans = [
        span
        for block in page.get_text("dict")["blocks"]
        for line in block.get("lines", [])
        for span in line["spans"]
        if span["text"].strip()
    ]
    footer_tops = [span["bbox"][1] for span in spans if FOOTER_PATTERN.match(span["text"].strip())]
    bottom = min(footer_tops) if footer_tops else page.rect.height

    labels = {"Bib": _centre(bib_label)}
    name_left = None
    for word in words:
        same_header = abs(word[1] - bib_label[1]) < 15
        if word[4].upper() in ("START", "ORD.", "ORDER") and same_header and word[2] < bib_label[0]:
            labels["Start Order"] = _centre(word)
        if word[4].upper() == "NAME" and abs(word[1] - bib_label[1]) < 3:
            name_left = word[0]
    if name_left is None:
        raise StartListError(f"{name}: page {page_number}: NAME column not found.")

    bibs, orders = [], []
    for word in words:
        if not NUMBER_PATTERN.match(word[4]) or not bib_label[3] < word[1] < bottom - 1:
            continue
        if any(
            other[4].casefold() in ("group", "gruppe") and abs(other[1] - word[1]) < 2 and other[0] < word[0]
            for other in words
        ):
            continue  # heading 'Group 1 Women' of a team start list
        column = min(labels, key=lambda key: abs(labels[key] - _centre(word)))
        if abs(labels[column] - _centre(word)) > MAX_LABEL_DISTANCE:
            continue
        (bibs if column == "Bib" else orders).append(word)

    rows = []
    for word in sorted(bibs, key=_middle):
        y = _middle(word)
        # A span may also start with the bib ('21 JANCAR Ziga').
        column = [
            {**span, "text": re.sub(rf"^{re.escape(word[4])}\s+", "", span["text"].strip())}
            for span in spans
            if span["bbox"][0] < name_left + 40
            and (name_left - 5 <= span["bbox"][0] or span["text"].strip().startswith(f"{word[4]} "))
            and span["bbox"][2] > name_left + 5
            and abs(_middle(span["bbox"]) - y) < MAX_NAME_DISTANCE
        ]
        names = [span for span in column if span["flags"] & 16]
        if not names:
            # Name not printed bold: the closest 'SURNAME Firstname'.
            names = [span for span in column if looks_like_athlete(span["text"])]
        if not names:
            # Name in capitals only ('BARKATS BRANDINA DAVID'): the span at the NAME label.
            names = [span for span in column if abs(span["bbox"][0] - name_left) < 3]
        if not names:
            raise StartListError(f"{name}: page {page_number}: no name for bib {word[4]}.")
        athlete = min(names, key=lambda span: abs(_middle(span["bbox"]) - y))
        bib_text = word[4].lstrip("*")
        bib = bib_text if "-" in bib_text else int(bib_text)
        row = {"Bib": bib, "Name": GENDER_SUFFIX.sub("", athlete["text"].strip())}
        if "Start Order" in labels:
            same_line = [order for order in orders if abs(_middle(order) - y) < 3]
            if len(same_line) != 1:
                raise StartListError(
                    f"{name}: page {page_number}: start order of bib {word[4]} not found."
                )
            row["Start Order"] = int(same_line[0][4])
        rows.append(row)
    return rows


def extract_start_list(pdf_path: Path) -> list[dict[str, Any]]:
    """Return the athletes of a start list in start order."""
    name = pdf_path.name
    try:
        with pymupdf.open(pdf_path) as document:
            rows = [
                row
                for page_number, page in enumerate(document, start=1)
                for row in _page_rows(page, page_number, name)
            ]
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise StartListError(f"{name}: cannot read PDF: {error}") from error

    if not rows:
        raise StartListError(f"{name}: no athletes found in the start list.")
    bibs = [row["Bib"] for row in rows]
    duplicates = sorted({bib for bib in bibs if bibs.count(bib) > 1})
    if duplicates:
        raise StartListError(f"{name}: duplicate bibs {duplicates}.")
    for position, row in enumerate(rows, start=1):
        printed = row.get("Start Order")
        if printed is not None and printed != position:
            raise StartListError(
                f"{name}: start order {printed} of bib {row['Bib']} is printed at position {position}."
            )
        row["Start Order"] = position
    return rows


def write_start_list_sheet(
    workbook: Workbook, sheet_name: str, rows: list[dict[str, Any]]
) -> None:
    worksheet = workbook.create_sheet(sheet_name[:31])
    for column, key in enumerate(START_LIST_COLUMNS, start=1):
        cell = worksheet.cell(row=1, column=column, value=key)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        worksheet.column_dimensions[cell.column_letter].width = 26 if key == "Name" else 11
    for row_number, values in enumerate(rows, start=2):
        for column, key in enumerate(START_LIST_COLUMNS, start=1):
            worksheet.cell(row=row_number, column=column, value=values.get(key))
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
