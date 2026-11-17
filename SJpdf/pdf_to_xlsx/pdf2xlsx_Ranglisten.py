"""Combine the result tables of all PDFs into one 'Ranglisten' sheet.

Each round of each PDF (in the order of the Daten sheet) becomes a block of
rows in start order. The columns are fixed (RANGLISTEN_COLUMNS) on every
run, because the sheet is later loaded into a database:

- A:H identity and status, I:AB the round values (without round number),
  AC Total (only in the second competition round), AD JumpID, AE PDF.
- JumpID is the competition name plus the round number and type, e.g.
  '2026-09-13 FCM Einsiedeln HS117 01T', so it is unique in a database.

Competitions (RL): the first round is taken from the intermediate results
after the first round (RL1), because an athlete disqualified in the second
round loses the first-round values in RL. The jump order comes from the
start lists (SLR1, SLRF); without them, round 1 is in bib order and the
second round in the reverse order of RL1 (with a warning).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Any, Sequence

from openpyxl.styles import Alignment, Font
from openpyxl.workbook import Workbook

if __package__:
    from .pdf2xlsx_Resultate import (
        HEADER_FILL,
        NUMBER_FORMATS,
        RESULT_COLUMNS,
        ROUND_FIELDS,
        STATUS_CODES,
    )
else:
    from pdf2xlsx_Resultate import (
        HEADER_FILL,
        NUMBER_FORMATS,
        RESULT_COLUMNS,
        ROUND_FIELDS,
        STATUS_CODES,
    )


SHEET_NAME = "Ranglisten"
IDENTITY_FIELDS = RESULT_COLUMNS[:8]
RANGLISTEN_COLUMNS = (*IDENTITY_FIELDS, *ROUND_FIELDS, "Total", "JumpID", "PDF")
ROUND_TYPES = {"RLT": "T", "RLQ": "Q", "RTRIA": "P", "RL": "W"}
KIND_PATTERN = re.compile(r"(RLT|RLQ|RTRIA|RL)$", re.I)
FIRST_ROUND_REMARK = re.compile(r"\b1st round\b", re.I)
NOT_QUALIFIED = "Not Qualified"

COMPETITION_CODES = (
    (re.compile(r"INTER[- ]?CONTINENTAL CUP"), "ICOC"),
    (re.compile(r"CONTINENTAL CUP"), "COC"),
    (re.compile(r"GRAND PRIX"), "GP"),
    (re.compile(r"WORLD CUP"), "WC"),
    (re.compile(r"ALPEN ?CUP|FESA"), "AC"),
    (re.compile(r"FIS CUP"), "FC"),
)
GENDER_CODES = {"MEN": "M", "WOMEN": "W", "MIXED": "A"}


class RanglistenError(ValueError):
    """Raised when the rounds cannot be combined without doubt."""


@dataclass
class PdfResults:
    stem: str
    rows: list[dict[str, Any]]
    first_round_stem: str | None = None
    first_round_rows: list[dict[str, Any]] | None = None
    # Bibs in start order from the start lists (SLR1, SLRF), if available.
    start_order_1: list[int] | None = None
    start_order_2: list[int] | None = None


def pdf_kind(stem: str) -> str:
    match = KIND_PATTERN.search(stem)
    if match is None:
        raise RanglistenError(f"{stem}: unknown PDF type.")
    return match.group(1).upper()


def _bib_order(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    missing = [row.get("Name") for row in rows if row.get("Bib") is None]
    if missing:
        raise RanglistenError(f"athletes without bib: {missing}")
    return sorted(rows, key=lambda row: row["Bib"])


def _round_status(row: dict[str, Any], number: int) -> str | None:
    """DNS/DSQ/... belongs to the round without values."""
    status = row.get("Status")
    if status in STATUS_CODES and row.get(f"Distance {number}") is not None:
        return None
    return status


def _round_row(row: dict[str, Any], number: int) -> dict[str, Any]:
    values = {key: row.get(key) for key in IDENTITY_FIELDS}
    values["Status"] = _round_status(row, number)
    for name in ROUND_FIELDS:
        values[name] = row.get(f"{name} {number}")
    return values


def _simple_rounds(pdf: PdfResults) -> list[tuple[str, list[dict[str, Any]]]]:
    rows = _bib_order(pdf.rows)
    rounds = [(pdf.stem, [_round_row(row, 1) for row in rows])]
    second = [_round_row(row, 2) for row in rows if row.get("Distance 2") is not None]
    if second:
        rounds.append((pdf.stem, second))
    return rounds


def _competition_rounds(
    pdf: PdfResults, warnings: list[str]
) -> list[tuple[str, list[dict[str, Any]]]]:
    rl_rows = pdf.rows
    jumped = [row for row in rl_rows if row.get("Distance 2") is not None]

    if pdf.first_round_rows is None:
        warnings.append(
            f"{pdf.stem}: {pdf.stem}1.pdf not found; first round taken from {pdf.stem} "
            "(athletes disqualified in the second round have no first-round values)."
        )
        first = [_round_row(row, 1) for row in _bib_order(rl_rows)]
        second_rows = sorted(
            jumped, key=lambda row: (-(row.get("Round Rank 1") or 0), row["Bib"])
        )
        rounds = [(pdf.stem, first)]
        if second_rows:
            rounds.append((pdf.stem, [_second_round_row(row, row) for row in second_rows]))
        return rounds

    rl1_rows = pdf.first_round_rows
    rl1_by_bib = {row["Bib"]: row for row in rl1_rows if row.get("Bib") is not None}
    unknown = [row.get("Name") for row in rl_rows if row.get("Bib") not in rl1_by_bib]
    if unknown:
        raise RanglistenError(f"{pdf.stem}: athletes missing in {pdf.first_round_stem}: {unknown}")

    # Athletes who started in round 2: with values, or DSQ/DNS/NPS in round 2.
    second_round_status: dict[int, dict[str, Any]] = {}
    first_round_status: dict[int, str] = {}
    for row in rl_rows:
        status = row.get("Status")
        if status not in STATUS_CODES or row.get("Distance 2") is not None:
            continue
        rl1_row = rl1_by_bib[row["Bib"]]
        if FIRST_ROUND_REMARK.search(row.get("Remarks") or "") or rl1_row.get("Distance 1") is None:
            first_round_status[row["Bib"]] = status
        else:
            second_round_status[row["Bib"]] = row
    starters = {row["Bib"] for row in jumped} | set(second_round_status)

    first = []
    for rl1_row in _start_order(pdf, rl1_rows, warnings):
        values = _round_row(rl1_row, 1)
        if values["Round Total"] is None:
            values["Round Total"] = rl1_row.get("Total")
        if values["Round Rank"] is None:
            values["Round Rank"] = rl1_row.get("Rank")
        bib = rl1_row["Bib"]
        if values["Status"] is None:
            if bib in first_round_status:
                values["Status"] = first_round_status[bib]
            elif bib not in starters and rl1_row.get("Distance 1") is not None:
                values["Status"] = NOT_QUALIFIED
        first.append(values)

    not_qualified_rl = {row["Bib"] for row in rl_rows if row.get("Status") == NOT_QUALIFIED}
    not_qualified = {row["Bib"] for row in first if row["Status"] == NOT_QUALIFIED}
    if not_qualified_rl != not_qualified:
        raise RanglistenError(
            f"{pdf.stem}: 'Not Qualified' differs between {pdf.stem} and "
            f"{pdf.first_round_stem} (bibs {sorted(not_qualified_rl ^ not_qualified)})."
        )

    if pdf.start_order_2 is not None:
        order = [bib for bib in pdf.start_order_2 if bib in starters]
        # A starter missing in SLRF is inserted at the place given by the
        # reverse RL1 ranking: before the first athlete ranked better in RL1.
        rank = {row["Bib"]: row.get("Rank") for row in rl1_rows}
        for bib in sorted(starters - set(order), key=lambda bib: -(rank.get(bib) or 0)):
            own = rank.get(bib)
            index = next(
                (
                    index for index, other in enumerate(order)
                    if own is not None and rank.get(other) is not None and rank[other] < own
                ),
                len(order),
            )
            order.insert(index, bib)
            warnings.append(
                f"{pdf.stem}: {_name(rl1_by_bib[bib])} (bib {bib}) is missing in the "
                f"final-round start list (SLRF); second-round position estimated from "
                f"the reverse ranking of {pdf.first_round_stem} (start {index + 1})."
            )
        missing_text = "second-round starters missing in the final-round start list (SLRF)"
    else:
        # Without SLRF: reverse order of the first-round ranking (RL1).
        warnings.append(
            f"{pdf.stem}: final-round start list (SLRF) not found; second round "
            f"in the reverse order of {pdf.first_round_stem}."
        )
        ranking = [row for row in rl1_rows if row.get("Rank") is not None]
        order = [row["Bib"] for row in reversed(ranking) if row["Bib"] in starters]
        missing_text = "second-round starters without first-round rank"
    if set(order) != starters:
        raise RanglistenError(f"{pdf.stem}: {missing_text}: {sorted(starters - set(order))}")
    rl_by_bib = {row["Bib"]: row for row in rl_rows}
    second = []
    for bib in order:
        if bib in second_round_status:
            values = {key: None for key in RANGLISTEN_COLUMNS}
            for key in IDENTITY_FIELDS:
                if key != "Rank":
                    values[key] = rl_by_bib[bib].get(key) or rl1_by_bib[bib].get(key)
            second.append(values)
        else:
            second.append(_second_round_row(rl_by_bib[bib], rl1_by_bib[bib]))

    rounds = [(pdf.first_round_stem or pdf.stem, first)]
    if second:
        rounds.append((pdf.stem, second))
    return rounds


def _start_order(
    pdf: PdfResults, rows: list[dict[str, Any]], warnings: list[str]
) -> list[dict[str, Any]]:
    """First-round rows in the order of the start list SLR1 (else bib order)."""
    if pdf.start_order_1 is None:
        warnings.append(
            f"{pdf.stem}: first-round start list (SLR1) not found; first round in bib order."
        )
        return _bib_order(rows)
    if any(row.get("Bib") is None for row in rows):
        _bib_order(rows)  # raises with the names
    order = [bib for bib in pdf.start_order_1 if bib in {row["Bib"] for row in rows}]
    # An athlete missing in SLR1 is inserted at his bib position: before the
    # first athlete with a higher bib (lower bib in a descending start list).
    descending = len(pdf.start_order_1) > 1 and pdf.start_order_1[0] > pdf.start_order_1[-1]
    for row in _bib_order([row for row in rows if row["Bib"] not in set(order)]):
        bib = row["Bib"]
        index = next(
            (
                index for index, other in enumerate(order)
                if (other < bib if descending else other > bib)
            ),
            len(order),
        )
        order.insert(index, bib)
        warnings.append(
            f"{pdf.stem}: {_name(row)} (bib {bib}) is missing in the first-round start "
            f"list (SLR1); first-round position estimated from the bib (start {index + 1})."
        )
    position = {bib: index for index, bib in enumerate(order)}
    return sorted(rows, key=lambda row: position[row["Bib"]])


def _name(row: dict[str, Any]) -> str:
    return str(row.get("Name") or "athlete")


def _second_round_row(rl_row: dict[str, Any], rl1_row: dict[str, Any]) -> dict[str, Any]:
    values = _round_row(rl_row, 2)
    for key in IDENTITY_FIELDS:
        if values.get(key) is None and key not in ("Rank", "Status"):
            values[key] = rl1_row.get(key)
    values["Total"] = rl_row.get("Total")
    return values


def build_ranglisten(
    pdfs: Sequence[PdfResults], competition: str
) -> tuple[list[dict[str, Any]], list[str]]:
    """Return the Ranglisten rows (in jump order) and warnings.

    competition is the competition name (competition_file_name), the first
    part of every JumpID.
    """
    warnings: list[str] = []
    rows: list[dict[str, Any]] = []
    jump_number = 0
    for pdf in pdfs:
        kind = pdf_kind(pdf.stem)
        rounds = (
            _competition_rounds(pdf, warnings) if kind == "RL" else _simple_rounds(pdf)
        )
        for source, round_rows in rounds:
            jump_number += 1
            for values in round_rows:
                row = {key: values.get(key) for key in RANGLISTEN_COLUMNS}
                row["JumpID"] = f"{competition} {jump_number:02d}{ROUND_TYPES[kind]}"
                row["PDF"] = source
                rows.append(row)
    return rows, warnings


def write_ranglisten_sheet(workbook: Workbook, rows: list[dict[str, Any]]) -> None:
    worksheet = workbook.create_sheet(SHEET_NAME, 1)
    for column, key in enumerate(RANGLISTEN_COLUMNS, start=1):
        cell = worksheet.cell(row=1, column=column, value=key)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        worksheet.column_dimensions[cell.column_letter].width = {
            "Name": 26, "Club": 30, "Date of Birth": 12, "Status": 13, "JumpID": 36, "PDF": 16,
        }.get(key, 9)
    for row_number, values in enumerate(rows, start=2):
        for column, key in enumerate(RANGLISTEN_COLUMNS, start=1):
            value = values.get(key)
            if value is None:
                continue
            cell = worksheet.cell(row=row_number, column=column, value=value)
            number_format = NUMBER_FORMATS.get(key)
            if number_format:
                cell.number_format = number_format
    worksheet.freeze_panes = "D2"
    worksheet.auto_filter.ref = worksheet.dimensions


def competition_file_name(values: Sequence[dict[str, Any]]) -> str:
    """Return e.g. '2026-09-13 FCM Einsiedeln HS117' from the Daten values."""
    if not values:
        raise RanglistenError("no PDFs for the file name.")
    dates = [value.get("Datum") for value in values]
    if any(not isinstance(day, date) for day in dates):
        raise RanglistenError("date missing in a PDF.")

    def single(key: str) -> Any:
        found = {value.get(key) for value in values}
        if len(found) != 1 or None in found:
            raise RanglistenError(f"'{key}' is not the same in all PDFs: {sorted(map(str, found))}")
        return found.pop()

    title = " ".join(str(single("Titel")).upper().split())
    code = next((code for pattern, code in COMPETITION_CODES if pattern.search(title)), None)
    if code is None:
        raise RanglistenError(f"unknown competition type '{title}'.")
    gender = GENDER_CODES.get(str(single("Gender")).upper())
    if gender is None:
        raise RanglistenError(f"unknown gender '{single('Gender')}'.")
    place = re.sub(r"\s*\([A-Z]{3}\)\s*$", "", str(single("Place"))).strip()
    hill_size = single("Hill Size")
    hill_size = int(hill_size) if float(hill_size).is_integer() else hill_size
    name = f"{max(dates).isoformat()} {code}{gender} {place} HS{hill_size}"
    if re.search(r'[<>:"/\\|?*]', name):
        raise RanglistenError(f"invalid file name '{name}'.")
    return name
