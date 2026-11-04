"""Extract block 1 (general competition data) from a FIS results PDF."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import pymupdf


DATA_ROWS = (
    "Titel",
    "Typ",
    "Datum",
    "Start Time",
    "Finish Time",
    "Competition",
    "Place",
    "Hill",
    "Gender",
    "Rangliste",
    "Hill Size",
    "HS 95%",
    "K-Punkt",
    "Meter Value",
    "Gate Factor",
    "HWC",
    "TWC",
)

DATE_PATTERN = re.compile(
    r"\b(?:MON|TUE|WED|THU|FRI|SAT|SUN)\s+"
    r"(\d{1,2})\s+([A-Z]{3})\s+(\d{4})\b",
    re.IGNORECASE,
)
MONTHS = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}
GENDER_PATTERN = re.compile(r"\b(MEN|WOMEN|MIXED)\b", re.IGNORECASE)
NUMBER = r"(-?\d+(?:[.,]\d+)?)"


class PdfParseError(ValueError):
    """Raised when a PDF does not contain a recognizable general-data header."""


def _number(value: str) -> int | float:
    normalized = value.replace(",", ".")
    parsed = float(normalized)
    return int(parsed) if parsed.is_integer() else parsed


def _competition_type(title: str) -> str:
    result = re.sub(r"\b(?:FIS|SKI|JUMPING|MEN|WOMEN|MIXED)\b", " ", title, flags=re.I)
    result = " ".join(result.split())
    return "FIS CUP" if result.casefold() == "cup" else result


def _header_lines(text: str) -> list[str]:
    return [" ".join(line.split()) for line in text.splitlines() if line.strip()]


def _find_date(text: str) -> date | None:
    match = DATE_PATTERN.search(text)
    if not match:
        return None
    day, month, year = match.groups()
    month_number = MONTHS.get(month.upper())
    if month_number is None:
        return None
    return date(int(year), month_number, int(day))


def _parse_hill_and_gender(hill_text: str, title: str) -> tuple[str | None, str | None]:
    gender_match = GENDER_PATTERN.search(hill_text) or GENDER_PATTERN.search(title)
    if gender_match is None:
        return hill_text or None, None

    gender = gender_match.group(1).upper()
    hill = GENDER_PATTERN.sub(" ", hill_text)
    return " ".join(hill.split()) or None, gender


def extract_data_from_text(text: str, pdf_name: str = "<PDF>") -> dict[str, Any]:
    """Parse the first-page header and Hill Data from extracted PDF text."""
    lines = _header_lines(text)
    title_index = next(
        (index for index, line in enumerate(lines) if re.match(r"^FIS\b", line, re.I)),
        None,
    )
    if title_index is None:
        raise PdfParseError(f"{pdf_name}: could not find the FIS title.")

    title = lines[title_index]
    date_index = next(
        (index for index, line in enumerate(lines) if DATE_PATTERN.search(line)),
        None,
    )
    competition = (
        DATE_PATTERN.sub("", lines[date_index]).strip() or None
        if date_index is not None
        else None
    )

    time_pattern = re.compile(
        r"\b(Start Time|End Time|Finish Time)\s*:?\s*(\d{1,2}:\d{2})\b",
        re.I,
    )
    start_time = None
    finish_time = None
    for line in lines:
        for match in time_pattern.finditer(line):
            label, time_value = match.groups()
            if label.casefold() == "start time":
                start_time = time_value
            else:
                finish_time = time_value

    results_pattern = re.compile(
        r"^(?:RESULTS\b|OFFICIAL RESULTS\b|INTERMEDIATE RESULTS\b)",
        re.I,
    )
    header_end = next(
        (
            index
            for index in range((date_index or title_index) + 1, len(lines))
            if results_pattern.match(lines[index])
            or re.match(r"^Hill Data\b", lines[index], re.I)
        ),
        len(lines),
    )
    detail_lines = []
    if date_index is not None:
        for line in lines[date_index + 1 : header_end]:
            line_without_date = DATE_PATTERN.sub("", line)
            line_without_times = time_pattern.sub("", line_without_date).strip()
            if line_without_times:
                detail_lines.append(line_without_times)

    place = None
    hill_text = ""
    if detail_lines:
        location_index = next(
            (
                index
                for index, line in enumerate(detail_lines)
                if re.search(r"\([A-Z]{3}\)", line)
            ),
            None,
        )
        if location_index is None:
            location_index = 0
        place = detail_lines[location_index]
        if location_index + 1 < len(detail_lines):
            hill_text = detail_lines[location_index + 1]

    hill, gender = _parse_hill_and_gender(hill_text, title)
    rangliste = next(
        (line for line in lines if results_pattern.match(line)),
        None,
    )

    values: dict[str, Any] = {
        "Titel": title,
        "Typ": _competition_type(title),
        "Datum": _find_date(text),
        "Start Time": start_time,
        "Finish Time": finish_time,
        "Competition": competition,
        "Place": place,
        "Hill": hill,
        "Gender": gender,
        "Rangliste": rangliste,
        "Hill Size": None,
        "HS 95%": None,
        "K-Punkt": None,
        "Meter Value": None,
        "Gate Factor": None,
        "HWC": None,
        "TWC": None,
    }

    normalized_text = " ".join(text.split())
    hill_size_match = re.search(
        rf"HILL\s+SIZE\s+HS\s*/\s*\(95%\)\s*{NUMBER}\s*m?\s*"
        rf"\(\s*{NUMBER}\s*m?\s*\)",
        normalized_text,
        re.I,
    )
    if hill_size_match:
        values["Hill Size"] = _number(hill_size_match.group(1))
        values["HS 95%"] = _number(hill_size_match.group(2))

    simple_patterns = {
        "K-Punkt": rf"\bK-POINT\s*{NUMBER}\s*m?\b",
        "Meter Value": rf"\bMETER VALUE\s*{NUMBER}\b",
        "Gate Factor": rf"\bGate Factor\s*{NUMBER}\b",
    }
    for key, pattern in simple_patterns.items():
        match = re.search(pattern, normalized_text, re.I)
        if match:
            values[key] = _number(match.group(1))

    wind_match = re.search(
        rf"\bWind Factors\s*-\s*head\s*/\s*tail\s*{NUMBER}\s*/\s*{NUMBER}",
        normalized_text,
        re.I,
    )
    if wind_match:
        values["HWC"] = _number(wind_match.group(1))
        values["TWC"] = _number(wind_match.group(2))

    return values


def extract_data(pdf_path: Path) -> dict[str, Any]:
    """Read the first page of a PDF once and extract its general-data fields."""
    try:
        with pymupdf.open(pdf_path) as document:
            if not document:
                raise PdfParseError(f"{pdf_path.name}: PDF has no pages.")
            text = document[0].get_text(sort=True)
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise PdfParseError(f"{pdf_path.name}: cannot read PDF: {error}") from error
    return extract_data_from_text(text, pdf_path.name)
