"""Extract general data, competition information, and footer from FIS PDFs."""

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
    "DG 1",
    "Start Time 1",
    "End Time 1",
    "Wind 1",
    "Base Gate 1",
    "Length 1",
    "Wheater 1",
    "Temp Air 1",
    "Temp Track 1",
    "Humidity 1",
    "DG 2",
    "Start Time 2",
    "End Time 2",
    "Wind 2",
    "Base Gate 2",
    "Length 2",
    "Wheater 2",
    "Temp Air 2",
    "Temp Track 2",
    "Humidity 2",
    "DG 3",
    "Start Time 3",
    "End Time 3",
    "Wind 3",
    "Base Gate 3",
    "Length 3",
    "Wheater 3",
    "Temp Air 3",
    "Temp Track 3",
    "Humidity 3",
    "Footer Datum",
    "Footer Ort",
    "Footer Codex",
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
ROUND_PATTERN = re.compile(
    r"\b(1st|first|2nd|second|3rd|third|trial)\s+round\b",
    re.IGNORECASE,
)


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


def _information_values(text: str) -> dict[str, Any]:
    lines = _header_lines(text)
    section_start = next(
        (
            index
            for index, line in enumerate(lines)
            if re.search(r"\bCompetition\s*/\s*Weather\s+Information\b", line, re.I)
        ),
        None,
    )
    if section_start is None:
        return {}
    section_end = next(
        (
            index
            for index in range(section_start + 1, len(lines))
            if re.match(r"^Statistics\b", lines[index], re.I)
        ),
        len(lines),
    )
    section = " ".join(lines[section_start + 1 : section_end])
    round_matches = list(ROUND_PATTERN.finditer(section))
    values: dict[str, Any] = {}
    round_numbers = {
        "1st": 1,
        "first": 1,
        "2nd": 2,
        "second": 2,
        "3rd": 3,
        "third": 3,
        "trial": 1,
    }
    row_pattern = re.compile(
        rf"(\d{{1,2}}:\d{{2}})\s*[-–—]\s*(\d{{1,2}}:\d{{2}})\s+"
        rf"{NUMBER}\s+{NUMBER}\s+{NUMBER}m?\s+(.+?)\s+"
        rf"{NUMBER}\s+{NUMBER}\s+{NUMBER}\b",
        re.I,
    )
    for index, round_match in enumerate(round_matches):
        round_text = round_match.group(1).casefold()
        round_number = round_numbers[round_text]
        end = round_matches[index + 1].start() if index + 1 < len(round_matches) else len(section)
        row_match = row_pattern.search(section, round_match.end(), end)
        if row_match is None:
            continue
        start_time, end_time, wind, gate, length, weather, air, track, humidity = (
            row_match.groups()
        )
        values.update(
            {
                f"DG {round_number}": round_match.group(0),
                f"Start Time {round_number}": start_time,
                f"End Time {round_number}": end_time,
                f"Wind {round_number}": _number(wind),
                f"Base Gate {round_number}": _number(gate),
                f"Length {round_number}": _number(length),
                f"Wheater {round_number}": " ".join(weather.split()),
                f"Temp Air {round_number}": _number(air),
                f"Temp Track {round_number}": _number(track),
                f"Humidity {round_number}": _number(humidity),
            }
        )
    return values


def _footer_values(text: str) -> dict[str, str]:
    lines = _header_lines(text)
    service_index = next(
        (
            index
            for index in range(len(lines) - 1, -1, -1)
            if re.search(
                r"(?:Timing\s+and\s+)?Data\s+Service\s+by\s+\S+",
                lines[index],
                re.I,
            )
        ),
        None,
    )
    if service_index is None:
        return {}
    footer_line = next(
        (line for line in reversed(lines[:service_index]) if line.strip()),
        "",
    )
    match = re.match(
        r"^\s*(.+?)\s*/\s*(.+?)\s*/\s*(\d+)\s*$",
        footer_line,
    )
    if match is None:
        return {}
    footer_date, place, codex = match.groups()
    return {
        "Footer Datum": footer_date.strip(),
        "Footer Ort": place.strip(),
        "Footer Codex": codex,
    }


def extract_data_from_text(text: str, pdf_name: str = "<PDF>") -> dict[str, Any]:
    """Parse general data and the competition/footer sections from PDF text."""
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
    values.update(_information_values(text))
    values.update(_footer_values(text))

    hill_size_pattern = re.compile(
        rf"\bHILL\b.*?95\s*%\s*\)?\s*{NUMBER}\s*m?\s*"
        rf"\(\s*{NUMBER}\s*m?\s*\)",
        re.I,
    )
    hill_size_match = next(
        (match for line in lines if (match := hill_size_pattern.search(line))),
        None,
    )
    normalized_text = " ".join(text.split())
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
    """Open a PDF once and extract data from its pages."""
    try:
        with pymupdf.open(pdf_path) as document:
            if not document:
                raise PdfParseError(f"{pdf_path.name}: PDF has no pages.")
            text = "\n".join(page.get_text(sort=True) for page in document)
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise PdfParseError(f"{pdf_path.name}: cannot read PDF: {error}") from error
    return extract_data_from_text(text, pdf_path.name)
