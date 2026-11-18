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
    "Pts / Gate DGW",  # from Liste_Schanzen_ID in SJ-Listen.xlsm
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
JUDGE_LINE_PATTERN = re.compile(r"^[A-E]:?\s")  # judge line of the jury list
GERMAN_DATE_PATTERN = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b")  # '15.03.2026'
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
    r"\b(1st|first|2nd|second|3rd|third|final|trial)\s+round\b",
    re.IGNORECASE,
)


class PdfParseError(ValueError):
    """Raised when a PDF does not contain a recognizable general-data header."""


def _number(value: str) -> int | float:
    normalized = value.replace(",", ".")
    parsed = float(normalized)
    return int(parsed) if parsed.is_integer() else parsed


# The title starts with FIS, optionally after up to three sponsor words.
TITLE_PATTERN = re.compile(r"^(?P<sponsor>(?:[A-Z0-9&.'-]+\s+){0,3}?)FIS\b", re.I)
COUNTRY_PATTERN = re.compile(r"\((?!FIS\))[A-Z]{3}\)")  # '(SUI)'; '(FIS)' is no country
OTHER_TITLE_PATTERN =re.compile(r"\b(?:CUP|GAMES|CHAMPIONSHIPS?|TOUR|TOURNAMENT|OPA|FESA)\b", re.I)


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


MONTH_NAMES = {
    name: number
    for number, names in enumerate(
        (
            ("January", "Januar"), ("February", "Februar"), ("March", "März"), ("April",),
            ("May", "Mai"), ("June", "Juni"), ("July", "Juli"), ("August",), ("September",),
            ("October", "Oktober"), ("November",), ("December", "Dezember"),
        ),
        start=1,
    )
    for name in names
}
LONG_DATE_PATTERN = re.compile(
    rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\.?\s*({'|'.join(MONTH_NAMES)})\s+(\d{{4}})\b", re.I
)  # '26th October 2024', '23.Januar 2021'
PRINT_DATE_PATTERN = re.compile(r"\bErstellt am:\s*(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})\b")


def _local_date(lines: list[str]) -> date | None:
    """Date of local organisers (WinSpringen): '15.03.2026' or '26th October
    2024' in the header, otherwise the print date 'Erstellt am: 20.09.25'."""
    head = "\n".join(lines[:15])
    match = GERMAN_DATE_PATTERN.search(head)
    if match:
        day, month, year = (int(part) for part in match.groups())
        return date(year, month, day)
    match = LONG_DATE_PATTERN.search(head)
    if match:
        month = next(number for name, number in MONTH_NAMES.items() if name.casefold() == match.group(2).casefold())
        return date(int(match.group(3)), month, int(match.group(1)))
    match = PRINT_DATE_PATTERN.search("\n".join(lines))
    if match:
        day, month, year = (int(part) for part in match.groups())
        return date(year + 2000 if year < 100 else year, month, day)
    return None


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
    # "Final Round" (e.g. World Cup, Grand Prix) is the round after the last
    # numbered one: 1st Round, Final Round -> 1, 2.
    round_numbers["final"] = 1 + max(
        (
            round_numbers[match.group(1).casefold()]
            for match in round_matches
            if match.group(1).casefold() in round_numbers
        ),
        default=0,
    )
    has_first_round = any(
        round_numbers[match.group(1).casefold()] == 1
        and match.group(1).casefold() != "trial"
        for match in round_matches
    )
    for index, round_match in enumerate(round_matches):
        round_text = round_match.group(1).casefold()
        round_number = round_numbers[round_text]
        # The trial round shares the "1" columns; it is only used when the
        # PDF has no real first round, so it never overwrites competition data.
        if round_text == "trial" and has_first_round:
            continue
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


def local_hill_values(text: str) -> dict[str, Any]:
    """Hill data in the formats of local organisers, e.g. 'Hill Size / (95%)
    106m (100,5m)', 'Hill Size (95%) 105m (99.5m)', 'HILL-SIZE / (95%) 117m
    (111m)', 'Hill Size: 66', 'Metre value 2.0'."""
    number = r"(\d+(?:[.,]\d+)?)"
    # 'Hill Size', 'HILL-SIZE', optionally followed by '/ (95%)' or '(95%)' and ':'
    hill_size = r"Hill[- ]?Size\b(?:\s*/?\s*\(\s*95\s*%\s*\))?\s*:?\s*"
    patterns = {
        "Hill Size": rf"{hill_size}{number}\s*m?\b",
        "HS 95%": rf"{hill_size}\d+\s*m?\s*\(\s*{number}\s*m?\s*\)",
        "K-Punkt": rf"K-?Point\b\s*:?\s*{number}",
        "Meter Value": rf"Met(?:er|re)\s*Value\b\s*:?\s*{number}",
        "Gate Factor": rf"Gate\s*Fa[ck]tor\b\s*:?\s*{number}",
    }
    values: dict[str, Any] = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.I)
        if match:
            values[key] = _number(match.group(1))
    wind = re.search(
        rf"(?:head\s*/\s*tail|Head/Tail)\)?\s*:?\s*{number}\s*/\s*{number}", text, re.I
    )
    if wind:
        values["HWC"], values["TWC"] = _number(wind.group(1)), _number(wind.group(2))
    if "Hill Size" not in values:
        match = re.search(r"\bHS\s*(\d{2,3})\b", text)
        if match:
            values["Hill Size"] = int(match.group(1))
    return values


def extract_data_from_text(text: str, pdf_name: str = "<PDF>") -> dict[str, Any]:
    """Parse general data and the competition/footer sections from PDF text."""
    lines = _header_lines(text)
    if not lines:
        raise PdfParseError(f"{pdf_name}: PDF has no text.")
    title_index = next(
        (index for index, line in enumerate(lines) if TITLE_PATTERN.match(line)),
        None,
    )
    if title_index is not None:
        # Without a sponsor in front: 'VIESSMANN FIS SKI JUMPING WORLD CUP'
        # -> 'FIS SKI JUMPING WORLD CUP'.
        title_match = TITLE_PATTERN.match(lines[title_index])
        title = lines[title_index][title_match.end("sponsor"):].strip()
    else:
        # Alpen Cup (OPA/FESA), FESA games, local organisers: no 'FIS' title.
        title_index = next(
            (index for index, line in enumerate(lines[:12]) if OTHER_TITLE_PATTERN.search(line)),
            0,
        )
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
        # 'Einsiedeln (SUI)'; not '(FIS)' of an official, not the footer
        # '27 JUN 2023 / Zakopane (POL) / 3001'.
        location_index = next(
            (
                index
                for index, line in enumerate(detail_lines)
                if len(COUNTRY_PATTERN.findall(line)) == 1
                and " / " not in line and "(FIS)" not in line
                and not JUDGE_LINE_PATTERN.match(line)  # 'E HLAVA Josef (CZE)'
            ),
            None,
        )
        if location_index is None:
            location_index = 0
        place = detail_lines[location_index]
        if JUDGE_LINE_PATTERN.match(place):
            place = None  # no place in the header, only the jury list
    if place is not None:
        # 'Engelberg (SUI) Start Time Trial 11:00' -> 'Engelberg (SUI)'
        country = COUNTRY_PATTERN.search(place)
        if country:
            place = place[: country.end()]
        if location_index + 1 < len(detail_lines):
            hill_text = detail_lines[location_index + 1]

    if place is None:
        # Local organisers (WinSpringen): 'Hinterzarten (GER)' below the title,
        # the date as '15.03.2026'.
        head = lines[title_index + 1 : title_index + 6]
        location_index = next(
            (
                index for index, line in enumerate(head)
                if len(COUNTRY_PATTERN.findall(line)) == 1
                and " / " not in line and ":" not in line and "(FIS)" not in line
                and len(line.split()) <= 4  # 'Hinterzarten (GER)', not a line of officials
                and not JUDGE_LINE_PATTERN.match(line)
            ),
            None,
        )
        if location_index is not None:
            line = head[location_index]
            place = line[: COUNTRY_PATTERN.search(line).end()]
            following = [l for l in head[location_index + 1 :] if not GERMAN_DATE_PATTERN.search(l)]
            if not hill_text and following:
                hill_text = following[0]
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
    if values["Datum"] is None:
        values["Datum"] = _local_date(lines)
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

    # Formats of local organisers ('Hill Size (95%) 105m (99.5m)', 'Hill Size: 66').
    for key, value in local_hill_values(normalized_text).items():
        if values.get(key) is None:
            values[key] = value
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
