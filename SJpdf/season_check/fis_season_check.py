"""Check a whole FIS season against SJ-Listen.xlsm.

For every competition of the season (FIS calendar, as read by the race-ID
crawler) the script reads

- the FIS results page: FIS code, name, birth year and NSA of each athlete,
- one start list PDF (1st round, qualification or training; else a results
  PDF): date of birth of the athletes and the hill data (place, HS, ...).

Athletes missing in Liste_Springer_ID and hills missing in Liste_Schanzen_ID
are written to Fehlende_JumperID.xlsx in the SJpdf data folder, in the
column order of the lists, so the rows can be copied into SJ-Listen.xlsm.
A date of birth missing in the PDFs (and the gender of athletes who only
jumped in mixed teams) is taken from the FIS athlete page.

Downloads are cached outside the synced data folder, so a second run is fast.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import os
import re
import sys
import tempfile
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
from fis_web import (  # noqa: E402
    REQUEST_DELAY_SECONDS,
    calendar_events,
    calendar_url,
    detail_url,
    open_url,
    parse_document,
    read_text,
)
from pdf_to_xlsx.pdf2xlsx_Daten import PdfParseError, extract_data_from_text  # noqa: E402
from pdf_to_xlsx.pdf2xlsx_Listen import (  # noqa: E402
    SJ_LISTEN_FILE,
    SjListen,
    _key,
    hill_name,
    proposed_jumper_id,
    read_sj_listen,
)
from raceid_crawler.fis_raceid_crawler import _competition_rows  # noqa: E402
from sj_paths import DATA_DIRECTORY  # noqa: E402

_spec = importlib.util.spec_from_file_location(
    "fis_pdf_download", PROJECT_ROOT / "pdf_download" / "FIS_pdf-download.py"
)
_downloader = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_downloader)


OUTPUT_FILE = DATA_DIRECTORY / "Fehlende_JumperID.xlsx"
CACHE_DIRECTORY = (
    Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "SJpdf" / "cache"
)
RESULTS_URL = "https://www.fis-ski.com/DB/general/results.html?sectorcode=JP&raceid={raceid}"
BIOGRAPHY_URL = (
    "https://www.fis-ski.com/DB/general/athlete-biography.html"
    "?sectorcode=jp&competitorid={competitorid}"
)
# Start lists with date of birth first; the final-round list has none.
PDF_PRIORITY = (
    "start list 1st round",
    "start list qualification",
    "start list training",
    "start list trial",
    "start list",
    "official results",
    "qualification results",
    "results 1st round",
    "results training",
    "results trial",
)
MONTHS = {
    name: number
    for number, name in enumerate(
        ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"),
        start=1,
    )
}
DOB_PATTERN = re.compile(r"\b(\d{1,2}) (" + "|".join(MONTHS) + r") ((?:19|20)\d{2})\b")
ATHLETE_LINE = re.compile(r"^\s*(?:\d{1,3}\s+){1,2}(\S.*?\S)\s{2,}([A-Z]{3})(?=\s|$)")
FIS_GENDER = {"M": "M", "W": "F"}  # Liste_Springer_ID uses M / F
YEAR_ONLY = "nur Jahrgang (1.1.)"
JUMPER_HEADERS = ("JumperID", "Name", "NSA", "Gender", "DoB", "FIS_Code", "JG")
HILL_HEADERS = (
    "SchanzenID", "HS", "95%", "K", "Pts./m", "Headwind \nPts./m/s",
    "Tailwind \nPts./m/s", "Gate \nFactor \nPts./m", "Pts / Gate \nfür DGW",
    "UTC Sommer", "UTC Winter",
)
HEADER_FILL = PatternFill(fill_type="solid", fgColor="FF17358A")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
MARK_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
YEAR_FILL = PatternFill(fill_type="solid", fgColor="FFFFEB9C")  # light yellow


# --- downloads ---------------------------------------------------------------

class Fetcher:
    """Download with cache and the FIS request delay (only for real requests)."""

    def __init__(self, cache_directory: Path = CACHE_DIRECTORY, log: Callable[[str], None] = print):
        self.cache_directory = cache_directory
        self.log = log
        self._last_request = 0.0

    def get(self, url: str) -> bytes:
        cache_file = self.cache_directory / hashlib.sha1(url.encode()).hexdigest()
        if cache_file.is_file():
            return cache_file.read_bytes()
        wait = REQUEST_DELAY_SECONDS - (time.monotonic() - self._last_request)
        if wait > 0:
            time.sleep(wait)

        def report_retry(delay: int) -> None:
            self.log(f"FIS-Drosselung (HTTP 429); neuer Versuch in {delay} Sekunden.")

        with open_url(url, timeout=60, on_retry=report_retry) as response:
            content = response.read()
        self._last_request = time.monotonic()
        self.cache_directory.mkdir(parents=True, exist_ok=True)
        cache_file.write_bytes(content)
        return content

    def text(self, url: str) -> str:
        return self.get(url).decode("utf-8", errors="replace")


# --- parsing -----------------------------------------------------------------

@dataclass
class Athlete:
    fis_code: str
    name: str
    nsa: str
    birth_year: int | None
    competitorid: str | None
    genders: set[str] = field(default_factory=set)
    dob: date | None = None
    dob_source: str = ""
    club: str = ""
    first_competition: str = ""


@dataclass
class Hill:
    hill_id: str
    values: dict[str, Any]
    competitions: list[str] = field(default_factory=list)


def results_athletes(html: str) -> list[dict[str, Any]]:
    """Athletes of a FIS results page.

    Team competitions list a team row (e.g. 'JAPAN', no birth year) with its
    athletes below; their nation is empty and taken from the team row.
    """
    athletes = []
    team_nation = ""
    document = parse_document(html)
    for row in document.find_all(lambda item: item.tag == "a" and "table-row" in item.classes):
        match = re.search(r"competitorid=(\d+)", row.attrs.get("href", ""))
        line = row.find_all(lambda item: item.tag == "div" and "justify-sb" in item.classes)
        if not match or not line:
            continue
        cells = [child for child in line[0].children if hasattr(child, "tag") and child.tag == "div"]
        texts = [cell.text for cell in cells]
        # rank, bib, FIS code, name, birth year, nation, ...
        if len(texts) < 6 or not texts[2].isdigit():
            continue
        nation = texts[5].split()[0] if texts[5].split() else ""
        if not texts[4].strip() and texts[3] == texts[3].upper():
            team_nation = nation  # team row
            continue
        nation = nation or team_nation
        athletes.append(
            {
                "fis_code": texts[2],
                "name": texts[3],
                "birth_year": int(texts[4]) if texts[4].isdigit() else None,
                "nsa": nation,
                "competitorid": match.group(1),
            }
        )
    return athletes


def _dob(text: str) -> date | None:
    match = DOB_PATTERN.search(text)
    if not match:
        return None
    day, month, year = match.groups()
    try:
        return date(int(year), MONTHS[month], int(day))
    except ValueError:
        return None


def pdf_people(text: str) -> dict[tuple[str, str], tuple[date, str]]:
    """{(name, NSA): (date of birth, club)} of a FIS start list or results PDF.

    The athlete line (bib, NAME First, NSA) is followed by a line with club
    and date of birth.
    """
    people = {}
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = ATHLETE_LINE.match(line)
        if not match or DOB_PATTERN.search(line):
            continue
        for following in lines[index + 1 : index + 3]:
            if ATHLETE_LINE.match(following):
                break
            dob_match = DOB_PATTERN.search(following)
            if dob_match:
                club = following[: dob_match.start()].strip()
                people[(_key(match.group(1)), _key(match.group(2)))] = (_dob(following), club)
                break
    return people


def _number(text: str) -> float | int:
    value = float(text.replace(",", "."))
    return int(value) if value.is_integer() else value


def _local_hill_values(text: str) -> dict[str, Any]:
    """Hill data in the formats of local organisers, e.g. 'Hill Size / (95%)
    106m (100,5m)', 'HILL-SIZE / (95%) 117m (111m)', 'Hill Size: 66'."""
    number = r"(\d+(?:[.,]\d+)?)"
    # 'Hill Size', 'HILL-SIZE', optionally followed by '/ (95%)' and ':'
    hill_size = r"Hill[- ]?Size\b(?:\s*/\s*\(\s*95\s*%\s*\))?\s*:?\s*"
    patterns = {
        "Hill Size": rf"{hill_size}{number}\s*m?\b",
        "HS 95%": rf"{hill_size}\d+\s*m?\s*\(\s*{number}\s*m?\s*\)",
        "K-Punkt": rf"K-?Point\b\s*:?\s*{number}",
        "Meter Value": rf"Meter\s*Value\b\s*:?\s*{number}",
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


def pdf_hill(text: str, place: str) -> dict[str, Any] | None:
    """Hill data of a PDF (FIS format, else 'Hill Size: 66' of local organisers)."""
    try:
        values = extract_data_from_text(text, "")
    except PdfParseError:
        values = {}  # no FIS title: PDF of a local organiser
    if values.get("Hill Size") is None:
        values.update(_local_hill_values(text))
        if values.get("Hill Size") is None:
            return None
    # The place of the PDF ('Einsiedeln (SUI)') when it has a country code,
    # else the place of the FIS calendar.
    pdf_place = values.get("Place") or ""
    hill_id = hill_name(pdf_place if re.search(r"\([A-Z]{3}\)", pdf_place) else place, values["Hill Size"])
    return {"SchanzenID": hill_id, **values} if hill_id else None


def biography(html: str) -> tuple[date | None, str | None]:
    """Date of birth and gender (M/F) of a FIS athlete page."""
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", html))
    dob = None
    match = re.search(r"Birthdate[ |]+(\d{2})-(\d{2})-(\d{4})", text)
    if match:
        day, month, year = (int(part) for part in match.groups())
        try:
            dob = date(year, month, day)
        except ValueError:
            dob = None
    gender = None
    match = re.search(r"Gender[ |]+(Male|Female)", text)
    if match:
        gender = "M" if match.group(1) == "Male" else "F"
    return dob, gender


def choose_pdf(attachments: Iterable[Any]):
    labelled = [(" ".join(attachment.label.split()).casefold(), attachment) for attachment in attachments]
    for wanted in PDF_PRIORITY:
        for label, attachment in labelled:
            if label.startswith(wanted) and "final round" not in label:
                return attachment
    return None


def pdf_text(content: bytes) -> str:
    with pymupdf.open(stream=content, filetype="pdf") as document:
        return "\n".join(page.get_text(sort=True) for page in document)


# --- season ------------------------------------------------------------------

@dataclass
class SeasonData:
    athletes: dict[str, Athlete] = field(default_factory=dict)  # by FIS code
    hills: dict[str, Hill] = field(default_factory=dict)  # by key of SchanzenID
    competitions: int = 0
    problems: list[str] = field(default_factory=list)


def collect_season(season: int, fetcher: Fetcher, log: Callable[[str], None] = print) -> SeasonData:
    data = SeasonData()
    events = calendar_events(fetcher.text(calendar_url(season)))
    for number, (eventid, place) in enumerate(events, start=1):
        log(f"[{number}/{len(events)}] Event {eventid}: {place}")
        html = fetcher.text(detail_url(eventid, season))
        records = _competition_rows(html, eventid, place, season, warn=data.problems.append)
        attachments = {competition.codex: competition.attachments for competition in _downloader.competition_rows(html)}
        for record in records:
            data.competitions += 1
            label = f"{record['date']} {record['category']}{record['gender']} {place} {record['codex']}"
            people: dict = {}
            attachment = choose_pdf(attachments.get(record["codex"], ()))
            if attachment is None:
                data.problems.append(f"{label}: kein Startlisten- oder Resultat-PDF.")
            else:
                try:
                    text = pdf_text(fetcher.get(attachment.url))
                except Exception as error:  # noqa: BLE001 - report and go on
                    data.problems.append(f"{label}: PDF nicht lesbar ({error}).")
                    text = ""
                people = pdf_people(text)
                hill = pdf_hill(text, place) if text else None
                if hill is None and text:
                    data.problems.append(f"{label}: Schanze (HS) nicht im PDF gefunden.")
                elif hill is not None:
                    entry = data.hills.setdefault(_key(hill["SchanzenID"]), Hill(hill["SchanzenID"], hill))
                    entry.competitions.append(label)

            found_athletes = results_athletes(fetcher.text(RESULTS_URL.format(raceid=record["raceid"])))
            if not found_athletes:
                data.problems.append(
                    f"{label}: keine Athleten auf der FIS-Resultatseite (z. B. Teamwettkampf "
                    "ohne Einzelnamen oder Resultat fehlt)."
                )
            for found in found_athletes:
                athlete = data.athletes.get(found["fis_code"])
                if athlete is None:
                    athlete = Athlete(
                        found["fis_code"], found["name"], found["nsa"], found["birth_year"],
                        found["competitorid"], first_competition=label,
                    )
                    data.athletes[found["fis_code"]] = athlete
                if record["gender"] in FIS_GENDER:
                    athlete.genders.add(FIS_GENDER[record["gender"]])
                person = people.get((_key(found["name"]), _key(found["nsa"])))
                if person and athlete.dob is None and person[0] is not None:
                    athlete.dob, athlete.club, athlete.dob_source = person[0], person[1], f"PDF {label}"
    return data


@dataclass
class Comparison:
    missing: list[Athlete]
    check: list[tuple[Athlete, Any]]  # FIS code listed with other name or NSA
    missing_hills: list[Hill]


def compare(data: SeasonData, listen: SjListen) -> Comparison:
    missing, check = [], []
    for athlete in data.athletes.values():
        listed = listen.by_fis_code(athlete.fis_code)
        if listed is not None:
            if (_key(listed.name), _key(listed.nsa)) != (_key(athlete.name), _key(athlete.nsa)):
                check.append((athlete, listed))
        elif listen.jumper_id(athlete.name, athlete.nsa) is None:
            missing.append(athlete)
    missing_hills = [hill for key, hill in data.hills.items() if key not in listen.hills]
    return Comparison(missing, check, missing_hills)


def complete_from_biography(athletes: Iterable[Athlete], fetcher: Fetcher, log=print) -> None:
    """Date of birth and gender from the FIS athlete page where still missing."""
    for athlete in athletes:
        if (athlete.dob is None or len(athlete.genders) != 1) and athlete.competitorid:
            log(f"FIS-Athletenseite: {athlete.name} ({athlete.nsa})")
            dob, gender = biography(fetcher.text(BIOGRAPHY_URL.format(competitorid=athlete.competitorid)))
            if athlete.dob is None and dob is not None:
                athlete.dob, athlete.dob_source = dob, "FIS-Athletenseite"
            if len(athlete.genders) != 1 and gender:
                athlete.genders = {gender}
        if athlete.dob is None and athlete.birth_year:
            # Only the birth year is known: 1 January, as in Liste_Springer_ID.
            athlete.dob, athlete.dob_source = date(athlete.birth_year, 1, 1), YEAR_ONLY


# --- workbook ----------------------------------------------------------------

def _header(sheet, headers: Iterable[str], column: int = 1) -> None:
    for offset, header in enumerate(headers):
        cell = sheet.cell(row=1, column=column + offset, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT


def write_workbook(result: Comparison, listen: SjListen, season: int, data: SeasonData, path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Fehlende_JumperID"
    _header(sheet, JUMPER_HEADERS)
    info = ("Club", "Quelle DoB", "Erster Wettkampf", "Hinweis")
    _header(sheet, info, column=len(JUMPER_HEADERS) + 2)
    used_ids = {_key(jumper.jumper_id) for jumper in listen.jumpers}
    rows = []
    for athlete in result.missing:
        jumper_id = proposed_jumper_id(athlete.name, athlete.nsa)
        rows.append((jumper_id, athlete))
    rows.sort(key=lambda item: item[0])
    proposed_count: dict[str, int] = {}
    for jumper_id, _athlete in rows:
        proposed_count[_key(jumper_id)] = proposed_count.get(_key(jumper_id), 0) + 1
    for row_number, (jumper_id, athlete) in enumerate(rows, start=2):
        gender = next(iter(athlete.genders)) if len(athlete.genders) == 1 else None
        year = athlete.dob.year if athlete.dob else athlete.birth_year
        values = (
            jumper_id, athlete.name, athlete.nsa, gender,
            athlete.dob, int(athlete.fis_code), year,
        )
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_number, column=column, value=value)
            if column == 5 and value is not None:
                cell.number_format = "yyyy-mm-dd"
        notes = []
        if _key(jumper_id) in used_ids or proposed_count[_key(jumper_id)] > 1:
            notes.append("JumperID schon vergeben, bitte anpassen")
            sheet.cell(row=row_number, column=1).fill = MARK_FILL
        if gender is None:
            notes.append("Gender unbekannt")
            sheet.cell(row=row_number, column=4).fill = MARK_FILL
        if athlete.dob is None:
            notes.append("Geburtsdatum unbekannt")
            sheet.cell(row=row_number, column=5).fill = MARK_FILL
        elif athlete.dob_source == YEAR_ONLY:
            notes.append("Geburtsdatum: nur Jahrgang bekannt")
            sheet.cell(row=row_number, column=5).fill = YEAR_FILL
        extra = (athlete.club, athlete.dob_source, athlete.first_competition, "; ".join(notes))
        for offset, value in enumerate(extra):
            sheet.cell(row=row_number, column=len(JUMPER_HEADERS) + 2 + offset, value=value or None)
    widths = {"A": 18, "B": 30, "C": 6, "D": 8, "E": 12, "F": 10, "G": 6, "I": 30, "J": 40, "K": 40, "L": 40}
    for letter, width in widths.items():
        sheet.column_dimensions[letter].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(JUMPER_HEADERS) + 1 + len(info)).column_letter}{max(len(rows) + 1, 1)}"

    check = workbook.create_sheet("Prüfen")
    _header(check, ("FIS_Code", "Name FIS", "NSA FIS", "JumperID Liste", "Name Liste", "NSA Liste", "Erster Wettkampf"))
    for row_number, (athlete, listed) in enumerate(sorted(result.check, key=lambda item: item[1].jumper_id), start=2):
        for column, value in enumerate(
            (int(athlete.fis_code), athlete.name, athlete.nsa, listed.jumper_id, listed.name, listed.nsa, athlete.first_competition),
            start=1,
        ):
            check.cell(row=row_number, column=column, value=value)
    for letter, width in {"A": 10, "B": 30, "C": 8, "D": 18, "E": 30, "F": 9, "G": 40}.items():
        check.column_dimensions[letter].width = width

    hills = workbook.create_sheet("Fehlende_SchanzenID")
    _header(hills, HILL_HEADERS)
    _header(hills, ("Wettkämpfe",), column=len(HILL_HEADERS) + 2)
    keys = ("SchanzenID", "Hill Size", "HS 95%", "K-Punkt", "Meter Value", "HWC", "TWC", "Gate Factor")
    for row_number, hill in enumerate(sorted(result.missing_hills, key=lambda hill: hill.hill_id), start=2):
        for column, key in enumerate(keys, start=1):
            hills.cell(row=row_number, column=column, value=hill.values.get(key))
        hills.cell(row=row_number, column=len(HILL_HEADERS) + 2, value="; ".join(hill.competitions))
    hills.column_dimensions["A"].width = 30
    hills.column_dimensions["M"].width = 80

    summary = workbook.create_sheet("Übersicht", 0)
    lines = [
        ("Saison", season),
        ("Erstellt", datetime.now().replace(microsecond=0)),
        ("Wettkämpfe", data.competitions),
        ("Springer in der Saison", len(data.athletes)),
        ("Fehlende JumperID", len(result.missing)),
        ("FIS-Code mit anderem Namen/NSA (Prüfen)", len(result.check)),
        ("Schanzen in der Saison", len(data.hills)),
        ("Fehlende SchanzenID", len(result.missing_hills)),
        ("", ""),
        ("Einfügen", "Fehlende_JumperID A:G in Liste_Springer_ID ab Spalte A einfügen "
                     "(Age berechnet die Tabelle); Fehlende_SchanzenID A:H in Liste_Schanzen_ID, "
                     "Pts / Gate für DGW und UTC ergänzen."),
    ]
    for row_number, (label, value) in enumerate(lines, start=1):
        summary.cell(row=row_number, column=1, value=label).font = Font(bold=True)
        cell = summary.cell(row=row_number, column=2, value=value)
        if isinstance(value, datetime):
            cell.number_format = "yyyy-mm-dd hh:mm:ss"
    row_number = len(lines) + 2
    summary.cell(row=row_number, column=1, value="Meldungen").font = Font(bold=True)
    for offset, problem in enumerate(data.problems or ["keine"], start=1):
        summary.cell(row=row_number + offset, column=2, value=problem)
    summary.column_dimensions["A"].width = 40
    summary.column_dimensions["B"].width = 110
    workbook.active = 1
    workbook.save(path)
    workbook.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Fehlende JumperID und SchanzenID einer FIS-Saison.")
    parser.add_argument("--season", type=int, required=True, help="FIS-Saison, z. B. 2026 (= 2025/26)")
    parser.add_argument("--output", type=Path, default=OUTPUT_FILE)
    parser.add_argument("--listen", type=Path, default=SJ_LISTEN_FILE)
    args = parser.parse_args(argv)

    listen = read_sj_listen(args.listen)
    fetcher = Fetcher()
    data = collect_season(args.season, fetcher)
    result = compare(data, listen)
    complete_from_biography(result.missing, fetcher)
    write_workbook(result, listen, args.season, data, args.output)
    print(
        f"{len(result.missing)} fehlende JumperID, {len(result.check)} zu prüfen, "
        f"{len(result.missing_hills)} fehlende SchanzenID: {args.output}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
