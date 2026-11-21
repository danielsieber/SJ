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
import csv
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
from pdf_to_xlsx.pdf2xlsx_Daten import (  # noqa: E402
    PdfParseError,
    extract_data_from_text,
    local_hill_values,
)
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
# PDF archive of pdf_archive/fis_pdf_archive.py: PDFs found there are not downloaded again.
ARCHIVE_DIRECTORY = Path(r"C:\downloaded_FIS_pdf")
MAX_SKIPPED_HILL_SIZE = 30  # competitions on hills up to HS30 are not evaluated
ARCHIVE_LINKS_FILENAME = "pdf_links.csv"
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
JUMPER_HEADERS = ("JumperID_pdf", "JumperID", "Name", "NSA", "Gender", "DoB", "FIS_Code", "JG", "Age")
HILL_HEADERS = (
    "SchanzenID_pdf", "SchanzenID", "HS", "95%", "K", "Pts./m", "Headwind \nPts./m/s",
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

    def __init__(
        self,
        cache_directory: Path = CACHE_DIRECTORY,
        log: Callable[[str], None] = print,
        use_cache: bool = True,
        local_files: dict[str, Path] | None = None,
    ):
        self.cache_directory = cache_directory
        self.log = log
        self.use_cache = use_cache  # False: always read fresh (pages that still change)
        self.local_files = local_files or {}  # FIS link -> PDF already downloaded
        self._last_request = 0.0

    def get(self, url: str) -> bytes:
        local_file = self.local_files.get(url)
        if local_file is not None and local_file.is_file():
            return local_file.read_bytes()
        cache_file = self.cache_directory / hashlib.sha1(url.encode()).hexdigest()
        if self.use_cache and cache_file.is_file():
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


# Moved to pdf2xlsx_Daten (also used by pdf2xlsx); old name for the tests.
_local_hill_values = local_hill_values


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
    # by FIS code, name and NSA: each spelling and nation of a jumper is checked
    athletes: dict[tuple[str, str, str], Athlete] = field(default_factory=dict)
    hills: dict[str, Hill] = field(default_factory=dict)  # by key of SchanzenID
    competitions: int = 0
    problems: list[str] = field(default_factory=list)


def archive_links(archive_directory: Path | None) -> dict[str, Path]:
    """{FIS link: PDF} of all event folders of the PDF archive (pdf_links.csv)."""
    links: dict[str, Path] = {}
    if archive_directory is None or not archive_directory.is_dir():
        return links
    for links_file in archive_directory.glob(f"*/{ARCHIVE_LINKS_FILENAME}"):
        with links_file.open(encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle):
                if row.get("url") and row.get("pdf"):
                    links[row["url"]] = links_file.parent / row["pdf"]
    return links


def parse_seasons(values: Iterable[str]) -> list[int]:
    """'2026', '2024-2026' or several values -> sorted seasons."""
    seasons: set[int] = set()
    for value in values:
        for part in str(value).replace(",", " ").split():
            first, _, last = part.partition("-")
            seasons.update(range(int(first), int(last or first) + 1))
    return sorted(seasons)


def season_label(seasons: Iterable[int]) -> str:
    seasons = sorted(seasons)
    return str(seasons[0]) if len(seasons) == 1 else f"{seasons[0]}-{seasons[-1]}"


def collect_season(
    season: int, fetcher: Fetcher, log: Callable[[str], None] = print, data: SeasonData | None = None
) -> SeasonData:
    """Athletes and hills of a season (added to data when given)."""
    data = data if data is not None else SeasonData()
    events = calendar_events(fetcher.text(calendar_url(season)))
    for number, (eventid, place) in enumerate(events, start=1):
        log(f"Saison {season} [{number}/{len(events)}] Event {eventid}: {place}")
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
                if hill is not None and (hill.get("Hill Size") or 999) <= MAX_SKIPPED_HILL_SIZE:
                    data.competitions -= 1  # small hills (up to HS30): not evaluated
                    continue
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
                athlete = add_athlete(data, found, label)
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


def add_athlete(data: SeasonData, found: dict[str, Any], label: str) -> Athlete:
    """The athlete of a FIS results page, once per FIS code, name and NSA.

    A name change or a change of nation ('BUECHEL Alina' AUT -> LIE) is a new
    entry with its own first competition, so a later spelling is checked
    against SJ-Listen too, not only the one of the first competition.
    """
    key = (found["fis_code"], _key(found["name"]), _key(found["nsa"]))
    athlete = data.athletes.get(key)
    if athlete is None:
        athlete = Athlete(
            found["fis_code"], found["name"], found["nsa"], found["birth_year"],
            found["competitorid"], first_competition=label,
        )
        data.athletes[key] = athlete
    return athlete


def compare(data: SeasonData, listen: SjListen) -> Comparison:
    """An athlete is listed when name and NSA are in Liste_Springer_ID, or when
    an entry with the same FIS code has the JumperID the rule gives for the
    FIS name ('BROBECK Hannah Christina' -> 'NOR BROBE HANNA') or has it as
    JumperID_pdf ('HAYBOE MICHAE'). A FIS code listed only under another
    JumperID (name change, other spelling) goes to 'Prüfen'.
    """
    missing, check = [], []
    # Team results at FIS have no nation: such an entry is checked with the
    # nation where the same name also appears with one.
    with_nsa = {(athlete.fis_code, _key(athlete.name)) for athlete in data.athletes.values() if athlete.nsa}
    for athlete in data.athletes.values():
        if not athlete.nsa and (athlete.fis_code, _key(athlete.name)) in with_nsa:
            continue
        if listen.jumper_id(athlete.name, athlete.nsa) is not None:
            continue
        same_code = listen.all_by_fis_code(athlete.fis_code)
        proposed = _key(proposed_jumper_id(athlete.name, athlete.nsa))
        if any(proposed in (_key(jumper.jumper_id), _key(jumper.pdf_id)) for jumper in same_code):
            continue
        if same_code:
            check.append((athlete, same_code[0]))
        else:
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
    used_ids = {_key(jumper.jumper_id) for jumper in listen.jumpers} | {
        _key(jumper.pdf_id) for jumper in listen.jumpers if jumper.pdf_id}
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
            jumper_id, jumper_id, athlete.name, athlete.nsa, gender,
            athlete.dob, int(athlete.fis_code), year, f"=YEAR(TODAY())-H{row_number}",
        )
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_number, column=column, value=value)
            if column == 6 and value is not None:
                cell.number_format = "yyyy-mm-dd"
        notes = []
        if _key(jumper_id) in used_ids or proposed_count[_key(jumper_id)] > 1:
            notes.append("JumperID schon vergeben, bitte anpassen")
            sheet.cell(row=row_number, column=2).fill = MARK_FILL
        if gender is None:
            notes.append("Gender unbekannt")
            sheet.cell(row=row_number, column=5).fill = MARK_FILL
        if athlete.dob is None:
            notes.append("Geburtsdatum unbekannt")
            sheet.cell(row=row_number, column=6).fill = MARK_FILL
        elif athlete.dob_source == YEAR_ONLY:
            notes.append("Geburtsdatum: nur Jahrgang bekannt")
            sheet.cell(row=row_number, column=6).fill = YEAR_FILL
        extra = (athlete.club, athlete.dob_source, athlete.first_competition, "; ".join(notes))
        for offset, value in enumerate(extra):
            sheet.cell(row=row_number, column=len(JUMPER_HEADERS) + 2 + offset, value=value or None)
    widths = {"A": 18, "B": 18, "C": 30, "D": 6, "E": 8, "F": 12, "G": 10, "H": 6, "I": 6,
              "K": 30, "L": 40, "M": 40, "N": 40}
    for letter, width in widths.items():
        sheet.column_dimensions[letter].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(JUMPER_HEADERS) + 1 + len(info)).column_letter}{max(len(rows) + 1, 1)}"

    # Prüfen: the jumper as listed under its FIS code, with the JumperID the
    # PDFs give (JumperID_pdf): rows to add to Liste_Springer_ID, e.g.
    # 'HAYBOE MICHAE' -> 'AUT HAYBOE MICHAE' for 'Hayboeck Michael' without NSA.
    check = workbook.create_sheet("Prüfen")
    _header(check, JUMPER_HEADERS)
    check_info = ("Name FIS", "NSA FIS", "Erster Wettkampf")
    _header(check, check_info, column=len(JUMPER_HEADERS) + 2)
    check_rows = sorted(
        ((proposed_jumper_id(athlete.name, athlete.nsa), athlete, listed) for athlete, listed in result.check),
        key=lambda item: (item[2].jumper_id, item[0]),
    )
    for row_number, (pdf_id, athlete, listed) in enumerate(check_rows, start=2):
        values = (
            pdf_id, listed.jumper_id, listed.name, listed.nsa or None, listed.gender, listed.dob,
            int(athlete.fis_code), listed.year, f"=YEAR(TODAY())-H{row_number}",
        )
        for column, value in enumerate(values, start=1):
            cell = check.cell(row=row_number, column=column, value=value)
            if column == 6 and value is not None:
                cell.number_format = "yyyy-mm-dd"
        for offset, value in enumerate((athlete.name, athlete.nsa, athlete.first_competition)):
            check.cell(row=row_number, column=len(JUMPER_HEADERS) + 2 + offset, value=value or None)
    for letter, width in {"A": 18, "B": 18, "C": 30, "D": 6, "E": 8, "F": 12, "G": 10, "H": 6, "I": 6,
                          "K": 30, "L": 8, "M": 40}.items():
        check.column_dimensions[letter].width = width
    check.freeze_panes = "A2"

    hills = workbook.create_sheet("Fehlende_SchanzenID")
    _header(hills, HILL_HEADERS)
    _header(hills, ("Wettkämpfe",), column=len(HILL_HEADERS) + 2)
    # SchanzenID_pdf and SchanzenID: the spelling of the PDF (change SchanzenID
    # to the name of a listed hill when it is the same hill).
    keys = ("SchanzenID", "SchanzenID", "Hill Size", "HS 95%", "K-Punkt", "Meter Value", "HWC", "TWC", "Gate Factor")
    for row_number, hill in enumerate(sorted(result.missing_hills, key=lambda hill: hill.hill_id), start=2):
        for column, key in enumerate(keys, start=1):
            hills.cell(row=row_number, column=column, value=hill.values.get(key))
        hills.cell(row=row_number, column=len(HILL_HEADERS) + 2, value="; ".join(hill.competitions))
    hills.column_dimensions["A"].width = 30
    hills.column_dimensions["B"].width = 30
    hills.column_dimensions["N"].width = 80

    summary = workbook.create_sheet("Übersicht", 0)
    lines = [
        ("Saison", season),
        ("Erstellt", datetime.now().replace(microsecond=0)),
        ("Wettkämpfe", data.competitions),
        ("Springer in der Saison", len({athlete.fis_code for athlete in data.athletes.values()})),
        ("Fehlende JumperID", len(result.missing)),
        ("FIS-Code mit anderem Namen/NSA (Prüfen)", len(result.check)),
        ("Schanzen in der Saison", len(data.hills)),
        ("Fehlende SchanzenID", len(result.missing_hills)),
        ("", ""),
        ("Einfügen", "Fehlende_JumperID A:H und Prüfen A:H in Liste_Springer_ID ab Spalte A einfügen "
                     "(Age berechnet die Tabelle; Prüfen: gleiche JumperID, neue JumperID_pdf); "
                     "Fehlende_SchanzenID A:I in Liste_Schanzen_ID, "
                     "SchanzenID (B) bei anderer Schreibweise einer bekannten Schanze anpassen, "
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
    parser = argparse.ArgumentParser(description="Fehlende JumperID und SchanzenID von FIS-Saisons.")
    parser.add_argument(
        "--season", nargs="+", required=True,
        help="FIS-Saison, z. B. 2026 (= 2025/26); mehrere: 2024 2025 2026 oder 2024-2026",
    )
    parser.add_argument("--output", type=Path, help=f"Standard: {OUTPUT_FILE} (mehrere Saisons: mit Saisons im Namen)")
    parser.add_argument("--listen", type=Path, default=SJ_LISTEN_FILE)
    parser.add_argument(
        "--archive", type=Path, default=ARCHIVE_DIRECTORY,
        help=f"PDF-Archiv (pdf_archive), dessen PDFs nicht neu geladen werden (Standard {ARCHIVE_DIRECTORY})",
    )
    args = parser.parse_args(argv)
    seasons = parse_seasons(args.season)
    label = season_label(seasons)
    output = args.output or (
        OUTPUT_FILE if len(seasons) == 1 else OUTPUT_FILE.with_name(f"{OUTPUT_FILE.stem} {label}.xlsx")
    )

    listen = read_sj_listen(args.listen)
    local_files = archive_links(args.archive)
    if local_files:
        print(f"{len(local_files)} PDFs aus dem Archiv {args.archive}.")
    fetcher = Fetcher(local_files=local_files)
    data = SeasonData()
    for season in seasons:
        collect_season(season, fetcher, data=data)
    result = compare(data, listen)
    complete_from_biography(result.missing, fetcher)
    write_workbook(result, listen, label, data, output)
    print(
        f"{len(result.missing)} fehlende JumperID, {len(result.check)} zu prüfen, "
        f"{len(result.missing_hills)} fehlende SchanzenID: {output}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
