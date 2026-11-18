import argparse
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sj_paths import DATA_DIRECTORY  # noqa: E402
from fis_web import (  # noqa: E402
    REQUEST_DELAY_SECONDS,
    SECTOR_CODE,
    calendar_events as _calendar_events,
    calendar_url as _calendar_url,
    current_season,
    detail_url as _detail_url,
    open_url,
    parse_document as _parse_document,
    read_text,
)


DEFAULT_SEASON = current_season()
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT_PATH = DATA_DIRECTORY / "SJ-Listen.xlsm"
SHEET_NAME = "Event_ID"
HEADERS = (
    "season",
    "date",
    "eventid",
    "raceid",
    "codex",
    "place",
    "hill",
    "cat1",
    "cat2",
    "gender",
    "EventNr",
    "list_event",
)
# keep leading zeros (codex 0284, list_event 001 ...)
TEXT_COLUMNS = ("eventid", "raceid", "codex", "list_event")
EXCEL_EPOCH = date(1899, 12, 30)
HEADER_FONT_COLOR = 0xFFFFFF  # Excel colours are BGR: white
HEADER_FILL_COLOR = 0x8A3517  # RGB 17358A (dark blue) as BGR
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # do not run the workbook's macros
# cat2 of a qualification: the championship it belongs to, else summer
# (May to October) Grand Prix, else World Cup.
SUMMER_MONTHS = range(5, 11)
CHAMPIONSHIP_CATEGORIES = ("OWG", "WSC", "SFWC", "WJC")
# cat2 of other categories that differ from cat1 (FIS Cup is FC, FESA is
# the Alpen Cup AC).
CATEGORY_2 = {"FIS": "FC", "FESA": "AC"}
# Competitions of one event lie within 5 calendar days (e.g. Friday to Tuesday).
EVENT_SPAN = timedelta(days=4)
# FIS describes each competition as e.g. "Individual Large Hill Qualification".
HILL_TYPES = {
    "Small": "SH",
    "Medium": "MH",
    "Normal": "NH",
    "Large": "LH",
    "Flying": "FH",
}
HILL_PATTERN = re.compile(r"\b(" + "|".join(HILL_TYPES) + r") Hill\b")


def _warn(message: str) -> None:
    print(f"Warnung: {message}", file=sys.stderr)


def _download(url: str) -> str:
    def report_retry(delay: int) -> None:
        print(
            f"FIS-Drosselung (HTTP 429); neuer Versuch in {delay} Sekunden.",
            file=sys.stderr,
        )

    with open_url(url, timeout=30, on_retry=report_retry) as response:
        return read_text(response)


def _competition_rows(
    html: str,
    eventid: str,
    place: str,
    season: int,
    warn=_warn,
) -> list[dict]:
    document = _parse_document(html)
    competitions = []
    for row in document.find_all(lambda item: "table-row" in item.classes):
        links = row.find_all(
            lambda item: item.tag == "a" and "results.html" in item.attrs.get("href", "")
        )
        if not links:
            continue

        race_ids = set()
        for link in links:
            query = parse_qs(urlparse(link.attrs["href"]).query)
            if query.get("sectorcode") == [SECTOR_CODE] and query.get("raceid"):
                race_ids.update(query["raceid"])
        if len(race_ids) != 1:
            continue
        raceid = race_ids.pop()

        codexes = []
        for link in row.find_all(lambda item: item.tag == "a"):
            codex_match = re.match(r"^\s*(\d{4})(?:\b|$)", link.text)
            if codex_match:
                codexes.append(codex_match.group(1))
        categories = [
            link.text
            for link in links
            if 2 <= len(link.text) <= 6
            and link.text.isalnum()
            and link.text.isupper()
            and link.text not in {"M", "W"}
        ]
        genders = [link.text for link in links if link.text in {"M", "W", "A"}]
        if not codexes or not categories or not genders:
            continue
        date_values = {
            node.attrs["data-date"]
            for node in row.find_all(
                lambda item: "timezone-date" in item.classes
                and item.attrs.get("data-date")
            )
        }
        if len(date_values) != 1:
            warn(
                f"Wettkampf {eventid}/{codexes[0]} übersprungen: FIS-Datum "
                f"ist nicht eindeutig ({sorted(date_values)})."
            )
            continue
        date_value = date_values.pop()
        try:
            competition_date = date.fromisoformat(date_value)
        except ValueError:
            warn(
                f"Wettkampf {eventid}/{codexes[0]} übersprungen: "
                f"ungültiges FIS-Datum {date_value!r}."
            )
            continue

        hill_match = HILL_PATTERN.search(row.text)
        hill = HILL_TYPES[hill_match.group(1)] if hill_match else None

        for gender in dict.fromkeys(genders):
            competitions.append(
                {
                    "season": season,
                    "date": competition_date,
                    "eventid": eventid,
                    "raceid": raceid,
                    "codex": codexes[0],
                    "place": place,
                    "category": categories[0],
                    "gender": gender,
                    "hill": hill,
                }
            )
    return competitions


def crawl_season(season: int) -> list[dict]:
    events = _calendar_events(_download(_calendar_url(season)))
    records = []
    for index, (eventid, place) in enumerate(events, start=1):
        print(f"[{index}/{len(events)}] Event {eventid}: {place}")
        html = _download(_detail_url(eventid, season))
        records.extend(_competition_rows(html, eventid, place, season))
        if index < len(events):
            time.sleep(REQUEST_DELAY_SECONDS)
    return records


def _sorted(records: list[dict]) -> list[dict]:
    return sorted(
        records,
        key=lambda record: (record["season"], record["date"], int(record["codex"])),
    )


def event_numbers(records: list[dict]) -> list[int]:
    """Consecutive event number for each record of ``_sorted(records)``.

    One event (e.g. World Cup men in Klingenthal: qualification, Saturday and
    Sunday) has the same season, place, cat2, gender and hill type (NH/LH/FH)
    and all its competitions lie within 5 calendar days of the first one.
    Events are numbered by the date of their last competition, then their
    first competition and their smallest codex.
    """
    groups = []
    last_number = 0
    open_events = {}  # key -> (group, date of the first competition)
    for record, category_2 in zip(records, second_categories(records)):
        key = (
            record["season"],
            record["place"],
            category_2,
            record["gender"],
            record.get("hill"),
        )
        current = open_events.get(key)
        if current is None or record["date"] - current[1] > EVENT_SPAN:
            last_number += 1
            current = (last_number, record["date"])
            open_events[key] = current
        groups.append(current[0])

    members = {}
    for group, record in zip(groups, records):
        members.setdefault(group, []).append(record)
    order = sorted(
        members,
        key=lambda group: (
            max(record["date"] for record in members[group]),
            min(record["date"] for record in members[group]),
            min(int(record["codex"]) for record in members[group]),
        ),
    )
    number_of = {group: number for number, group in enumerate(order, start=1)}
    return [number_of[group] for group in groups]


def second_category(category: str, competition_date: date, championship: str | None = None) -> str:
    """cat2: the category, but FIS becomes FC, FESA becomes AC and QUA becomes the championship
    (e.g. WSC) it belongs to, else GP (May to October) or WC."""
    if category != "QUA":
        return CATEGORY_2.get(category, category)
    if championship:
        return championship
    return "GP" if competition_date.month in SUMMER_MONTHS else "WC"


def _championship(qualification: dict, records: list[dict]) -> str | None:
    """Championship category held on the same hill (place, gender, hill type)
    within 5 calendar days of the qualification, if any."""
    for record in records:
        if (
            record["category"] in CHAMPIONSHIP_CATEGORIES
            and record["season"] == qualification["season"]
            and record["place"] == qualification["place"]
            and record["gender"] == qualification["gender"]
            and record.get("hill") == qualification.get("hill")
            and abs(record["date"] - qualification["date"]) <= EVENT_SPAN
        ):
            return record["category"]
    return None


def second_categories(records: list[dict]) -> list[str]:
    """cat2 for each record."""
    return [
        second_category(
            record["category"],
            record["date"],
            _championship(record, records) if record["category"] == "QUA" else None,
        )
        for record in records
    ]


def sheet_rows(records: list[dict]) -> list[list]:
    """Header and records sorted by EventNr, date and codex, as Excel values.

    EventNr counts the events in date order. Dates become Excel serial numbers
    (shown as yyyy-mm-dd), so no time zone conversion can shift them.
    """
    records = _sorted(records)
    numbered = sorted(
        zip(event_numbers(records), records, second_categories(records)),
        key=lambda item: (item[0], item[1]["date"], int(item[1]["codex"])),
    )
    rows = [list(HEADERS)]
    for number, record, category_2 in numbered:
        values = dict(record)
        values["date"] = (record["date"] - EXCEL_EPOCH).days
        values["cat1"] = record["category"]
        values["cat2"] = category_2
        values["EventNr"] = number
        values["list_event"] = list_event(
            number,
            record["date"],
            category_2,
            record["gender"],
            record["place"],
            record.get("hill"),
            record["codex"],
        )
        rows.append([values[header] for header in HEADERS])
    return rows


def list_event(
    number: int,
    competition_date: date,
    category_2: str,
    gender: str,
    place: str,
    hill,
    codex: str,
) -> str:
    """Name of a competition for a drop-down list:
    '001 2026-07-04 CHIW Rastbuchl SH 0284'."""
    return " ".join(
        part
        for part in (
            f"{number:03d}",
            competition_date.isoformat(),
            f"{category_2}{gender}",
            place,
            hill or "",
            codex,
        )
        if part
    )


def _local_date_format(excel) -> str:
    """'yyyy-mm-dd' with the date letters of the Excel language (German: JJJJ-MM-TT)."""
    # Application.International: xlYearCode 19, xlMonthCode 20, xlDayCode 21
    # (late-bound COM returns the whole 1-based array).
    international = excel.International
    year, month, day = international[18], international[19], international[20]
    return f"{year * 4}-{month * 2}-{day * 2}"


class WorkbookOpenError(RuntimeError):
    """Raised when the workbook is open in Excel and cannot be written."""


def _excel_application():
    import win32com.client

    # A separate Excel instance: a workbook the user has open stays untouched.
    return win32com.client.DispatchEx("Excel.Application")


def check_writable(workbook_path: Path) -> None:
    """Fail before crawling when the workbook is missing or open in Excel."""
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel-Datei nicht gefunden: {workbook_path}")
    lock_file = workbook_path.with_name(f"~${workbook_path.name}")
    if lock_file.exists():
        raise WorkbookOpenError(
            f"{workbook_path.name} ist in Excel geöffnet. Bitte schliessen und neu starten."
        )


def write_event_sheet(records: list[dict], workbook_path: Path, excel_factory=None) -> None:
    """Replace the content of the sheet Event_ID in SJ-Listen.xlsm.

    Excel itself writes the workbook, so macros, buttons, tables and formulas
    of the other sheets are kept. Only the sheet Event_ID is changed and the
    file is saved once.
    """
    if not records:
        raise ValueError(f"Keine Wettkämpfe gefunden; {SHEET_NAME} bleibt unverändert.")
    check_writable(workbook_path)

    rows = sheet_rows(records)
    excel = (excel_factory or _excel_application)()
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), 0, False)
        if workbook.ReadOnly:
            raise WorkbookOpenError(
                f"{workbook_path.name} ist schreibgeschützt oder in Excel geöffnet. "
                "Bitte schliessen und neu starten."
            )
        if SHEET_NAME not in [sheet.Name for sheet in workbook.Worksheets]:
            raise ValueError(f"Sheet {SHEET_NAME} fehlt in {workbook_path.name}.")
        sheet = workbook.Worksheets(SHEET_NAME)
        old_headers = [sheet.Cells(1, index).Value for index in range(1, len(HEADERS) + 1)]
        if sheet.AutoFilterMode:
            sheet.AutoFilterMode = False
        # Only the values: formats, conditional formats and column widths stay.
        sheet.UsedRange.ClearContents()

        data = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(rows), len(HEADERS)))
        for index, header in enumerate(HEADERS, start=1):
            column = sheet.Range(sheet.Cells(2, index), sheet.Cells(len(rows), index))
            if header in TEXT_COLUMNS:
                column.NumberFormat = "@"
            elif header == "date":
                column.NumberFormatLocal = _local_date_format(excel)
        data.Value = tuple(tuple(row) for row in rows)

        # Format only columns that are new in this sheet; the user's formats of
        # the existing columns are kept.
        for index, header in enumerate(HEADERS, start=1):
            if old_headers[index - 1] == header:
                continue
            cell = sheet.Cells(1, index)
            cell.Font.Bold = True
            cell.Font.Color = HEADER_FONT_COLOR
            cell.Interior.Color = HEADER_FILL_COLOR
            data.Columns(index).AutoFit()
        data.AutoFilter()
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Erstellt eine Excel-Liste der FIS-Skisprung-Wettkämpfe."
    )
    parser.add_argument(
        "--season",
        type=int,
        default=DEFAULT_SEASON,
        help=f"FIS-Saison (Standard: aktuelle Saison, {DEFAULT_SEASON})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help="Excel-Datei mit dem Sheet Event_ID (Standard: SJ-Listen.xlsm im SJpdf-Datenordner)",
    )
    args = parser.parse_args()

    print(f"FIS Ski Jumping ({SECTOR_CODE}), Saison {args.season}")
    try:
        check_writable(args.output)
        records = crawl_season(args.season)
        write_event_sheet(records, args.output)
    except (WorkbookOpenError, FileNotFoundError, ValueError) as error:
        print(f"Fehler: {error}", file=sys.stderr)
        raise SystemExit(1) from error
    print(f"{len(records)} Wettkämpfe in {args.output} (Sheet {SHEET_NAME}) gespeichert.")


if __name__ == "__main__":
    main()
