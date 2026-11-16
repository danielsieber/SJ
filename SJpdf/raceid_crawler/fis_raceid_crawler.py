import argparse
import re
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

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
DEFAULT_OUTPUT_PATH = DATA_DIRECTORY / "raceid" / "raceid.xlsx"
HEADERS = (
    "season",
    "date",
    "eventid",
    "raceid",
    "codex",
    "place",
    "category",
    "gender",
)


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


def write_workbook(records: list[dict], output_path: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name, gender_filter in (
        ("men", "M"),
        ("women", "W"),
        ("mixed", "A"),
        ("all", None),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append(HEADERS)
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17358A")

        sheet_records = sorted(
            (
                record
                for record in records
                if gender_filter is None or record["gender"] == gender_filter
            ),
            key=lambda record: (
                record["season"],
                record["date"],
                int(record["codex"]),
            ),
        )
        for record in sheet_records:
            sheet.append([record[header] for header in HEADERS])

        for cell in sheet["B"][1:]:
            cell.number_format = "yyyy-mm-dd"
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            width = max(len(str(cell.value or "")) for cell in column_cells) + 2
            sheet.column_dimensions[column_cells[0].column_letter].width = width

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


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
        help="Excel-Ausgabedatei (Standard: raceid/raceid.xlsx im SJpdf-Datenordner)",
    )
    args = parser.parse_args()

    print(f"FIS Ski Jumping ({SECTOR_CODE}), Saison {args.season}")
    records = crawl_season(args.season)
    write_workbook(records, args.output)
    print(f"{len(records)} Wettkämpfe in {args.output} gespeichert.")


if __name__ == "__main__":
    main()
