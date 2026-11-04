import argparse
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


SECTOR_CODE = "JP"
DEFAULT_SEASON = 2027
REQUEST_DELAY_SECONDS = 2
MAX_RATE_LIMIT_RETRIES = 5
BASE_URL = "https://www.fis-ski.com/DB"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "data" / "raceid" / "raceid.xlsx"
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


@dataclass
class Element:
    tag: str
    attrs: dict[str, str]
    parent: "Element | None" = None
    children: list[object] = field(default_factory=list)

    @property
    def classes(self) -> set[str]:
        return set(self.attrs.get("class", "").split())

    @property
    def text(self) -> str:
        return " ".join(
            " ".join(child.split()) if isinstance(child, str) else child.text
            for child in self.children
            if isinstance(child, Element) or child.strip()
        ).strip()

    def find_all(self, predicate) -> list["Element"]:
        found = []
        for child in self.children:
            if isinstance(child, Element):
                if predicate(child):
                    found.append(child)
                found.extend(child.find_all(predicate))
        return found


class _DocumentParser(HTMLParser):
    _VOID_TAGS = {
        "area", "base", "br", "col", "embed", "hr", "img", "input",
        "link", "meta", "param", "source", "track", "wbr",
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Element("document", {})
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        element = Element(tag, dict(attrs), self.stack[-1])
        self.stack[-1].children.append(element)
        if tag not in self._VOID_TAGS:
            self.stack.append(element)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in self._VOID_TAGS:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, 0, -1):
            if self.stack[index].tag == tag:
                del self.stack[index:]
                break

    def handle_data(self, data):
        self.stack[-1].children.append(data)


def _parse_document(html: str) -> Element:
    parser = _DocumentParser()
    parser.feed(html)
    parser.close()
    return parser.root


def _download(url: str) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
            )
        },
    )
    for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
        try:
            with urlopen(request, timeout=30) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                return response.read().decode(charset, errors="replace")
        except HTTPError as error:
            if error.code != 429 or attempt == MAX_RATE_LIMIT_RETRIES:
                raise RuntimeError(
                    f"FIS-Seite konnte nicht geladen werden ({url}): {error}"
                ) from error

            retry_after = error.headers.get("Retry-After") if error.headers else None
            delay = _retry_delay(retry_after, attempt)
            print(
                f"FIS-Drosselung (HTTP 429); neuer Versuch in {delay} Sekunden.",
                file=sys.stderr,
            )
            time.sleep(delay)
        except (URLError, TimeoutError) as error:
            raise RuntimeError(
                f"FIS-Seite konnte nicht geladen werden ({url}): {error}"
            ) from error

    raise RuntimeError(f"FIS-Seite konnte nach Wiederholungen nicht geladen werden: {url}")


def _retry_delay(retry_after: str | None, attempt: int) -> int:
    if retry_after:
        try:
            return max(1, int(retry_after))
        except ValueError:
            try:
                retry_at = parsedate_to_datetime(retry_after)
                if retry_at.tzinfo is None:
                    retry_at = retry_at.replace(tzinfo=timezone.utc)
                return max(
                    1,
                    int((retry_at - datetime.now(timezone.utc)).total_seconds()),
                )
            except (TypeError, ValueError, OverflowError):
                pass
    return min(300, 30 * 2**attempt)


def _calendar_url(season: int) -> str:
    params = {
        "eventselection": "",
        "seasoncode": season,
        "sectorcode": SECTOR_CODE,
        "disciplinecode": "",
        "gendercode": "",
        "racecodex": "",
        "seasonmonth": f"X-{season}",
        "seasonselection": "",
    }
    return f"{BASE_URL}/ski-jumping/calendar-results.html?{urlencode(params)}"


def _detail_url(eventid: str, season: int) -> str:
    params = {"sectorcode": SECTOR_CODE, "eventid": eventid, "seasoncode": season}
    return f"{BASE_URL}/general/event-details.html?{urlencode(params)}"


def _calendar_events(html: str) -> list[tuple[str, str]]:
    document = _parse_document(html)
    calendar = next(
        (node for node in document.find_all(lambda item: item.attrs.get("id") == "calendardata")),
        None,
    )
    if calendar is None:
        raise ValueError("FIS-Kalender enthält keinen Veranstaltungsbereich (calendardata).")

    events = []
    for row in calendar.find_all(
        lambda item: "table-row" in item.classes and item.attrs.get("id", "").isdigit()
    ):
        eventid = row.attrs["id"]
        place_nodes = row.find_all(
            lambda item: "clip" in item.classes
            and "bold" in item.classes
            and "font_md_large" in item.classes
        )
        if not place_nodes:
            place_nodes = row.find_all(
                lambda item: "clip" in item.classes
                and "bold" in item.classes
                and "clip-xs" not in item.classes
            )
        place = place_nodes[0].text if place_nodes else ""
        events.append((eventid, place))
    return events


def _competition_rows(html: str, eventid: str, place: str, season: int) -> list[dict]:
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
            raise ValueError(
                f"FIS-Datum für Wettkampf {eventid}/{codexes[0]} "
                f"ist nicht eindeutig: {sorted(date_values)}"
            )
        try:
            competition_date = date.fromisoformat(date_values.pop())
        except ValueError as error:
            raise ValueError(
                f"Ungültiges FIS-Datum für Wettkampf {eventid}/{codexes[0]}"
            ) from error

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
        help=f"FIS-Saison (Standard: {DEFAULT_SEASON})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help="Excel-Ausgabedatei (Standard: data/raceid/raceid.xlsx im Projektordner)",
    )
    args = parser.parse_args()

    print(f"FIS Ski Jumping ({SECTOR_CODE}), Saison {args.season}")
    records = crawl_season(args.season)
    write_workbook(records, args.output)
    print(f"{len(records)} Wettkämpfe in {args.output} gespeichert.")


if __name__ == "__main__":
    main()
