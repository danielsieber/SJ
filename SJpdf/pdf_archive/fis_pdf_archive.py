"""Download all FIS ski jumping PDFs of several seasons into one folder per event.

An event is a series of competitions with the same EventNr as in the sheet
Event_ID (same place, cat2, gender and hill type within 5 days, see
raceid_crawler/fis_raceid_crawler.py). Its folder is named like list_event
without EventNr and codex, with the date of its last competition:

    C:\\downloaded_FIS_pdf\\2027-03-14 COCM Zakopane LH

Each folder holds every PDF of all its competitions (start lists, results,
program, ...) under the FIS file name, plus pdf_links.csv (PDF name and FIS
link, read by pdf2xlsx). A PDF whose link is already in pdf_links.csv is not
downloaded again, so an interrupted run continues where it stopped and a
later run of the current season only adds new PDFs.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "season_check"))
from fis_web import (  # noqa: E402
    REQUEST_DELAY_SECONDS,
    calendar_events,
    calendar_url,
    current_season,
    detail_url,
    open_url,
)
from fis_season_check import Fetcher, _downloader  # noqa: E402
from raceid_crawler.fis_raceid_crawler import (  # noqa: E402
    _competition_rows,
    _sorted,
    event_numbers,
    second_categories,
)

OUTPUT_DIRECTORY = Path(r"C:\downloaded_FIS_pdf")
FIRST_SEASON = 2016
LINKS_FILENAME = "pdf_links.csv"  # as written by pdf_download/FIS_pdf-download.py
LOG_FILENAME = "archiv_log.csv"
LOG_HEADERS = ("Zeit", "Saison", "Ordner", "Codex", "PDFs neu", "PDFs gesamt", "Meldung")


@dataclass
class Event:
    """One EventNr of a season: its competitions and their PDF links."""

    season: int
    number: int
    records: list[dict] = field(default_factory=list)
    category_2: str = ""
    attachments: dict[str, object] = field(default_factory=dict)  # url -> attachment

    @property
    def last_date(self) -> date:
        return max(record["date"] for record in self.records)

    @property
    def codexes(self) -> list[str]:
        return sorted({record["codex"] for record in self.records}, key=int)


def folder_name(event: Event) -> str:
    """'2027-03-14 COCM Zakopane LH': list_event without EventNr and codex,
    with the date of the last competition."""
    first = event.records[0]
    parts = (
        event.last_date.isoformat(),
        f"{event.category_2}{first['gender']}",
        first["place"],
        first.get("hill") or "",
    )
    name = " ".join(part for part in parts if part)
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")


def season_events(season: int, fetcher: Fetcher, log: Callable[[str], None] = print) -> list[Event]:
    """All events (EventNr) of a season with the PDF links of their competitions."""
    records = []
    attachments_by_codex: dict[str, list] = defaultdict(list)
    events = calendar_events(fetcher.text(calendar_url(season)))
    for index, (eventid, place) in enumerate(events, start=1):
        log(f"Saison {season} [{index}/{len(events)}] Event {eventid}: {place}")
        html = fetcher.text(detail_url(eventid, season))
        records.extend(_competition_rows(html, eventid, place, season, warn=log))
        for competition in _downloader.competition_rows(html):
            attachments_by_codex[competition.codex].extend(competition.attachments)
    records = _sorted(records)
    grouped: dict[int, Event] = {}
    for record, number, category_2 in zip(records, event_numbers(records), second_categories(records)):
        # All competitions of one EventNr have the same cat2 (see event_numbers).
        event = grouped.setdefault(number, Event(season, number, category_2=category_2))
        event.records.append(record)
        for attachment in attachments_by_codex.get(record["codex"], ()):
            event.attachments.setdefault(attachment.url, attachment)
    return [grouped[number] for number in sorted(grouped)]


def read_links(folder: Path) -> dict[str, str]:
    """{FIS link: PDF name} of a folder."""
    links_file = folder / LINKS_FILENAME
    if not links_file.is_file():
        return {}
    with links_file.open(encoding="utf-8-sig", newline="") as handle:
        return {row["url"]: row["pdf"] for row in csv.DictReader(handle) if row.get("url")}


def _append_link(folder: Path, pdf_name: str, url: str) -> None:
    links_file = folder / LINKS_FILENAME
    new_file = not links_file.is_file()
    with links_file.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if new_file:
            writer.writerow(["pdf", "url"])
        writer.writerow([pdf_name, url])


class PdfDownloader:
    """Download PDFs with the FIS request delay (no cache: they go to the archive)."""

    def __init__(self, log: Callable[[str], None] = print):
        self.log = log
        self._last_request = 0.0

    def download(self, url: str) -> tuple[str, bytes]:
        wait = REQUEST_DELAY_SECONDS - (time.monotonic() - self._last_request)
        if wait > 0:
            time.sleep(wait)

        def report_retry(delay: int) -> None:
            self.log(f"FIS-Drosselung (HTTP 429); neuer Versuch in {delay} Sekunden.")

        try:
            with open_url(url, timeout=60, on_retry=report_retry) as response:
                content = response.read()
                filename = _downloader._pdf_filename(response.headers, url)
        finally:
            self._last_request = time.monotonic()
        if not content.startswith(b"%PDF"):
            raise ValueError(f"kein gültiges PDF: {url}")
        return filename, content


def archive_event(event: Event, output_directory: Path, downloader: PdfDownloader, log=print) -> tuple[Path, int, list[str]]:
    """Download the missing PDFs of an event; return (folder, new PDFs, problems)."""
    folder = output_directory / folder_name(event)
    known = read_links(folder)
    missing = [attachment for url, attachment in event.attachments.items() if url not in known]
    problems = []
    added = 0
    if missing:
        folder.mkdir(parents=True, exist_ok=True)
    for attachment in missing:
        label = " ".join(attachment.label.split())
        try:
            filename, content = downloader.download(attachment.url)
        except Exception as error:  # noqa: BLE001 - report, keep going
            problems.append(f"{label}: {error}")
            continue
        destination = _downloader._unique_path(folder, filename)
        destination.write_bytes(content)
        _append_link(folder, destination.name, attachment.url)
        added += 1
        log(f"  {folder.name}: {destination.name}")
    return folder, added, problems


def _log_row(output_directory: Path, values) -> None:
    log_file = output_directory / LOG_FILENAME
    new_file = not log_file.is_file()
    with log_file.open("a", encoding="utf-8-sig" if new_file else "utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        if new_file:
            writer.writerow(LOG_HEADERS)
        writer.writerow(values)


def archive_seasons(
    first: int, last: int, output_directory: Path, fetcher: Fetcher | None = None, log=print
) -> None:
    """Newest season first, so the most recent events are there first."""
    output_directory.mkdir(parents=True, exist_ok=True)
    downloader = PdfDownloader(log)
    for season in range(last, first - 1, -1):
        # FIS pages of a finished season do not change: cached. Pages of the
        # current season get new PDFs: always read fresh.
        season_fetcher = fetcher or Fetcher(log=log, use_cache=season < current_season())
        events = season_events(season, season_fetcher, log)
        total_new = 0
        for event in events:
            if not event.attachments:
                continue  # no PDFs at FIS (yet)
            folder, added, problems = archive_event(event, output_directory, downloader, log)
            total_new += added
            if added or problems:
                _log_row(output_directory, (
                    datetime.now().replace(microsecond=0).isoformat(sep=" "),
                    season, folder.name, " ".join(event.codexes), added,
                    len(read_links(folder)), "; ".join(problems),
                ))
        log(f"Saison {season}: {len(events)} Events, {total_new} neue PDFs.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Alle FIS-Skisprung-PDFs pro Event herunterladen.")
    parser.add_argument("--from", dest="first", type=int, default=FIRST_SEASON, help=f"erste Saison (Standard {FIRST_SEASON})")
    parser.add_argument("--to", dest="last", type=int, default=current_season(), help="letzte Saison (Standard: aktuelle)")
    parser.add_argument("--output", type=Path, default=OUTPUT_DIRECTORY)
    args = parser.parse_args(argv)
    if args.first > args.last:
        parser.error("--from muss kleiner oder gleich --to sein")
    archive_seasons(args.first, args.last, args.output)
    print(f"Fertig: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
