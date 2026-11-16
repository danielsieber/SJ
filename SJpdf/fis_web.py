"""Shared helpers for reading FIS ski jumping web pages.

Used by ``raceid_crawler`` and ``pdf_download``: a small HTML tree parser,
HTTP requests with HTTP 429 retries, FIS URLs, and the calendar parser.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from typing import Callable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

BASE_URL = "https://www.fis-ski.com/DB"
SECTOR_CODE = "JP"
REQUEST_DELAY_SECONDS = 2
MAX_RATE_LIMIT_RETRIES = 5
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
)


def current_season(today: date | None = None) -> int:
    """Return the FIS season code: a season starting in July 2026 is 2027."""
    today = date.today() if today is None else today
    return today.year + 1 if today.month >= 7 else today.year


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
        parts = []
        for child in self.children:
            if isinstance(child, Element):
                parts.append(child.text)
            elif child.strip():
                parts.append(child.strip())
        return " ".join(" ".join(parts).split())

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


def parse_document(html: str) -> Element:
    parser = _DocumentParser()
    parser.feed(html)
    parser.close()
    return parser.root


def retry_delay(retry_after: str | None, attempt: int) -> int:
    """Seconds to wait after HTTP 429, from Retry-After or exponential backoff."""
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


def open_url(
    url: str,
    timeout: int = 30,
    on_retry: Callable[[int], None] | None = None,
):
    """Open ``url`` and return the response; retry HTTP 429 with backoff.

    ``on_retry`` is called with the delay in seconds before each retry.
    """
    request = Request(url, headers={"User-Agent": USER_AGENT})
    for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
        try:
            return urlopen(request, timeout=timeout)
        except HTTPError as error:
            if error.code != 429 or attempt == MAX_RATE_LIMIT_RETRIES:
                raise RuntimeError(f"FIS-Anfrage fehlgeschlagen ({url}): {error}") from error
            header = error.headers.get("Retry-After") if error.headers else None
            delay = retry_delay(header, attempt)
            if on_retry is not None:
                on_retry(delay)
            time.sleep(delay)
        except (URLError, TimeoutError) as error:
            raise RuntimeError(f"FIS-Anfrage fehlgeschlagen ({url}): {error}") from error
    raise RuntimeError(f"FIS-Anfrage konnte nicht abgeschlossen werden: {url}")


def read_text(response) -> str:
    """Decode an HTTP response body with its declared charset."""
    charset = response.headers.get_content_charset() or "utf-8"
    return response.read().decode(charset, errors="replace")


def calendar_url(season: int) -> str:
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


def detail_url(eventid: str, season: int) -> str:
    params = {"sectorcode": SECTOR_CODE, "eventid": eventid, "seasoncode": season}
    return f"{BASE_URL}/general/event-details.html?{urlencode(params)}"


def calendar_events(html: str) -> list[tuple[str, str]]:
    """Return ``(eventid, place)`` for every event on a FIS calendar page."""
    document = parse_document(html)
    calendar = next(
        iter(document.find_all(lambda item: item.attrs.get("id") == "calendardata")),
        None,
    )
    if calendar is None:
        raise ValueError("FIS-Kalender enthält keinen Veranstaltungsbereich (calendardata).")

    events = []
    for row in calendar.find_all(
        lambda item: "table-row" in item.classes and item.attrs.get("id", "").isdigit()
    ):
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
        events.append((row.attrs["id"], place))
    return events
