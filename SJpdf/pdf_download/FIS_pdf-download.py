import queue
import re
import shutil
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from email.message import Message
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RACEID_FILE = PROJECT_ROOT / "data" / "raceid" / "raceid.xlsx"
PDF_DIRECTORY = PROJECT_ROOT / "data" / "pdf"
BASE_URL = "https://www.fis-ski.com/DB"
SECTOR_CODE = "JP"
REQUEST_DELAY_SECONDS = 2
MAX_RATE_LIMIT_RETRIES = 5
CODEX_PATTERN = re.compile(r"^\s*(\d{4})(?:\b|$)")


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


def normalize_codex(value: str) -> str:
    codex = value.strip()
    if not codex.isdigit() or not 2 <= len(codex) <= 4:
        raise ValueError("Bitte einen zwei- bis vierstelligen numerischen Codex eingeben.")
    return codex.zfill(4)


@dataclass(frozen=True)
class Candidate:
    season: int
    eventid: str
    codex: str
    gender: str
    place: str = ""
    event_date: str = ""
    category: str = ""

    @property
    def description(self) -> str:
        date_text = self.event_date
        if isinstance(date_text, (date, datetime)):
            date_text = date_text.strftime("%Y-%m-%d")
        details = " - ".join(part for part in (str(date_text), self.place) if part)
        return (
            f"{self.codex} | {self.season} | {details or 'Veranstaltung'} "
            f"| {self.gender} | Event {self.eventid}"
        )


@dataclass(frozen=True)
class Attachment:
    url: str
    label: str


@dataclass(frozen=True)
class Competition:
    codex: str
    raceid: str
    gender: str
    event_date: str
    category: str
    attachments: tuple[Attachment, ...]


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_workbook_candidates(workbook_path: Path, codex: str) -> list[Candidate]:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Race-ID-Datei nicht gefunden: {workbook_path}")

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        sheet_name = "all" if "all" in workbook.sheetnames else None
        sheet_names = [sheet_name] if sheet_name else [
            name for name in ("men", "women", "mixed") if name in workbook.sheetnames
        ]
        candidates = {}
        for name in sheet_names:
            rows = workbook[name].iter_rows(values_only=True)
            headers = next(rows, None)
            if headers is None:
                continue
            columns = {str(header).strip().lower(): index for index, header in enumerate(headers)}
            required = {"season", "eventid", "codex", "gender"}
            missing = required.difference(columns)
            if missing:
                raise ValueError(
                    f"Spalten fehlen im Sheet '{name}': {', '.join(sorted(missing))}"
                )

            for row in rows:
                row_codex = _cell_text(row[columns["codex"]]).zfill(4)
                if row_codex != codex:
                    continue
                season = int(row[columns["season"]])
                eventid = _cell_text(row[columns["eventid"]])
                gender = _cell_text(row[columns["gender"]]).upper()
                key = (season, eventid, row_codex, gender)
                event_date = row[columns["date"]] if "date" in columns else ""
                if isinstance(event_date, datetime):
                    event_date = event_date.date()
                candidates[key] = Candidate(
                    season=season,
                    eventid=eventid,
                    codex=row_codex,
                    gender=gender,
                    place=_cell_text(row[columns["place"]]) if "place" in columns else "",
                    event_date=event_date,
                    category=(
                        _cell_text(row[columns["category"]])
                        if "category" in columns
                        else ""
                    ),
                )
        return sorted(
            candidates.values(),
            key=lambda candidate: (candidate.season, candidate.event_date, candidate.eventid),
        )
    finally:
        workbook.close()


def _request(url: str):
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
            return urlopen(request, timeout=45)
        except HTTPError as error:
            if error.code != 429 or attempt == MAX_RATE_LIMIT_RETRIES:
                raise RuntimeError(f"FIS-Anfrage fehlgeschlagen ({url}): {error}") from error
            retry_after = error.headers.get("Retry-After") if error.headers else None
            delay = _retry_delay(retry_after, attempt)
            time.sleep(delay)
        except (URLError, TimeoutError) as error:
            raise RuntimeError(f"FIS-Anfrage fehlgeschlagen ({url}): {error}") from error
    raise RuntimeError(f"FIS-Anfrage konnte nicht abgeschlossen werden: {url}")


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
    document = parse_document(html)
    calendar = next(
        (
            node
            for node in document.find_all(lambda item: item.attrs.get("id") == "calendardata")
        ),
        None,
    )
    if calendar is None:
        raise ValueError("FIS-Kalender enthält keinen Veranstaltungsbereich.")
    events = []
    for row in calendar.find_all(
        lambda item: "table-row" in item.classes and item.attrs.get("id", "").isdigit()
    ):
        place_nodes = row.find_all(
            lambda item: "clip" in item.classes
            and "bold" in item.classes
            and "font_md_large" in item.classes
        )
        place = place_nodes[0].text if place_nodes else ""
        events.append((row.attrs["id"], place))
    return events


def competition_rows(html: str) -> list[Competition]:
    document = parse_document(html)
    competitions = []
    seen_races = set()
    for row in document.find_all(lambda item: "table-row" in item.classes):
        anchors = row.find_all(lambda item: item.tag == "a")
        result_links = [
            link
            for link in anchors
            if "results.html" in link.attrs.get("href", "")
        ]
        race_ids = set()
        for link in result_links:
            query = parse_qs(urlparse(link.attrs["href"]).query)
            if query.get("sectorcode") == [SECTOR_CODE] and query.get("raceid"):
                race_ids.update(query["raceid"])
        if len(race_ids) != 1:
            continue
        raceid = race_ids.pop()
        if raceid in seen_races:
            continue
        seen_races.add(raceid)

        codexes = []
        for link in anchors:
            codex_match = CODEX_PATTERN.match(link.text)
            if codex_match:
                codexes.append(codex_match.group(1))
        genders = {
            node.text
            for node in row.find_all(lambda item: "gender__item" in item.classes)
            if node.text in {"M", "W", "A"}
        }
        genders.update(link.text for link in anchors if link.text in {"M", "W", "A"})
        date_nodes = row.find_all(
            lambda item: "timezone-date" in item.classes and item.attrs.get("data-date")
        )
        event_date = date_nodes[0].attrs["data-date"] if date_nodes else ""
        categories = [
            link.text
            for link in result_links
            if 2 <= len(link.text) <= 6
            and link.text.isalnum()
            and link.text.isupper()
            and link.text not in {"M", "W", "A"}
        ]
        attachments = []
        for link in anchors:
            href = link.attrs.get("href", "")
            if "/competition-attachment/" not in href or not href.lower().endswith(".pdf"):
                continue
            label_node = link.parent
            while label_node is not None and "drop-btn__item" not in label_node.classes:
                label_node = label_node.parent
            label = label_node.text if label_node else link.text
            attachments.append(Attachment(href, label))

        for codex in dict.fromkeys(codexes):
            for gender in sorted(genders):
                competitions.append(
                    Competition(
                        codex=codex,
                        raceid=raceid,
                        gender=gender,
                        event_date=event_date,
                        category=categories[0] if categories else "",
                        attachments=tuple(attachments),
                    )
                )
    return competitions


def find_fis_candidates(
    codex: str,
    seasons: list[int],
    progress=lambda message: None,
) -> list[Candidate]:
    for season in sorted(set(seasons), reverse=True):
        progress(f"Codex {codex} fehlt in der Excel-Datei; durchsuche FIS-Saison {season}.")
        with _request(_calendar_url(season)) as response:
            calendar_html = response.read().decode(
                response.headers.get_content_charset() or "utf-8",
                errors="replace",
            )
        events = _calendar_events(calendar_html)
        for index, (eventid, place) in enumerate(events, start=1):
            progress(
                f"Suche Codex {codex} auf FIS-Eventseiten ({index}/{len(events)}): "
                f"{place or eventid}"
            )
            with _request(_detail_url(eventid, season)) as response:
                event_html = response.read().decode(
                    response.headers.get_content_charset() or "utf-8",
                    errors="replace",
                )
            matches = [
                Candidate(
                    season=season,
                    eventid=eventid,
                    codex=competition.codex,
                    gender=competition.gender,
                    place=place,
                    event_date=competition.event_date,
                    category=competition.category,
                )
                for competition in competition_rows(event_html)
                if competition.codex == codex
            ]
            if matches:
                return matches
            if index < len(events):
                time.sleep(REQUEST_DELAY_SECONDS)
    return []


def find_candidates(
    codex: str,
    workbook_path: Path = RACEID_FILE,
    progress=lambda message: None,
) -> list[Candidate]:
    candidates = read_workbook_candidates(workbook_path, codex)
    if candidates:
        return candidates

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        seasons = set()
        for sheet_name in workbook.sheetnames:
            rows = workbook[sheet_name].iter_rows(values_only=True)
            headers = next(rows, None)
            if headers is None or "season" not in headers:
                continue
            season_column = headers.index("season")
            seasons.update(int(row[season_column]) for row in rows if row[season_column])
    finally:
        workbook.close()
    if not seasons:
        seasons = {2027}
    return find_fis_candidates(codex, sorted(seasons), progress)


def _pdf_filename(headers, url: str) -> str:
    disposition = headers.get("Content-Disposition", "")
    message = Message()
    if disposition:
        message["Content-Disposition"] = disposition
    filename = message.get_filename()
    if not filename:
        filename = Path(urlparse(url).path).name
    filename = Path(filename.replace("\\", "/")).name
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
    if not filename or not filename.lower().endswith(".pdf"):
        raise ValueError(f"FIS lieferte keinen gültigen PDF-Dateinamen für {url}")
    return filename


def _unique_path(directory: Path, filename: str) -> Path:
    original = Path(filename)
    path = directory / original.name
    suffix = 2
    while path.exists():
        path = directory / f"{original.stem} ({suffix}){original.suffix}"
        suffix += 1
    return path


def archive_existing_files(
    output_directory: Path = PDF_DIRECTORY,
    progress=lambda message: None,
) -> Path | None:
    if not output_directory.is_dir():
        return None

    files = [path for path in output_directory.iterdir() if path.is_file()]
    if not files:
        return None

    folder_name_file = output_directory / "myFolderName.txt"
    if not folder_name_file.is_file():
        raise FileNotFoundError(
            "PDF-Ordner enthält Dateien, aber data/pdf/myFolderName.txt fehlt. "
            "Dateien bleiben unverändert; bitte die Zuordnung ergänzen."
        )

    folder_name = folder_name_file.read_text(encoding="utf-8-sig").strip()
    if (
        not folder_name
        or folder_name in {".", ".."}
        or Path(folder_name).name != folder_name
        or any(character in folder_name for character in '<>:"/\\|?*')
        or any(ord(character) < 32 for character in folder_name)
    ):
        raise ValueError(
            f"Ungültiger Ordnername in {folder_name_file}; "
            "Dateien bleiben unverändert."
        )

    archive_directory = output_directory / "archiv" / folder_name
    if archive_directory.resolve() == output_directory.resolve():
        raise ValueError(
            f"Der Archivordner darf nicht der Download-Ordner sein: {archive_directory}"
        )
    archive_directory.mkdir(parents=True, exist_ok=True)

    for source in files:
        destination = _unique_path(archive_directory, source.name)
        shutil.move(str(source), str(destination))
        progress(f"Archiviert: {source.name} -> {destination}")

    return archive_directory


def folder_name_for_group(
    candidate: Candidate,
    competitions: list[Competition],
) -> str:
    group = [
        competition
        for competition in competitions
        if competition.gender == candidate.gender
    ]
    dated_competitions = []
    for competition in group:
        if competition.event_date:
            try:
                competition_day = date.fromisoformat(competition.event_date)
            except ValueError as error:
                raise ValueError(
                    f"Ungültiges FIS-Datum für Codex {competition.codex}: "
                    f"{competition.event_date}"
                ) from error
            dated_competitions.append((competition_day, competition))

    if not dated_competitions:
        raise ValueError(
            f"Für Event {candidate.eventid} und Gender {candidate.gender} "
            "wurde kein Wettkampfdatum gefunden."
        )
    newest_date = max(competition_day for competition_day, _ in dated_competitions)
    newest_competitions = [
        competition
        for competition_day, competition in dated_competitions
        if competition_day == newest_date
    ]
    categories = {competition.category for competition in newest_competitions}
    categories.discard("")
    if len(categories) != 1:
        raise ValueError(
            f"Excel-Kategorie für das neueste Datum {newest_date} von Event "
            f"{candidate.eventid} ist nicht eindeutig: {sorted(categories)}"
        )
    if not candidate.place:
        raise ValueError(
            f"Veranstaltungsort für Event {candidate.eventid} fehlt in der Race-ID-Tabelle."
        )

    category = categories.pop()
    folder_name = (
        f"{newest_date.isoformat()} {candidate.place.strip()} "
        f"{category}{candidate.gender}"
    )
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", folder_name).strip(" .")


def download_candidate(
    candidate: Candidate,
    output_directory: Path = PDF_DIRECTORY,
    progress=lambda message: None,
) -> list[Path]:
    progress(
        f"Lade FIS-Event {candidate.eventid} ({candidate.season}) für "
        f"Codex {candidate.codex}, Gender {candidate.gender}."
    )
    archived_directory = archive_existing_files(output_directory, progress)
    if archived_directory is not None:
        progress(f"Vorherige PDFs archiviert in {archived_directory}")

    with _request(_detail_url(candidate.eventid, candidate.season)) as response:
        html = response.read().decode(
            response.headers.get_content_charset() or "utf-8",
            errors="replace",
        )

    competitions = competition_rows(html)
    target_competitions = [
        competition
        for competition in competitions
        if competition.gender == candidate.gender
    ]
    if not any(competition.codex == candidate.codex for competition in target_competitions):
        raise ValueError(
            f"Codex {candidate.codex} mit Gender {candidate.gender} wurde auf "
            f"FIS-Event {candidate.eventid} nicht gefunden."
        )

    folder_name = folder_name_for_group(candidate, competitions)
    attachments = {}
    for competition in target_competitions:
        for attachment in competition.attachments:
            attachments.setdefault(attachment.url, attachment)
    if not attachments:
        raise ValueError(
            f"Keine PDF-Dateien für Gender {candidate.gender} auf "
            f"FIS-Event {candidate.eventid} gefunden."
        )

    output_directory.mkdir(parents=True, exist_ok=True)
    downloaded = []
    total = len(attachments)
    attachment_items = list(attachments.values())
    for index, attachment in enumerate(attachment_items, start=1):
        progress(f"Lade PDF {index}/{total}: {attachment.label}")
        with _request(attachment.url) as response:
            content = response.read()
            filename = _pdf_filename(response.headers, attachment.url)
            content_type = response.headers.get("Content-Type", "").lower()
        if not content.startswith(b"%PDF") or (
            content_type and "application/pdf" not in content_type
        ):
            raise ValueError(
                f"Der FIS-Anhang ist kein gültiges PDF: {attachment.label or attachment.url}"
            )

        destination = _unique_path(output_directory, filename)
        with destination.open("xb") as pdf_file:
            pdf_file.write(content)
        downloaded.append(destination)
        progress(f"Gespeichert: {destination.name}")
        if index < total:
            time.sleep(REQUEST_DELAY_SECONDS)
    folder_name_file = output_directory / "myFolderName.txt"
    folder_name_file.write_text(folder_name, encoding="utf-8")
    progress(f"Ordnername gespeichert: {folder_name_file}")
    return downloaded


class PdfDownloadApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.events = queue.Queue()
        self.root.title("FIS Ranglisten-PDF Download")
        self.root.geometry("620x340")
        self.root.minsize(520, 280)

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Codex eingeben (zwei- bis vierstellig; führende Nullen werden ergänzt):",
        ).pack(anchor="w")

        self.codex_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.codex_var, width=18)
        entry.pack(anchor="w", pady=(6, 10))
        entry.bind("<Return>", lambda _event: self.start_search())
        self.download_button = ttk.Button(
            frame, text="PDFs herunterladen", command=self.start_search
        )
        self.download_button.pack(anchor="w")

        self.status = tk.StringVar(value=f"PDFs werden in {PDF_DIRECTORY} gespeichert.")
        ttk.Label(
            frame,
            textvariable=self.status,
            wraplength=570,
            justify="left",
        ).pack(anchor="w", pady=(16, 8))

        self.log = tk.Text(frame, height=9, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)
        entry.focus_set()
        self.root.after(100, self._poll_events)

    def _write_log(self, message: str):
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.status.set(message)

    def _start_worker(self, function, *args):
        self.download_button.configure(state="disabled")

        def progress(message):
            self.events.put(("progress", message))

        def run():
            try:
                result = function(*args, progress=progress)
            except Exception as error:
                self.events.put(("error", str(error)))
            else:
                if result is None:
                    self.events.put(("complete", None))
                else:
                    self.events.put(("result", result))

        threading.Thread(target=run, daemon=True).start()

    def start_search(self):
        try:
            codex = normalize_codex(self.codex_var.get())
        except ValueError as error:
            messagebox.showerror("Ungültiger Codex", str(error), parent=self.root)
            return

        self.codex_var.set(codex)
        self._write_log(f"Suche Codex {codex} ...")
        self._start_worker(self._search, codex)

    def _search(self, codex: str, progress):
        candidates = find_candidates(codex, progress=progress)
        if not candidates:
            raise ValueError(f"Codex {codex} wurde weder in Excel noch bei FIS gefunden.")
        return candidates

    def _choose_candidate(self, candidates: list[Candidate]):
        if len(candidates) == 1:
            self._start_worker(self._download, candidates[0])
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Wettkampf auswählen")
        dialog.transient(self.root)
        dialog.grab_set()
        ttk.Label(
            dialog,
            text="Der Codex kommt mehrfach vor. Bitte den gewünschten Wettkampf wählen:",
            padding=12,
            wraplength=460,
        ).pack(anchor="w")
        listbox = tk.Listbox(dialog, width=76, height=min(12, len(candidates)))
        listbox.pack(padx=12, fill="both", expand=True)
        for candidate in candidates:
            listbox.insert("end", candidate.description)
        listbox.selection_set(0)

        def choose(_event=None):
            selection = listbox.curselection()
            if not selection:
                return
            candidate = candidates[selection[0]]
            dialog.destroy()
            self._start_worker(self._download, candidate)

        listbox.bind("<Double-Button-1>", choose)
        ttk.Button(dialog, text="Auswählen", command=choose).pack(pady=12)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

    def _download(self, candidate: Candidate, progress):
        downloaded = download_candidate(candidate, progress=progress)
        progress(f"Fertig: {len(downloaded)} PDFs heruntergeladen.")

    def _poll_events(self):
        while True:
            try:
                event, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event == "progress":
                self._write_log(payload)
            elif event == "result":
                self.download_button.configure(state="normal")
                self._choose_candidate(payload)
            elif event == "error":
                self.download_button.configure(state="normal")
                self._write_log(f"Fehler: {payload}")
                messagebox.showerror("Download fehlgeschlagen", payload, parent=self.root)
            elif event == "complete":
                self.download_button.configure(state="normal")
        self.root.after(100, self._poll_events)


def main():
    root = tk.Tk()
    PdfDownloadApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
