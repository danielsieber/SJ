import csv
import queue
import re
import shutil
import sys
import threading
import time
import tkinter as tk
import warnings
from dataclasses import dataclass
from datetime import date, datetime
from email.message import Message
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.parse import parse_qs, urlparse

from openpyxl import load_workbook

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
    parse_document,
    read_text,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
# Sheet Event_ID is written by raceid_crawler/fis_raceid_crawler.py.
RACEID_FILE = DATA_DIRECTORY / "SJ-Listen.xlsm"
EVENT_SHEET = "Event_ID"
PDF_DIRECTORY = DATA_DIRECTORY / "pdf"
# Original FIS link of every downloaded PDF (read by pdf_to_xlsx).
LINKS_FILENAME = "pdf_links.csv"
CODEX_PATTERN = re.compile(r"^\s*(\d{4})(?:\b|$)")


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


def _load_event_workbook(workbook_path: Path):
    with warnings.catch_warnings():
        # openpyxl warns about the print area of other sheets in SJ-Listen.xlsm.
        warnings.simplefilter("ignore", UserWarning)
        return load_workbook(workbook_path, read_only=True, data_only=True)


def _event_sheet_names(workbook) -> list[str]:
    """Event_ID (SJ-Listen.xlsm), else 'all' or men/women/mixed (old raceid.xlsx)."""
    for name in (EVENT_SHEET, "all"):
        if name in workbook.sheetnames:
            return [name]
    names = [name for name in ("men", "women", "mixed") if name in workbook.sheetnames]
    if not names:
        raise ValueError(f"Sheet {EVENT_SHEET} fehlt in {workbook.path or 'der Excel-Datei'}.")
    return names


def read_workbook_candidates(workbook_path: Path, codex: str) -> list[Candidate]:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel-Datei mit dem Sheet Event_ID nicht gefunden: {workbook_path}")

    workbook = _load_event_workbook(workbook_path)
    try:
        sheet_names = _event_sheet_names(workbook)
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
                raw_codex = _cell_text(row[columns["codex"]])
                if not raw_codex or raw_codex.zfill(4) != codex:
                    continue
                row_codex = raw_codex.zfill(4)
                season_text = _cell_text(row[columns["season"]])
                eventid = _cell_text(row[columns["eventid"]])
                if not season_text.isdigit() or not eventid:
                    continue
                season = int(season_text)
                gender = _cell_text(row[columns["gender"]]).upper()
                key = (season, eventid, row_codex, gender)
                event_date = row[columns["date"]] if "date" in columns else ""
                if isinstance(event_date, datetime):
                    event_date = event_date.date()
                if event_date is None:
                    event_date = ""
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
            key=lambda candidate: (
                candidate.season,
                _date_sort_text(candidate.event_date),
                candidate.eventid,
            ),
        )
    finally:
        workbook.close()


def _date_sort_text(value) -> str:
    """Sort dates and ISO date strings together; missing dates sort first."""
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value or "")


def _request(url: str):
    return open_url(url, timeout=45)


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
            calendar_html = read_text(response)
        events = _calendar_events(calendar_html)
        for index, (eventid, place) in enumerate(events, start=1):
            progress(
                f"Suche Codex {codex} auf FIS-Eventseiten ({index}/{len(events)}): "
                f"{place or eventid}"
            )
            with _request(_detail_url(eventid, season)) as response:
                event_html = read_text(response)
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

    workbook = _load_event_workbook(workbook_path)
    try:
        seasons = set()
        for sheet_name in _event_sheet_names(workbook):
            rows = workbook[sheet_name].iter_rows(values_only=True)
            headers = next(rows, None)
            if headers is None or "season" not in headers:
                continue
            season_column = headers.index("season")
            for row in rows:
                season_text = _cell_text(row[season_column])
                if season_text.isdigit():
                    seasons.add(int(season_text))
    finally:
        workbook.close()
    if not seasons:
        seasons = {current_season()}
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
            f"Veranstaltungsort für Event {candidate.eventid} fehlt im Sheet Event_ID."
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
        html = read_text(response)

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
    # Write the folder name before the first PDF: if a download fails midway,
    # the next run can still archive the partial files under the right name.
    folder_name_file = output_directory / "myFolderName.txt"
    folder_name_file.write_text(folder_name, encoding="utf-8")
    progress(f"Ordnername gespeichert: {folder_name_file}")

    links_file = output_directory / LINKS_FILENAME
    with links_file.open("w", encoding="utf-8", newline="") as handle:
        csv.writer(handle).writerow(["pdf", "url"])

    downloaded = []
    total = len(attachments)
    attachment_items = list(attachments.values())
    for index, attachment in enumerate(attachment_items, start=1):
        progress(f"Lade PDF {index}/{total}: {attachment.label}")
        with _request(attachment.url) as response:
            content = response.read()
            filename = _pdf_filename(response.headers, attachment.url)
        # FIS sometimes serves PDFs as application/octet-stream, so the file
        # signature decides rather than the Content-Type header.
        if not content.startswith(b"%PDF"):
            raise ValueError(
                f"Der FIS-Anhang ist kein gültiges PDF: {attachment.label or attachment.url}"
            )

        destination = _unique_path(output_directory, filename)
        with destination.open("xb") as pdf_file:
            pdf_file.write(content)
        with links_file.open("a", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerow([destination.name, attachment.url])
        downloaded.append(destination)
        progress(f"Gespeichert: {destination.name}")
        if index < total:
            time.sleep(REQUEST_DELAY_SECONDS)
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
