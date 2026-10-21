"""Utilities for downloading FIS result PDFs."""

from __future__ import annotations

import ctypes
import os
import threading
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path


FIS_PDF_URL_TEMPLATE = (
    "https://medias1.fis-ski.com/pdf/"
    "{season}/JP/{codex}/{season}JP{codex}{document_type}.pdf"
)
FIS_PDF_DOCUMENT_TYPES = ("RL", "SL", "TNG", "TRIAL", "RL1")
FIS_PDF_DOWNLOAD_WORKERS = 4
FIS_PDF_DOWNLOAD_TIMEOUT = 30
FIS_PDF_POPUP_TITLE = "FIS PDF-Download"


def fis_season(event_date: date) -> int:
    """Return the FIS season year; June already belongs to the next season."""
    return event_date.year + (1 if event_date.month >= 6 else 0)


def build_fis_pdf_url(season: int, codex: str, document_type: str) -> str:
    """Build the official FIS PDF URL for one document."""
    return FIS_PDF_URL_TEMPLATE.format(
        season=season,
        codex=codex,
        document_type=document_type,
    )


def download_fis_pdf(url: str, dest_path: Path) -> tuple[bool, str | None]:
    """
    Downloads a PDF from *url* and saves it to *dest_path*.

    Returns:
        (True, None)              on success.
        (False, error_message)    on failure.
    """
    try:
        dest_path = Path(dest_path)
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        }
        req = urllib.request.Request(url, headers=headers)

        with urllib.request.urlopen(req, timeout=FIS_PDF_DOWNLOAD_TIMEOUT) as response:
            pdf_bytes = response.read()

        # FIS may occasionally return an HTML error page with a successful
        # status code. Do not store that page with a .pdf extension.
        if not pdf_bytes.lstrip().startswith(b"%PDF"):
            return False, "Antwort ist keine PDF-Datei."

        # Write atomically so an interrupted request never leaves a file that
        # will be mistaken for a completed download on the next export.
        temporary_path = dest_path.with_suffix(dest_path.suffix + ".part")
        temporary_path.write_bytes(pdf_bytes)
        os.replace(temporary_path, dest_path)
        return True, None

    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.reason}"
    except urllib.error.URLError as e:
        return False, f"URL-Fehler: {e.reason}"
    except Exception as e:
        return False, str(e)


def _show_download_popup(downloaded: list[str], destination: Path) -> None:
    """Show a short native Windows message after the background work ends."""
    if downloaded:
        message = "Neu gespeichert:\n" + "\n".join(downloaded)
    else:
        message = "Keine neuen PDF-Dateien heruntergeladen."
    message += f"\n\nOrdner: {destination}"
    try:
        ctypes.windll.user32.MessageBoxW(0, message, FIS_PDF_POPUP_TITLE, 0x40)
    except Exception:
        # A notification failure must never affect the already saved workbook.
        print(message)


def _download_pdf_batch(
    season_codex_pairs: tuple[tuple[int, str], ...], destination: Path
) -> None:
    """Download all missing documents and notify the user on completion."""
    downloaded: list[str] = []
    try:
        destination.mkdir(parents=True, exist_ok=True)
        requests: list[tuple[str, Path]] = []
        for season, codex in season_codex_pairs:
            for document_type in FIS_PDF_DOCUMENT_TYPES:
                target = destination / f"{season}-{codex}-{document_type}.pdf"
                if target.exists():
                    continue
                requests.append(
                    (build_fis_pdf_url(season, codex, document_type), target)
                )

        def download_one(item: tuple[str, Path]) -> str | None:
            url, target = item
            success, _error = download_fis_pdf(url, target)
            return target.name if success else None

        if requests:
            with ThreadPoolExecutor(max_workers=FIS_PDF_DOWNLOAD_WORKERS) as executor:
                downloaded = [
                    name for name in executor.map(download_one, requests) if name
                ]
    except Exception as error:
        print(f"Fehler beim PDF-Hintergrunddownload: {error}")
    finally:
        _show_download_popup(sorted(downloaded), destination)


def start_fis_pdf_downloads(
    season_codex_pairs, destination: str | Path
) -> threading.Thread | None:
    """Start non-blocking PDF downloads in the configured PDF directory."""
    pairs = tuple(sorted(set(season_codex_pairs)))
    if not pairs:
        return None
    destination = Path(destination).resolve()
    worker = threading.Thread(
        target=_download_pdf_batch,
        args=(pairs, destination),
        name="fis-pdf-download",
        daemon=True,
    )
    worker.start()
    return worker
