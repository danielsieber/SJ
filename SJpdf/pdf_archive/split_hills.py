"""Split an event folder of the PDF archive by hill size.

An event (EventNr) can hold competitions on two hills, e.g. the Universiade
2017 in Erzurum (HS109 and HS140) in

    2017-02-05 COCM Erzurum - Kiremitliktepe LH

All PDFs of one codex belong to one hill. The hill size of a codex is read
from its result PDFs (RL, RL1, RLQ, RLT, RTRIA; else from its other PDFs).
When the codexes of a folder have more than one hill size, each hill size
gets its own folder, named like the event folder plus the hill size:

    2017-02-05 COCM Erzurum - Kiremitliktepe LH HS109
    2017-02-05 COCM Erzurum - Kiremitliktepe LH HS140

The PDFs of each codex and their rows of pdf_links.csv are moved there; the
workbooks of the event folder are deleted (archive_xlsx writes them again)
and the empty event folder is removed. PDFs downloaded later into the event
folder are moved to the folder of their codex (or of their hill size).
A codex without a readable hill size goes to the folder with most codexes.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path
from typing import Callable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
from pdf_to_xlsx.pdf2xlsx_Daten import PdfParseError, extract_data  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
import pdf_links  # noqa: E402

WORKBOOK_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2} .+\.xlsx$")  # written by archive_xlsx
CODEX_PATTERN = re.compile(r"^(\d{4}JP\d{4})", re.I)
RESULT_PATTERN = re.compile(r"^\d{4}JP\d{4}(?:RL\d?|RLQ|RLT\d*|RTRIA)$", re.I)
SPLIT_SUFFIX = re.compile(r" HS(\d+)$")


def split_folders(folder: Path) -> list[Path]:
    """The hill size folders of an event folder ('<event folder> HS109', ...)."""
    if not folder.parent.is_dir():
        return []
    return sorted(
        path for path in folder.parent.iterdir()
        if path.is_dir() and path.name.startswith(f"{folder.name} HS")
        and SPLIT_SUFFIX.search(path.name)
        and path.name[: SPLIT_SUFFIX.search(path.name).start()] == folder.name
    )


def _read_links(folder: Path) -> list[dict[str, str]]:
    return pdf_links.read_rows(folder)


def _write_links(folder: Path, rows: list[dict[str, str]]) -> None:
    pdf_links.write_rows(folder, rows)  # pdf_links.csv and pdf_links.xlsx


def codex_of(path: Path) -> str | None:
    match = CODEX_PATTERN.match(path.stem)
    return match.group(1).upper() if match else None


def codex_hill_sizes(pdfs: list[Path]) -> dict[str, int | None]:
    """{codex: hill size} (most frequent of its result PDFs, else of its other PDFs)."""
    by_codex: dict[str, list[Path]] = {}
    for pdf in pdfs:
        codex = codex_of(pdf)
        if codex is not None:
            by_codex.setdefault(codex, []).append(pdf)
    sizes: dict[str, int | None] = {}
    for codex, paths in sorted(by_codex.items()):
        results = [path for path in paths if RESULT_PATTERN.match(path.stem)]
        sizes[codex] = None
        for group in (results, [path for path in paths if path not in results]):
            found = Counter()
            for path in group:
                try:
                    size = extract_data(path).get("Hill Size")
                except (PdfParseError, OSError):
                    continue
                if size is not None:
                    found[int(size)] += 1
            if found:
                sizes[codex] = found.most_common(1)[0][0]
                break
    return sizes


def _finish(
    folder: Path, targets: set[Path], links: dict[str, dict[str, str]],
    moved: dict[Path, list[dict[str, str]]],
) -> None:
    """Write the moved link rows, clean up the event folder and remove it when empty."""
    # Links of PDFs moved before (interrupted run) go to the folder of their PDF.
    for name in list(links):
        if not (folder / name).exists():
            target = next((t for t in sorted(targets) if (t / name).exists()), None)
            if target is not None:
                moved.setdefault(target, []).append(links.pop(name))
    for target, rows in moved.items():
        old = {row["pdf"]: row for row in _read_links(target)}
        old.update({row["pdf"]: row for row in rows})
        _write_links(target, list(old.values()))
    _write_links(folder, list(links.values()))
    for workbook in folder.glob("*.xlsx"):
        if WORKBOOK_PATTERN.match(workbook.name):
            workbook.unlink(missing_ok=True)  # rewritten per hill by archive_xlsx
    if not any(folder.iterdir()):
        folder.rmdir()


def split_by_hill_size(folder: Path, log: Callable[[str], None] = print) -> list[Path]:
    """Split the event folder when its codexes have different hill sizes.

    Return the folders that now hold the event's PDFs (the event folder
    itself when nothing was split)."""
    existing = split_folders(folder)
    pdfs = sorted(folder.glob("*.pdf")) if folder.is_dir() else []
    if not pdfs:
        if existing and folder.is_dir():
            _finish(folder, set(existing), {row["pdf"]: row for row in _read_links(folder)}, {})
        return existing or ([folder] if folder.is_dir() else [])
    sizes = codex_hill_sizes(pdfs)
    known = {size for size in sizes.values() if size is not None}
    if not existing and len(known) <= 1:
        return [folder]

    # Codexes already split keep their folder; a codex has one hill size.
    target_of_codex: dict[str, Path] = {}
    for split in existing:
        for pdf in split.glob("*.pdf"):
            codex = codex_of(pdf)
            if codex is not None:
                target_of_codex.setdefault(codex, split)
    targets: dict[int, Path] = {
        int(SPLIT_SUFFIX.search(split.name).group(1)): split for split in existing
    }
    for size in sorted(known):
        targets.setdefault(size, folder.parent / f"{folder.name} HS{size}")
    counts = Counter(size for size in sizes.values() if size is not None)
    for split in existing:
        size = int(SPLIT_SUFFIX.search(split.name).group(1))
        counts[size] += len({codex_of(pdf) for pdf in split.glob("*.pdf")})
    largest = max(targets, key=lambda size: counts.get(size, 0))

    links = {row["pdf"]: row for row in _read_links(folder)}
    moved: dict[Path, list[dict[str, str]]] = {}
    for pdf in pdfs:
        codex = codex_of(pdf)
        target = target_of_codex.get(codex) if codex else None
        if target is None:
            size = sizes.get(codex) if codex else None
            target = targets[size if size is not None else largest]
        target.mkdir(parents=True, exist_ok=True)
        destination = target / pdf.name
        try:
            if destination.exists():
                destination.unlink()  # same PDF downloaded again
            pdf.replace(destination)
        except PermissionError as error:  # PDF open in another program: next run
            log(f"  {folder.name}: {pdf.name} nicht verschoben ({error.strerror}).")
            continue
        if pdf.name in links:
            moved.setdefault(target, []).append(links.pop(pdf.name))
    _finish(folder, set(targets.values()), links, moved)
    result = sorted(set(targets.values()))
    log(f"  {folder.name}: nach Schanzengrösse aufgeteilt in "
        + ", ".join(path.name[len(folder.name) + 1:] for path in result))
    return result
