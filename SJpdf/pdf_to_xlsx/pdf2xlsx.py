"""Create an Excel overview of data blocks in FIS result PDFs."""

from __future__ import annotations

import argparse
from copy import copy
from datetime import date, datetime, time, timedelta
import re
import sys
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sj_paths import DATA_DIRECTORY  # noqa: E402

if __package__:
    from .pdf2xlsx_Daten import DATA_ROWS, PdfParseError, extract_data
    from .pdf2xlsx_Listen import (
        SJ_LISTEN_FILE, ListenError, SjListen, hill_name, proposed_jumper_id, read_sj_listen,
    )
    from .pdf2xlsx_Notes import LINK_FONT, Note, read_pdf_links, write_notes_sheet
    from .pdf2xlsx_Ranglisten import (
        ROUND_KEY,
        SHEET_NAME as RANGLISTEN_SHEET,
        PdfResults,
        RanglistenError,
        build_ranglisten,
        competition_file_name,
        write_ranglisten_sheet,
    )
    from .pdf2xlsx_Resultate import (
        TEAM_KEY,
        ResultsParseError,
        extract_results,
        write_error_sheet,
        write_results_sheet,
    )
    from .pdf2xlsx_Startlisten import (
        StartListError,
        extract_start_list,
        start_list_paths,
        write_start_list_sheet,
    )
else:
    from pdf2xlsx_Daten import DATA_ROWS, PdfParseError, extract_data
    from pdf2xlsx_Listen import (
        SJ_LISTEN_FILE, ListenError, SjListen, hill_name, proposed_jumper_id, read_sj_listen,
    )
    from pdf2xlsx_Notes import LINK_FONT, Note, read_pdf_links, write_notes_sheet
    from pdf2xlsx_Ranglisten import (
        ROUND_KEY,
        SHEET_NAME as RANGLISTEN_SHEET,
        PdfResults,
        RanglistenError,
        build_ranglisten,
        competition_file_name,
        write_ranglisten_sheet,
    )
    from pdf2xlsx_Resultate import (
        TEAM_KEY,
        ResultsParseError,
        extract_results,
        write_error_sheet,
        write_results_sheet,
    )
    from pdf2xlsx_Startlisten import (
        StartListError,
        extract_start_list,
        start_list_paths,
        write_start_list_sheet,
    )


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PDF_DIRECTORY = DATA_DIRECTORY / "pdf"
OUTPUT_FILE = DATA_DIRECTORY / "latest event.xlsx"
PDF_NAME_PATTERN = re.compile(
    # RLT1, RLT2, ...: several official trainings; RL1 is read with its RL.
    r"^\d{4}JP\d{4}(?:RLT\d?|RLQ|RTRIA|RL)\.pdf$",
    re.IGNORECASE,
)
UNAVAILABLE_FIELD_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFFFC7CE",
)
SECTION_FILL = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
RANGLISTE_FONT_COLOR = "FF0000FF"
# Section rows of Daten: the rounds and the footer.
GRAY_ROWS = tuple(
    DATA_ROWS.index(key) + 2 for key in ("DG 1", "DG 2", "DG 3", "Footer Datum")
)
DGW_ROW = "Pts / Gate DGW"


def default_pdf_directory() -> Path:
    """Return data/pdf in the shared SJpdf data folder (see sj_paths.py)."""
    return PDF_DIRECTORY


def find_result_pdfs(pdf_directory: Path) -> list[Path]:
    """Return supported PDFs directly in the input folder, in stable order."""
    if not pdf_directory.exists():
        return []
    return sorted(
        (
            path for path in pdf_directory.iterdir()
            if path.is_file()
            and path.suffix.casefold() == ".pdf"
            and PDF_NAME_PATTERN.fullmatch(path.name)
        ),
        key=lambda path: path.name.casefold(),
    )


def _cleared_workbook(output_file: Path) -> Workbook:
    """Load the workbook (or create it) with only Daten and its first column."""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists():
        workbook = load_workbook(output_file)
    else:
        workbook = Workbook()

    worksheet = (
        workbook["Daten"]
        if "Daten" in workbook.sheetnames
        else workbook.create_sheet("Daten")
    )
    for sheet_name in workbook.sheetnames:
        if sheet_name != "Daten":
            del workbook[sheet_name]
    if worksheet.max_column > 1:
        worksheet.delete_cols(2, worksheet.max_column - 1)
    return workbook


def reset_workbook(output_file: Path) -> None:
    """Keep only the Daten sheet and its first column in the existing workbook."""
    workbook = _cleared_workbook(output_file)
    try:
        workbook.save(output_file)
    finally:
        workbook.close()


def _sort_key(values: dict[str, object]) -> tuple[bool, int, bool, int]:
    competition_date = values.get("Datum")
    date_missing = competition_date is None
    date_order = competition_date.toordinal() if hasattr(competition_date, "toordinal") else 0

    start_time = values.get("Start Time")
    time_match = re.fullmatch(r"(\d{1,2}):(\d{2})", str(start_time or ""))
    time_missing = time_match is None
    time_order = (
        int(time_match.group(1)) * 60 + int(time_match.group(2))
        if time_match
        else 0
    )
    return date_missing, date_order, time_missing, time_order


def _read_results(pdf_path: Path, workbook: Workbook, warnings: list[str]) -> list | None:
    """Add the result sheet of one PDF; return its rows or None on failure."""
    try:
        rows = extract_results(pdf_path)
    except ResultsParseError as error:
        warnings.append(str(error))
        write_error_sheet(workbook, pdf_path.stem, str(error))
        return None
    write_results_sheet(workbook, pdf_path.stem, rows)
    return rows


def _first_round_path(pdf_path: Path) -> Path | None:
    """Return the RL1 PDF (intermediate results) next to an RL PDF, if any."""
    if not re.search(r"\dRL$", pdf_path.stem, re.I):
        return None
    first_round_path = pdf_path.with_name(f"{pdf_path.stem}1{pdf_path.suffix}")
    return first_round_path if first_round_path.exists() else None


def _read_start_order(
    pdf_path: Path | None, workbook: Workbook, warnings: list[str]
) -> tuple[list[int] | None, bool]:
    """Add the start list sheet; return (bibs in start order, read ok)."""
    if pdf_path is None:
        return None, True
    try:
        rows = extract_start_list(pdf_path)
    except StartListError as error:
        warnings.append(str(error))
        write_error_sheet(workbook, pdf_path.stem, str(error))
        return None, False
    write_start_list_sheet(workbook, pdf_path.stem, rows)
    return [row["Bib"] for row in rows], True


def _document_paths(extracted: Sequence[tuple[Path, dict]]) -> list[Path]:
    """All PDFs in sheet order: start list SLR1 before RL1, SLRF before RL."""
    paths = []
    for pdf_path, _values in extracted:
        if re.search(r"\dRL1$", pdf_path.stem, re.I):
            start_list = start_list_paths(pdf_path.with_name(f"{pdf_path.stem[:-1]}.pdf"))[0]
        else:
            start_list = start_list_paths(pdf_path)[1]
        if start_list is not None:
            paths.append(start_list)
        paths.append(pdf_path)
    return paths


def _add_result_sheets(
    workbook: Workbook, extracted: Sequence[tuple[Path, dict]], warnings: list[str]
) -> list[PdfResults] | None:
    """Add one sheet per PDF (RL1 and start lists included); None if any PDF failed.

    Sheet order of a competition: SLR1, RL1, SLRF, RL.
    """
    pdf_results: list[PdfResults] | None = []
    for pdf_path, _values in extracted:
        if re.search(r"\dRL1$", pdf_path.stem, re.I):
            continue  # read together with its RL PDF
        first_list, final_list = start_list_paths(pdf_path)
        first_round_stem = first_round_rows = None
        first_round_path = _first_round_path(pdf_path)
        start_order_1, ok_1 = _read_start_order(first_list, workbook, warnings)
        if first_round_path is not None:
            first_round_stem = first_round_path.stem
            first_round_rows = _read_results(first_round_path, workbook, warnings)
            if first_round_rows is None:
                pdf_results = None
        start_order_2, ok_2 = _read_start_order(final_list, workbook, warnings)
        rows = _read_results(pdf_path, workbook, warnings)
        if rows is not None and any(row.get(TEAM_KEY) for row in rows):
            # Team events: the later groups of a start list have no start order
            # yet; without it the rounds are ordered by the team standings.
            ok_1 = ok_2 = True
        if rows is None or not (ok_1 and ok_2):
            pdf_results = None
        if pdf_results is not None:
            pdf_results.append(
                PdfResults(
                    pdf_path.stem, rows, first_round_stem, first_round_rows,
                    start_order_1, start_order_2,
                )
            )
    return pdf_results


def _read_listen(listen_file: Path | None, notes: list[Note]) -> SjListen | None:
    """SJ-Listen.xlsm or None (with a warning) when it cannot be read."""
    if listen_file is None or isinstance(listen_file, SjListen):
        return listen_file
    try:
        return read_sj_listen(listen_file)
    except ListenError as error:
        notes.append(Note(
            "Warnung",
            f"SJ-Listen: {error} JumperID und {DGW_ROW} wurden nicht ermittelt.",
        ))
        return None


def _add_hill_values(
    extracted: Sequence[tuple[Path, dict]], listen: SjListen | None, notes: list[Note]
) -> None:
    """Set 'Pts / Gate DGW' of each PDF from Liste_Schanzen_ID."""
    if listen is None:
        return
    found: dict[str, object] = {}
    missing: dict[str, None] = {}
    for pdf_path, values in extracted:
        hill = hill_name(values.get("Place"), values.get("Hill Size"))
        if hill is None:
            missing.setdefault(f"von {pdf_path.stem} (Place oder Hill Size fehlt)", None)
        elif not listen.has_hill(hill):
            missing.setdefault(hill, None)
        elif not isinstance(listen.dgw(hill), (int, float)):
            notes.append(Note(
                "Warnung",
                f"Daten: SchanzenID {hill} hat in Liste_Schanzen_ID keinen Wert "
                f"'Pts / Gate für DGW'.",
            ))
        else:
            values[DGW_ROW] = listen.dgw(hill)
            found[hill] = values[DGW_ROW]
    for hill in missing:
        notes.append(Note(
            "Warnung", f"Daten: SchanzenID {hill} fehlt in Liste_Schanzen_ID (SJ-Listen.xlsm)."
        ))
    for hill, value in found.items():
        notes.append(Note(
            "Info",
            f"Daten: SchanzenID {hill} in Liste_Schanzen_ID gefunden, {DGW_ROW} {value:g}.",
        ))


def _add_jumper_ids(rows: list[dict], listen: SjListen | None, notes: list[Note]) -> None:
    """Set JumperID of each Ranglisten row: built by the rule from NSA and
    name, checked against Liste_Springer_ID (see SjListen.resolve)."""
    if listen is None:
        return
    missing: dict[tuple, str] = {}
    for row in rows:
        row["JumperID"], problem = listen.resolve(row.get("Name"), row.get("NSA"))
        if problem is not None:
            missing.setdefault((row.get("Name"), row.get("NSA")), problem)
    for (name, nsa), problem in missing.items():
        notes.append(Note("Warnung", f"Ranglisten: {problem}: {name} ({nsa})."))
    if not missing:
        notes.append(Note("Info", "Ranglisten: Alle JumperID in SJ-Listen.xlsm vorhanden."))


def _clock(text: object) -> time | None:
    match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", str(text or "").strip())
    if not match:
        return None
    hours, minutes, seconds = (int(part or 0) for part in match.groups())
    return time(hours, minutes, seconds) if hours < 24 and minutes < 60 and seconds < 60 else None


def _add_jump_times(
    rows: list[dict], values_by_stem: dict[str, dict], warnings: list[str]
) -> None:
    """Date and time of each jump.

    The date of a round is the Datum of its PDF. The first jumper of a round
    gets the start time, the last one the end time of the round (DG n in
    Daten); the jumpers in between are spaced evenly. Rows without a jump
    (DNS, Not Qualified, ...) get the date but no time.
    """
    blocks: dict[str, list[dict]] = {}
    for row in rows:
        blocks.setdefault(row["JumpID"], []).append(row)
    for jump_id, block in blocks.items():
        source = block[0]["PDF"]
        number = block[0][ROUND_KEY]
        values = values_by_stem.get(source, {})
        day = values.get("Datum")
        for row in block:
            row["Date"] = day if isinstance(day, date) else None
        jumpers = [row for row in block if row.get("Distance") is not None]
        if not jumpers:
            continue
        start, end = _clock(values.get(f"Start Time {number}")), _clock(values.get(f"End Time {number}"))
        if start is None and end is None and number == 1 and values.get("DG 1") is None:
            # No weather block (e.g. qualification): the times of the header.
            start, end = _clock(values.get("Start Time")), _clock(values.get("Finish Time"))
        if not isinstance(day, date) or start is None or end is None:
            warnings.append(
                f"{source}: date, start or end time of round {number} missing in Daten; "
                f"no jump times for {jump_id}."
            )
            continue
        first = datetime.combine(day, start)
        duration = datetime.combine(day, end) - first
        if duration < timedelta(0):
            warnings.append(
                f"{source}: end time of round {number} is before its start time; "
                f"no jump times for {jump_id}."
            )
            continue
        step = duration / (len(jumpers) - 1) if len(jumpers) > 1 else timedelta(0)
        for index, row in enumerate(jumpers):
            moment = first + step * index
            row["Time"] = (moment + timedelta(microseconds=500_000)).replace(microsecond=0).time()


def _add_ranglisten(
    workbook: Workbook,
    pdf_results: list[PdfResults] | None,
    competition: str | None,
    warnings: list[str],
    listen: SjListen | None = None,
    notes: list[Note] | None = None,
    values_by_stem: dict[str, dict] | None = None,
) -> None:
    message = None
    if pdf_results is None:
        message = "not created because the results of at least one PDF could not be read."
    elif competition is None:
        message = "not created because the competition name (first part of JumpID) is unknown."
    if message is not None:
        warnings.append(f"{RANGLISTEN_SHEET}: {message}")
        write_error_sheet(workbook, RANGLISTEN_SHEET, message)
        workbook.move_sheet(RANGLISTEN_SHEET, 1 - workbook.sheetnames.index(RANGLISTEN_SHEET))
        return
    try:
        rows, ranglisten_warnings = build_ranglisten(pdf_results, competition)
    except (RanglistenError, ResultsParseError) as error:
        warnings.append(f"{RANGLISTEN_SHEET}: {error}")
        write_error_sheet(workbook, RANGLISTEN_SHEET, str(error))
        workbook.move_sheet(RANGLISTEN_SHEET, 1 - workbook.sheetnames.index(RANGLISTEN_SHEET))
        return
    warnings.extend(ranglisten_warnings)
    _add_jumper_ids(rows, listen, notes if notes is not None else [])
    _add_jump_times(rows, values_by_stem or {}, warnings)
    write_ranglisten_sheet(workbook, rows)


_DEFAULT = object()


def create_workbook(
    pdf_paths: Sequence[Path],
    output_file: Path,
    listen_file=_DEFAULT,
    competition: str | None = None,
    save_copy: bool = True,
    place: str | None = None,
) -> list[str]:
    """Extract each PDF once and save the selected fields in separate columns.

    Every PDF also gets its own sheet (named like the PDF) with the result
    table; for a competition (RL) the intermediate results after the first
    round (RL1) and the start lists (SLR1, SLRF) are read as well. The sheet Ranglisten combines all rounds
    in jump order. When results cannot be read without doubt, the sheet only
    shows the reason and the reason is returned as a warning.

    Each file is saved exactly once, so a sync client (Nextcloud) sees a
    single change. Besides output_file, a copy named after the competition
    (e.g. '2026-09-13 FCM Einsiedeln HS117.xlsx') is saved in the same
    folder. When there are no PDFs or a PDF cannot be read, the cleared
    workbook is saved so no outdated results remain.

    Jumpers (JumperID in Ranglisten) and hills (Pts / Gate DGW in Daten) are
    looked up in listen_file (default SJ-Listen.xlsm; None: no lookup; an
    SjListen already read is used as it is).

    competition replaces the name read from the PDFs (first part of JumpID,
    e.g. from the PDF archive folder); save_copy=False saves output_file only.
    place (FIS calendar) replaces a Place of a PDF without country code
    ('Engelberg (SUI)'), as printed by local organisers.
    """
    if listen_file is _DEFAULT:
        listen_file = SJ_LISTEN_FILE
    created = datetime.now().replace(microsecond=0)
    workbook = _cleared_workbook(output_file)
    worksheet = workbook["Daten"]
    try:
        if not pdf_paths:
            message = "No matching RLT, RLQ, RTRIA, or RL PDFs found."
            write_notes_sheet(workbook, created, [], [Note("Fehler", message)])
            workbook.save(output_file)
            raise FileNotFoundError(message)

        try:
            extracted = [
                (pdf_path, extract_data(pdf_path))
                for pdf_path in pdf_paths
            ]
            extracted.sort(key=lambda item: _sort_key(item[1]))
            # The RL1 column stands directly before its RL column.
            with_first_rounds = []
            for pdf_path, values in extracted:
                first_round_path = _first_round_path(pdf_path)
                if first_round_path is not None:
                    with_first_rounds.append((first_round_path, extract_data(first_round_path)))
                with_first_rounds.append((pdf_path, values))
            extracted = with_first_rounds
            if place:
                for _path, values in extracted:
                    if not re.search(r"\([A-Z]{3}\)", str(values.get("Place") or "")):
                        values["Place"] = place
        except PdfParseError as error:
            pdfs = [(path, None) for path in pdf_paths]
            write_notes_sheet(workbook, created, pdfs, [Note("Fehler", str(error))])
            workbook.save(output_file)
            raise

        warnings: list[str] = []
        listen_notes: list[Note] = []  # SJ-Listen: hill first, then jumpers
        listen = _read_listen(listen_file, listen_notes)
        _add_hill_values(extracted, listen, listen_notes)

        worksheet["A1"] = "pdf"

        for column, (pdf_path, values) in enumerate(extracted, start=2):
            worksheet.cell(row=1, column=column, value=pdf_path.stem)
            for row, key in enumerate(DATA_ROWS, start=2):
                cell = worksheet.cell(row=row, column=1, value=key)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(fill_type=None)  # old section rows of column A
                value = values.get(key)
                value_cell = worksheet.cell(row=row, column=column, value=value)
                if value is None:
                    value_cell.fill = UNAVAILABLE_FIELD_FILL
                if key == "Datum" and value is not None:
                    value_cell.number_format = "yyyy-mm-dd"

        for row in range(1, worksheet.max_row + 1):
            for column in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=column).alignment = Alignment(
                    horizontal="left"
                )

        notes: list[Note] = []
        links: dict[str, str] = {}
        documents = _document_paths(extracted)
        for folder in sorted({path.parent for path in documents}):
            folder_links = read_pdf_links(folder)
            if folder_links is None:
                notes.append(Note(
                    "Info",
                    f"Keine FIS-Links: {folder / 'pdf_links.csv'} fehlt "
                    "(PDFs vor dem Speichern der Links heruntergeladen).",
                ))
            else:
                links.update(folder_links)
                missing = [
                    path.stem for path in documents
                    if path.parent == folder and path.name not in folder_links
                ]
                if missing:
                    notes.append(Note("Info", f"Kein FIS-Link in pdf_links.csv für {', '.join(missing)}."))
        pdf_links = [(path, links.get(path.name)) for path in documents]
        for column, (path, _values) in enumerate(extracted, start=2):
            url = links.get(path.name)
            if url:
                cell = worksheet.cell(row=1, column=column)
                cell.hyperlink = url
                cell.font = LINK_FONT

        rangliste_row = DATA_ROWS.index("Rangliste") + 2
        for cell in worksheet[rangliste_row]:
            font = copy(cell.font)
            font.color = RANGLISTE_FONT_COLOR
            cell.font = font

        for row in GRAY_ROWS:
            for cell in worksheet[row]:
                cell.fill = SECTION_FILL

        if competition is not None:
            name = competition
        else:
            try:
                name = competition_file_name([values for _path, values in extracted])
            except RanglistenError as error:
                name = None
                warnings.append(f"No competition copy saved: {error}")
        pdf_results = _add_result_sheets(workbook, extracted, warnings)
        _add_ranglisten(
            workbook, pdf_results, name, warnings, listen, listen_notes,
            {path.stem: values for path, values in extracted},
        )

        worksheet.freeze_panes = "B12"
        worksheet.column_dimensions["A"].width = 20
        for column in range(2, len(extracted) + 2):
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=column).column_letter
            ].width = 32

        notes = [Note("Warnung", warning) for warning in warnings] + listen_notes + notes
        warnings = warnings + [note.message for note in listen_notes if note.kind == "Warnung"]
        copy_name = f"{name}.xlsx" if name is not None and save_copy else None
        write_notes_sheet(workbook, created, pdf_links, notes, copy_name)
        workbook.save(output_file)
        if copy_name is not None:
            workbook.save(output_file.with_name(copy_name))
        return warnings
    finally:
        workbook.close()


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract data blocks from FIS result PDFs."
    )
    parser.add_argument(
        "--pdf-dir",
        type=Path,
        default=default_pdf_directory(),
        help=f"Folder containing downloaded PDFs (default: {PDF_DIRECTORY})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help=f"Output workbook (default: {OUTPUT_FILE})",
    )
    args = parser.parse_args(argv)

    try:
        pdf_paths = find_result_pdfs(args.pdf_dir)
        warnings = create_workbook(pdf_paths, args.output)
    except PermissionError as error:
        print(
            f"PDF to Excel failed: Cannot write '{args.output}'. "
            "Close the workbook in Excel if it is open, then try again. "
            f"Also check that you have permission to write to the folder. ({error})",
            file=sys.stderr,
        )
        return 1
    except (FileNotFoundError, OSError, PdfParseError) as error:
        print(f"PDF to Excel failed: {error}", file=sys.stderr)
        return 1

    for warning in warnings:
        print(f"Warning: {warning}", file=sys.stderr)
    print(f"Processed {len(pdf_paths)} PDF(s): {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
