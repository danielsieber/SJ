"""Create an Excel overview of data blocks in FIS result PDFs."""

from __future__ import annotations

import argparse
from copy import copy
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
else:
    from pdf2xlsx_Daten import DATA_ROWS, PdfParseError, extract_data


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PDF_DIRECTORY = DATA_DIRECTORY / "pdf"
OUTPUT_FILE = DATA_DIRECTORY / "pdf_to_xlsx.xlsx"
PDF_NAME_PATTERN = re.compile(
    r"^\d{4}JP\d{4}(?:RLT|RLQ|RTRIA|RL)\.pdf$",
    re.IGNORECASE,
)
UNAVAILABLE_FIELD_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFFFC7CE",
)
SECTION_FILL = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
RANGLISTE_FONT_COLOR = "FF0000FF"
GRAY_ROWS = (19, 29, 39, 49)


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


def create_workbook(pdf_paths: Sequence[Path], output_file: Path) -> None:
    """Extract each PDF once and save the selected fields in separate columns.

    The workbook is saved exactly once, so a sync client (Nextcloud) sees a
    single change. When there are no PDFs or a PDF cannot be read, the
    cleared workbook is saved so no outdated results remain.
    """
    workbook = _cleared_workbook(output_file)
    worksheet = workbook["Daten"]
    try:
        if not pdf_paths:
            workbook.save(output_file)
            raise FileNotFoundError("No matching RLT, RLQ, RTRIA, or RL PDFs found.")

        try:
            extracted = [
                (pdf_path, extract_data(pdf_path))
                for pdf_path in pdf_paths
            ]
        except PdfParseError:
            workbook.save(output_file)
            raise

        worksheet["A1"] = "pdf"
        extracted.sort(key=lambda item: _sort_key(item[1]))

        for column, (pdf_path, values) in enumerate(extracted, start=2):
            worksheet.cell(row=1, column=column, value=pdf_path.stem)
            for row, key in enumerate(DATA_ROWS, start=2):
                cell = worksheet.cell(row=row, column=1, value=key)
                cell.font = Font(bold=True)
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

        rangliste_row = DATA_ROWS.index("Rangliste") + 2
        for cell in worksheet[rangliste_row]:
            font = copy(cell.font)
            font.color = RANGLISTE_FONT_COLOR
            cell.font = font

        for row in GRAY_ROWS:
            for cell in worksheet[row]:
                cell.fill = SECTION_FILL

        worksheet.freeze_panes = "B12"
        worksheet.column_dimensions["A"].width = 20
        for column in range(2, len(pdf_paths) + 2):
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=column).column_letter
            ].width = 32

        workbook.save(output_file)
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
        create_workbook(pdf_paths, args.output)
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

    print(f"Processed {len(pdf_paths)} PDF(s): {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
