"""Create an Excel overview of the general-data block in FIS result PDFs."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

if __package__:
    from .pdf2xlsx_1Daten import DATA_ROWS, PdfParseError, extract_data
else:
    from pdf2xlsx_1Daten import DATA_ROWS, PdfParseError, extract_data


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PDF_DIRECTORY = PROJECT_ROOT / "data" / "pdf"
OUTPUT_FILE = PROJECT_ROOT / "data" / "pdf_to_xlsx.xlsx"
PDF_NAME_PATTERN = re.compile(
    r"^\d{4}JP\d{4}(?:RLT|RLQ|RTRIA|RL)\.pdf$",
    re.IGNORECASE,
)
UNAVAILABLE_FIELD_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFFFC7CE",
)


def default_pdf_directory(
    project_root: Path = PROJECT_ROOT,
    home_directory: Path = Path.home(),
) -> Path:
    """Prefer the synced project copy, then use PDFs in this checkout."""
    project_pdf_directory = project_root / "data" / "pdf"
    synced_pdf_directory = (
        home_directory / "CLOUD-DANIEL" / "SJA" / "SJpdf" / "data" / "pdf"
    )
    for directory in (synced_pdf_directory, project_pdf_directory):
        if directory.is_dir() and find_result_pdfs(directory):
            return directory
    return synced_pdf_directory


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


def reset_workbook(output_file: Path) -> None:
    """Keep only the Daten sheet and its first column in the existing workbook."""
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

    try:
        workbook.save(output_file)
    finally:
        workbook.close()


def create_workbook(pdf_paths: Sequence[Path], output_file: Path) -> None:
    """Extract each PDF once and save the selected fields in separate columns."""
    reset_workbook(output_file)
    workbook = load_workbook(output_file)
    worksheet = workbook["Daten"]
    if not pdf_paths:
        workbook.close()
        raise FileNotFoundError("No matching RLT, RLQ, RTRIA, or RL PDFs found.")

    try:
        worksheet["A1"] = "pdf"

        for column, pdf_path in enumerate(pdf_paths, start=2):
            worksheet.cell(row=1, column=column, value=pdf_path.stem)
            values = extract_data(pdf_path)
            for row, key in enumerate(DATA_ROWS, start=2):
                cell = worksheet.cell(row=row, column=1, value=key)
                cell.font = Font(bold=True)
                value = values.get(key)
                value_cell = worksheet.cell(row=row, column=column, value=value)
                if value is None:
                    value_cell.fill = UNAVAILABLE_FIELD_FILL
                if key == "Datum" and value is not None:
                    value_cell.number_format = "yyyy-mm-dd"

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
        description="Extract the general-data block from FIS result PDFs."
    )
    parser.add_argument(
        "--pdf-dir",
        type=Path,
        default=default_pdf_directory(),
        help="Folder containing downloaded PDFs (default: synced project data/pdf, "
        "or checkout data/pdf if needed)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help=f"Output workbook (default: {OUTPUT_FILE})",
    )
    args = parser.parse_args(argv)

    try:
        reset_workbook(args.output)
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
