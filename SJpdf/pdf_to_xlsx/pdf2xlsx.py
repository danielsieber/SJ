"""Create an Excel overview of the general-data block in FIS result PDFs."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import Font

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


def find_result_pdfs(pdf_directory: Path) -> list[Path]:
    """Return supported result PDFs in a stable order."""
    return sorted(
        (
            path
            for path in pdf_directory.iterdir()
            if path.is_file() and PDF_NAME_PATTERN.fullmatch(path.name)
        ),
        key=lambda path: path.name.casefold(),
    )


def create_workbook(pdf_paths: Sequence[Path], output_file: Path) -> None:
    """Extract each PDF once and save the selected fields in separate columns."""
    if not pdf_paths:
        raise FileNotFoundError("No matching RLT, RLQ, RTRIA, or RL PDFs found.")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Daten"
    worksheet["A1"] = "pdf"

    for column, pdf_path in enumerate(pdf_paths, start=2):
        worksheet.cell(row=1, column=column, value=pdf_path.stem)
        values = extract_data(pdf_path)
        for row, key in enumerate(DATA_ROWS, start=2):
            cell = worksheet.cell(row=row, column=1, value=key)
            cell.font = Font(bold=True)
            value = values.get(key)
            value_cell = worksheet.cell(row=row, column=column, value=value)
            if key == "Datum" and value is not None:
                value_cell.number_format = "yyyy-mm-dd"

    worksheet.column_dimensions["A"].width = 20
    for column in range(2, len(pdf_paths) + 2):
        worksheet.column_dimensions[worksheet.cell(row=1, column=column).column_letter].width = 32

    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract the general-data block from FIS result PDFs."
    )
    parser.add_argument(
        "--pdf-dir",
        type=Path,
        default=PDF_DIRECTORY,
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
    except (FileNotFoundError, OSError, PdfParseError) as error:
        print(f"PDF to Excel failed: {error}", file=sys.stderr)
        return 1

    print(f"Processed {len(pdf_paths)} PDF(s): {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
