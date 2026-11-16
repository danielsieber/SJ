# FIS PDF to Excel

`pdf2xlsx.bat` reads matching FIS result PDFs directly from `pdf` in the
SJpdf data folder (see the main `SJpdf/README.md`). Subfolders, including
archived event folders, are not searched. It creates `pdf_to_xlsx.xlsx` in the
same data folder.
Only files named like
`2027JP0372RL`, `2027JP0372RLT`, `2027JP0372RLQ`, or `2027JP0372RTRIA`
(with a `.pdf` extension) are processed. Other PDF types are ignored.

The workbook contains the sheet `Daten`. Each supported PDF has its own column.
The first column contains general competition data, competition/weather
information for up to three rounds (including length), and the last-page
footer date, place, and codex. The competition information is written from
row 19; footer fields follow it at rows 49-51. The event date is stored as an
Excel date and displayed as `yyyy-mm-dd`. Unavailable individual fields are
left blank with a light red background. Each run removes all other sheets and
clears every column after column A in `Daten` before writing the current
results.

In addition, every processed PDF gets its own sheet named like the PDF
(for example `2027JP0281RLT`) with the result table (`pdf2xlsx_Resultate.py`).
There is one row per athlete, and every sheet has the same fixed columns:
rank, bib, name, NSA, U23, club, date of birth, status (Not Qualified, DSQ,
DNS, NPS), then per round speed (rank), distance (fall `*`, rank), distance
points, judge marks A-E, judge points, gate (gate before a change and the
coach-request mark `©`), gate points, wind, wind points, round total and round
rank, then total and remarks. Fields a PDF does not contain stay empty.
Shared ranks (`=5.`) are stored as plain numbers.

Values are read with their position on the page, so empty cells never shift
values into another column. Each PDF is checked arithmetically (distance and
judge points, round totals, total, and all ranks). If a check fails or the
layout is unknown, the sheet only contains the reason (red tab), and the
converter prints a warning; no unchecked values are written. A swap of two
judges (A-E) cannot be detected by these checks.

Install dependencies from this directory with:

```powershell
python -m pip install -r requirements.txt
```

Run `pdf2xlsx.bat`, or call `python pdf2xlsx.py` directly. Optional
`--pdf-dir` and `--output` arguments can be used to select another input
folder or output workbook. After a successful run, the batch file opens the
generated workbook in its associated application.
Close the output workbook in Excel before running the converter; Excel locks
open workbooks and prevents them from being overwritten.

Run the tests from the `SJpdf` directory with:

```powershell
python -m unittest discover -s pdf_to_xlsx\tests -t . -v
```
