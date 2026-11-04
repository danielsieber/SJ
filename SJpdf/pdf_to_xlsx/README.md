# FIS PDF to Excel

`pdf2xlsx.bat` reads matching FIS result PDFs recursively from
`..\data\pdf\archiv` and creates
`..\data\pdf_to_xlsx.xlsx`. Archived event subfolders are searched as well.
Only files named like
`2027JP0372RL`, `2027JP0372RLT`, `2027JP0372RLQ`, or `2027JP0372RTRIA`
(with a `.pdf` extension) are processed. Other PDF types are ignored.

The workbook contains the sheet `Daten`. Each supported PDF has its own column;
the first column contains the field names for block 1 (title, event details,
results heading, and Hill Data). The date is stored as an Excel date and
displayed as `yyyy-mm-dd`. Unavailable individual fields are left blank.

Install dependencies from this directory with:

```powershell
python -m pip install -r requirements.txt
```

Run `pdf2xlsx.bat`, or call `python pdf2xlsx.py` directly. Optional
`--pdf-dir` and `--output` arguments can be used to select another input
folder or output workbook.
Close the output workbook in Excel before running the converter; Excel locks
open workbooks and prevents them from being overwritten.

Run the tests from the `SJpdf` directory with:

```powershell
python -m unittest discover -s pdf_to_xlsx\tests -t . -v
```
