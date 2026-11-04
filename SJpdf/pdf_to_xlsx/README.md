# FIS PDF to Excel

`pdf2xlsx.bat` reads matching FIS result PDFs directly from
`%USERPROFILE%\CLOUD-DANIEL\SJA\SJpdf\data\pdf` first and then
`..\data\pdf` if no matching result PDFs are found in the synced folder.
Subfolders, including archived event folders, are not searched. It creates
`..\data\pdf_to_xlsx.xlsx`.
Only files named like
`2027JP0372RL`, `2027JP0372RLT`, `2027JP0372RLQ`, or `2027JP0372RTRIA`
(with a `.pdf` extension) are processed. Other PDF types are ignored.

The workbook contains the sheet `Daten`. Each supported PDF has its own column;
the first column contains the field names for block 1 (title, event details,
results heading, and Hill Data). The date is stored as an Excel date and
displayed as `yyyy-mm-dd`. Unavailable individual fields are left blank with a
light red background. Each run removes all other sheets and clears every
column after column A in `Daten` before writing the current results.

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
