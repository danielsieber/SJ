# FIS PDF Download

Start `start_pdf_download.bat` to open the Codex input window. Enter a two-,
three-, or four-digit Codex and press Enter or select **PDFs herunterladen**.
Shorter Codices are padded to four digits, preserving leading zeros. When the
Codex occurs in multiple events, select the intended event.

The downloader reads the sheet `Event_ID` of `SJ-Listen.xlsm` in the SJpdf data folder (see `SJpdf/README.md`), then downloads every PDF
attached to the same event and gender as the selected Codex. This includes
official results, start lists (SLR1/SLRF are used by `pdf_to_xlsx` for the
jump order), training documents, and other FIS attachments.
FIS's `Content-Disposition` filename is used for each PDF. Downloads are saved
directly in `pdf` inside the data folder, alongside a `myFolderName.txt` with the suggested
destination folder name:
`yyyy-mm-dd place categorygender` (for example `2026-09-13 Einsiedeln FCM`).
The original FIS link of each PDF is saved in `pdf_links.csv` (columns
`pdf`, `url`) in the same folder; `pdf_to_xlsx` uses it to link the PDFs.
The date is the latest competition date for that event and gender; category
and gender come from `Event_ID`. Before each download, any files
already directly in `..\data\pdf` are moved into
`..\data\pdf\archiv\<myFolderName>`; subfolders are left in place. If files
exist without a valid `myFolderName.txt`, the download stops without moving
them. Existing PDFs are not overwritten.

If the Codex is not yet in the Excel index, the tool searches the FIS calendar
for the indexed season and reports its progress.

Run the unit tests from the project folder with:

```powershell
python -m unittest discover -s pdf_download -p "test_*.py"
```
