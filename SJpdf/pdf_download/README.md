# FIS PDF Download

Start `start_pdf_download.bat` to open the Codex input window. Enter a two-,
three-, or four-digit Codex and press Enter or select **PDFs herunterladen**.
Shorter Codices are padded to four digits, preserving leading zeros. When the
Codex occurs in multiple events, select the intended event.

The downloader reads `..\data\raceid\raceid.xlsx`, then downloads every PDF
attached to the same event and gender as the selected Codex. This includes
official results, start lists, training documents, and other FIS attachments.
FIS's `Content-Disposition` filename is used for each PDF. Downloads are saved
directly in `..\data\pdf`, alongside a `myFolderName.txt` with the suggested
destination folder name:
`yyyy-mm-dd place categorygender` (for example `2026-09-13 Einsiedeln FCM`).
The date is the latest competition date for that event and gender; category
and gender come from the race-ID workbook. Before each download, any files
already directly in `..\data\pdf` are moved into
`..\data\pdf\archiv\<myFolderName>`; subfolders are left in place. If files
exist without a valid `myFolderName.txt`, the download stops without moving
them. Existing PDFs are not overwritten. The race-ID workbook remains in this
project's `..\data\raceid` folder.

If the Codex is not yet in the Excel index, the tool searches the FIS calendar
for the indexed season and reports its progress.

Run the unit tests from the project folder with:

```powershell
python -m unittest discover -s pdf_download -p "test_*.py"
```
