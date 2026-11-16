# SJpdf

Tools for FIS ski jumping PDFs:

- `raceid_crawler` – reads the FIS calendar into `raceid\raceid.xlsx`
- `pdf_download` – downloads all PDFs for a Codex into `pdf`
- `pdf_to_xlsx` – extracts result PDFs into `pdf_to_xlsx.xlsx`

## Data folder

PDFs and Excel files are not stored in Git. All three tools use the same data
folder, found by `sj_paths.py` in this order:

1. Environment variable `SJPDF_DATA_DIR`, if set
2. `C:\CLOUD-DANIEL\SJA\SJpdf\data` (Desktop)
3. `%USERPROFILE%\CLOUD-DANIEL\SJA\SJpdf\data` (Laptop)
4. `data` inside this checkout

Show which folder is used:

```powershell
py sj_paths.py
```

## Setup

The `.bat` files use the Python launcher `py` if available, otherwise `python`.

```powershell
py -m pip install -r pdf_to_xlsx\requirements.txt -r raceid_crawler\requirements.txt
```

## Shared code

- `sj_paths.py` – data folder lookup (see above)
- `fis_web.py` – FIS page requests (with HTTP 429 retries), HTML parsing,
  calendar parsing and the current season code (a season starting in July
  2026 is `2027`); used by `raceid_crawler` and `pdf_download`

Tests for the shared modules:

```powershell
py -m unittest test_sj_paths test_fis_web
```
