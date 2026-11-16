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

Tests for the shared path lookup:

```powershell
py -m unittest test_sj_paths
```
