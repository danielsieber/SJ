# Ski Jumping projects

This repository groups the ski-jumping applications:

- [SJA](./SJA/README.md) — competition analysis
- [SJC](./SJC/README.md) — ski-jumping charts
- [SJpdf](./SJpdf/pdf_download/README.md) — FIS race-ID index and PDF download

The repository contains program files, documentation, and tests. Local
workbooks, downloaded PDFs, and personal notes are intentionally excluded.

## Tests

Run each suite from its project directory:

```powershell
cd SJA
python -m unittest discover -s tests
cd ..\SJpdf
cd raceid_crawler
python -m unittest discover -s tests
cd ..
python -m unittest discover -s pdf_download -p "test_*.py"
```
