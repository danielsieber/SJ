# SJA – Ski Jumping Analysis

SJA is the Streamlit application for ski-jumping competition analysis.

## Run

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Competition workbooks and reference lists are local data and are not included
in this repository. The app uses the first existing FIS `Web` folder
(containing `Listen.xlsm`, `Data_FIS`, `Data_Weekend`, `Data_pdf`):

1. Environment variable `SJA_WEB_ROOT`, if set
2. `C:\CLOUD-DANIEL\Family\Sport\FIS\Web` (Desktop)
3. `%USERPROFILE%\CLOUD-DANIEL\Family\Sport\FIS\Web` (Laptop)

The command-line pipeline (`src/pipeline/cli`) reads `Data_FIS` from the same
`Web` folder and `Liste_Codex.xlsx` from the parent `FIS` folder.
