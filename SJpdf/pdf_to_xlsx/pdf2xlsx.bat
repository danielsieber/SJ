@echo off
setlocal
cd /d "%~dp0"
python "pdf2xlsx.py" %*
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo PDF to Excel failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
start "" "%~dp0..\data\pdf_to_xlsx.xlsx"
exit /b %exit_code%
