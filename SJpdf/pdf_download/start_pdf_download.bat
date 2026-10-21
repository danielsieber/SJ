@echo off
setlocal
cd /d "%~dp0"
python "FIS_pdf-download.py" %*
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo PDF downloader failed with exit code %exit_code%.
    pause
)
exit /b %exit_code%
