@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"

rem Seasons as argument (start_archive_xlsx.bat 2024-2026) or asked here.
rem Empty input: all folders of C:\downloaded_FIS_pdf. Folders whose xlsx is
rem newer than their PDFs and SJ-Listen.xlsm are skipped.
set "SEASON=%~1"
if "%SEASON%"=="" set /p "SEASON=Saisons (z. B. 2026 oder 2024-2026, leer = alle): "
set "ARGS="
if not "%SEASON%"=="" set "ARGS=--season %SEASON%"

%PYTHON% archive_xlsx.py %ARGS%
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo Archiv-xlsx failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
start "" "C:\downloaded_FIS_pdf\xlsx_log.csv"
exit /b %exit_code%
