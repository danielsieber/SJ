@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"

rem Overview of all folders of C:\downloaded_FIS_pdf: Archiv_Uebersicht.xlsx
%PYTHON% archive_overview.py %*
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo Archiv-Uebersicht failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
start "" "C:\downloaded_FIS_pdf\Archiv_Uebersicht.xlsx"
exit /b %exit_code%
