@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"
%PYTHON% "FIS_pdf-download.py" %*
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo PDF downloader failed with exit code %exit_code%.
    pause
)
exit /b %exit_code%
