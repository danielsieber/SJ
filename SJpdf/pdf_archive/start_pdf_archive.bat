@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"

rem Seasons as arguments (start_pdf_archive.bat 2016 2027) or asked here.
rem Empty input: from 2016 to the current season. The run can be stopped
rem (close the window) and started again; it continues where it stopped.
set "FROM=%~1"
set "TO=%~2"
if "%FROM%"=="" set /p "FROM=Erste Saison (leer = 2016): "
if "%TO%"=="" set /p "TO=Letzte Saison (leer = aktuelle): "
set "ARGS="
if not "%FROM%"=="" set "ARGS=%ARGS% --from %FROM%"
if not "%TO%"=="" set "ARGS=%ARGS% --to %TO%"

%PYTHON% fis_pdf_archive.py %ARGS%
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo PDF-Archiv failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
start "" "C:\downloaded_FIS_pdf"
exit /b %exit_code%
