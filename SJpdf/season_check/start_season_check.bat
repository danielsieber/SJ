@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"

rem Season as argument (start_season_check.bat 2026) or asked here.
set "SEASON=%~1"
if "%SEASON%"=="" set /p "SEASON=FIS-Saison (z. B. 2026 fuer 2025/26): "
if "%SEASON%"=="" (
    echo Keine Saison angegeben.
    pause
    exit /b 1
)

%PYTHON% fis_season_check.py --season %SEASON%
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo Saison-Check failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
for /f "usebackq delims=" %%D in (`%PYTHON% "%~dp0..\sj_paths.py"`) do set "DATA_DIR=%%D"

rem Let Nextcloud upload the new file before Excel locks it (see pdf2xlsx.bat).
if not defined SJPDF_SYNC_WAIT set "SJPDF_SYNC_WAIT=10"
if not "%SJPDF_SYNC_WAIT%"=="0" (
    echo Waiting %SJPDF_SYNC_WAIT% seconds for Nextcloud to upload the workbook ...
    timeout /t %SJPDF_SYNC_WAIT% /nobreak >nul
)
start "" "%DATA_DIR%\Fehlende_JumperID.xlsx"
exit /b %exit_code%
