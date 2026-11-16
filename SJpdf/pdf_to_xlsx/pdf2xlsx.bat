@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"
%PYTHON% "pdf2xlsx.py" %*
set "exit_code=%ERRORLEVEL%"
if not "%exit_code%"=="0" (
    echo.
    echo PDF to Excel failed with exit code %exit_code%.
    pause
    exit /b %exit_code%
)
for /f "usebackq delims=" %%D in (`%PYTHON% "%~dp0..\sj_paths.py"`) do set "DATA_DIR=%%D"

rem The data folder is synced by Nextcloud. Opening the workbook in Excel
rem makes Nextcloud lock it on the server; if the upload of the new version
rem is still running, the server rejects the lock ("Precondition Failed").
rem Wait a moment so the upload can finish first.
rem Set SJPDF_SYNC_WAIT to change the wait in seconds (0 = no wait).
if not defined SJPDF_SYNC_WAIT set "SJPDF_SYNC_WAIT=10"
if not "%SJPDF_SYNC_WAIT%"=="0" (
    echo Waiting %SJPDF_SYNC_WAIT% seconds for Nextcloud to upload the workbook ...
    timeout /t %SJPDF_SYNC_WAIT% /nobreak >nul
)
start "" "%DATA_DIR%\pdf_to_xlsx.xlsx"
exit /b %exit_code%
