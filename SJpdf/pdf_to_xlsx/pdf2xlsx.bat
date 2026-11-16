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
start "" "%DATA_DIR%\pdf_to_xlsx.xlsx"
exit /b %exit_code%
