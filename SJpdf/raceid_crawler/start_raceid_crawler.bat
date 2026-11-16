@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=python"
where py >nul 2>nul && set "PYTHON=py"

%PYTHON% fis_raceid_crawler.py %*
set "exit_code=%ERRORLEVEL%"

if not "%exit_code%"=="0" (
    echo.
    echo Crawler failed with exit code %exit_code%.
    pause
)

exit /b %exit_code%
