@echo off
title SJA - Ski Jumping Analyse
echo Starte Ski Jumping Analyse...
cd /d "%~dp0"
set PYTHON_CMD=python

py --version >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON_CMD=py
)

%PYTHON_CMD% -m streamlit run app.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo Fehler beim Starten von Streamlit!
    pause
)
