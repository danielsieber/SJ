@echo off
setlocal
cd /d "C:\repos\SJC"

start "" /B python app.py
powershell -NoProfile -Command "$deadline = (Get-Date).AddSeconds(30); do { try { $response = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:8050' -TimeoutSec 2; if ($response.StatusCode -eq 200) { Start-Process 'http://127.0.0.1:8050'; exit 0 } } catch { }; Start-Sleep -Milliseconds 500 } while ((Get-Date) -lt $deadline); exit 1"
if errorlevel 1 (
    echo.
    echo SJC konnte nicht gestartet werden. Bitte pruefe Python und die Installation der Abhaengigkeiten.
    pause
)
