# SJC – Ski Jumping Charts

Ein lokales Python/Dash-Grundgeruest fuer die spaetere Anzeige interaktiver
Skisprung-Grafiken. Die Kopfzeile listet `.xlsx`-Dateien aus dem Review-Ordner;
links lassen sich Grafiktyp und Judge-Rangliste auswaehlen. Der Ordner kann
ueber `SJC_REVIEW_FOLDER` konfiguriert werden. Sonst wird der erste vorhandene
dieser Ordner verwendet:

1. `C:\CLOUD-DANIEL\SJA\Review` (Desktop)
2. `%USERPROFILE%\CLOUD-DANIEL\SJA\Review` (Laptop)
3. `~/SJA/Review`

## Starten

```powershell
cd C:\repos\SJ\SJC
py -m pip install -r requirements.txt
py app.py
```

Anschliessend die von Dash angezeigte lokale Adresse im Browser oeffnen.
Der Wettkampftitel wird aus dem Namen der ausgewaehlten Arbeitsmappe gebildet.
Die Diagramm- und Ranglistenbuttons zeigen vorerst Platzhalter im Workspace;
Excel-Import und Diagramme folgen als naechster Ausbauschritt.
