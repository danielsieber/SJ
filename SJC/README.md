# SJC – Ski Jumping Charts

Ein lokales Python/Dash-Grundgeruest fuer die spaetere Anzeige interaktiver
Skisprung-Grafiken. Die Kopfzeile listet `.xlsx`-Dateien aus dem Review-Ordner;
links lassen sich Grafiktyp und Judge-Rangliste auswaehlen. Der Ordner kann
ueber `SJC_REVIEW_FOLDER` konfiguriert werden; standardmaessig wird
`~/SJA/Review` verwendet.

## Starten

```powershell
cd C:\repos\SJC
python -m pip install -r requirements.txt
python app.py
```

Anschliessend die von Dash angezeigte lokale Adresse im Browser oeffnen.
Der Wettkampftitel wird aus dem Namen der ausgewaehlten Arbeitsmappe gebildet.
Die Diagramm- und Ranglistenbuttons zeigen vorerst Platzhalter im Workspace;
Excel-Import und Diagramme folgen als naechster Ausbauschritt.
