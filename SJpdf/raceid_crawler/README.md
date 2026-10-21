# SJpdf – FIS-Ranglisten-Crawler

Der Crawler liest den FIS-Kalender für Skispringen (`sectorcode=JP`) einer
Saison, öffnet die Detailseiten der Veranstaltungen und speichert die
Wettkämpfe in `data/raceid/raceid.xlsx`. PDFs können unter `data/pdf` abgelegt
werden. Standardmäßig wird Saison 2027 verwendet.

Die FIS verwendet zwei unterschiedliche IDs:

- `eventid` identifiziert die Veranstaltung und steht in der Kalender-URL,
  zum Beispiel `eventid=63279`.
- `raceid` identifiziert den einzelnen Wettkampf und steht im Resultat-Link,
  zum Beispiel `raceid=7934` für Codex `0381`.

Die Excel-Datei enthält die Sheets `men`, `women`, `mixed` und `all`. Die
Spalten sind `season`, `date`, `eventid`, `raceid`, `codex`, `place`,
`category` und `gender`. Das FIS-Datum wird als Excel-Datum gespeichert. Alle
Sheets sind aufsteigend nach Saison, Datum und Codex sortiert. Im Sheet `mixed`
stehen die Wettkämpfe mit Gender `A`, zum Beispiel Teamwettkämpfe.

## Ausführen

Zum Starten unter Windows `start_raceid_crawler.bat` doppelklicken. Zusätzliche Argumente,
zum Beispiel eine andere Saison, können auch über die Kommandozeile übergeben
werden:

```powershell
.\start_raceid_crawler.bat --season 2027
```

Alternativ lässt sich der Crawler direkt mit Python starten:

```powershell
python -m pip install -r requirements.txt
python fis_raceid_crawler.py
python fis_raceid_crawler.py --season 2027 --output ..\data\raceid\raceid.xlsx
```

Die Abfragen werden nacheinander mit einer Pause ausgeführt. Bei einer
FIS-Drosselung (HTTP 429) wartet der Crawler entsprechend der Server-Angabe
und wiederholt den Abruf. Wenn FIS eine Seite nicht liefert oder das
Kalenderformat nicht erkannt wird, bricht der Lauf mit einer Fehlermeldung ab,
statt eine unvollständige Excel-Datei als erfolgreich auszugeben.

## Tests

```powershell
python -m unittest discover -s tests
```
