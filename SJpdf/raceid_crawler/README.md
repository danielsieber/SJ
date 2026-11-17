# SJpdf – FIS-Ranglisten-Crawler

Der Crawler liest den FIS-Kalender für Skispringen (`sectorcode=JP`) einer
Saison, öffnet die Detailseiten der Veranstaltungen und speichert die
Wettkämpfe im Sheet `Event_ID` von `SJ-Listen.xlsm` im SJpdf-Datenordner
(siehe `SJpdf/README.md`). PDFs liegen im selben Datenordner unter `pdf`.
Standardmäßig wird die aktuelle Saison (2027) verwendet.

Die FIS verwendet zwei unterschiedliche IDs:

- `eventid` identifiziert die Veranstaltung und steht in der Kalender-URL,
  zum Beispiel `eventid=63279`.
- `raceid` identifiziert den einzelnen Wettkampf und steht im Resultat-Link,
  zum Beispiel `raceid=7934` für Codex `0381`.

Jeder Lauf ersetzt den Inhalt von `Event_ID` durch die Wettkämpfe der
gecrawlten Saison. Die Spalten sind `season`, `date`, `eventid`, `raceid`,
`codex`, `place`, `hill`, `cat1`, `cat2`, `gender` (Gender `M`, `W` oder `A` für
Mixed, zum Beispiel Teamwettkämpfe) und `EventNr`. Das FIS-Datum ist ein Excel-Datum
(`yyyy-mm-dd`), `eventid`, `raceid` und `codex` sind Text (führende Nullen
bleiben). Die Zeilen sind nach `EventNr`, innerhalb nach Datum und dann nach Codex sortiert,
mit Filter über den ganzen Bereich.

`hill` ist der Schanzentyp aus der FIS-Beschreibung des Wettkampfs, z. B.
„Individual Large Hill“: `SH`, `MH`, `NH` (Normal Hill), `LH` (Large Hill) oder
`FH` (Flying Hill). Die Schanzengrösse (HS) steht nicht auf den FIS-Seiten des
Kalenders; sie steht erst in den Ergebnis-PDFs.

`cat1` ist die FIS-Kategorie (z. B. `WC`, `GP`, `COC`, `QUA`). `cat2` ist gleich
wie `cat1`, nur eine Qualifikation (`QUA`) erhält die Kategorie der
Meisterschaft, zu der sie gehört (`OWG`, `WSC`, `SFWC` oder `WJC` am selben Ort,
mit gleichem Gender und `hill`, innerhalb von 5 Kalendertagen), sonst `GP` (Mai
bis Oktober) bzw. `WC` (übrige Monate).

`EventNr` nummeriert die Wettkampfserien nach dem Datum ihres letzten
Wettkampfs (bei Gleichstand nach dem ersten Wettkampf, dann Codex): Wettkämpfe mit
gleichem Ort, gleichem `cat2`, gleichem Gender und gleichem `hill` innerhalb von
5 Kalendertagen ab dem ersten Wettkampf erhalten dieselbe Nummer. Die
Qualifikation gehört so zur Serie ihres Weltcups bzw. Grand Prix.

Excel selbst schreibt das Sheet (über `pywin32`, in einer eigenen,
unsichtbaren Excel-Instanz, ohne die Makros auszuführen). Dadurch bleiben
Makros, Schaltflächen, Tabellen und Formeln der anderen Sheets erhalten; die
Datei wird einmal gespeichert. Ist `SJ-Listen.xlsm` in Excel geöffnet, bricht
der Crawler mit einer Meldung ab: Datei schliessen und erneut starten. Findet
der Crawler keine Wettkämpfe, bleibt `Event_ID` unverändert.

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
python fis_raceid_crawler.py --season 2027 --output ..\data\SJ-Listen.xlsm
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
