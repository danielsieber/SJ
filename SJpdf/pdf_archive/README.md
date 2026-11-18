# PDF-Archiv: alle FIS-PDFs pro Event

`fis_pdf_archive.py` lädt alle Skisprung-PDFs der FIS (Startlisten,
Resultate, Programm, Standings, ...) mehrerer Saisons herunter, mit einem
Ordner pro Event in `C:\downloaded_FIS_pdf`.

Start mit Doppelklick auf `start_pdf_archive.bat`. Sie fragt nach der ersten
und letzten Saison (leer: 2016 bis zur aktuellen Saison) und öffnet am Schluss
den Ordner. Oder:

```powershell
python fis_pdf_archive.py --from 2016 --to 2027
```

## Ordner

Ein Event ist eine Serie mit derselben `EventNr` wie im Sheet `Event_ID`
(gleicher Ort, `cat2`, Gender und Schanzentyp innerhalb von 5 Tagen; die
Qualifikation gehört zu ihrem Weltcup oder Grand Prix). Der Ordnername ist
wie `list_event`, aber ohne EventNr und Codex, mit dem Datum des letzten
Wettkampfs:

```
C:\downloaded_FIS_pdf\2027-03-14 COCM Zakopane LH
```

Im Ordner liegen alle PDFs aller Wettkämpfe der Serie unter dem FIS-Namen
(z. B. `2026JP3098RL.pdf`) und `pdf_links.csv` mit dem FIS-Link jedes PDFs
(pdf2xlsx liest diese Datei für die Links in `Notes`).

## Ablauf und Dauer

- Die neueste Saison zuerst, dann rückwärts bis zur ersten.
- Die Saisons werden wie beim Race-ID-Crawler gelesen, nur im Speicher;
  `Event_ID` in `SJ-Listen.xlsm` wird nicht verändert.
- Ein PDF, dessen Link schon in `pdf_links.csv` steht, wird nicht neu
  geladen. Ein Lauf kann jederzeit abgebrochen (Fenster schliessen) und neu
  gestartet werden; er macht dort weiter. Ein späterer Lauf der aktuellen
  Saison lädt nur die neuen PDFs.
- Zwischen zwei FIS-Anfragen wartet das Skript 2 Sekunden. Ein PDF braucht
  so gut 2 Sekunden; alle Saisons ab 2016 sind rund 25'000 PDFs (etwa 6 GB)
  und 15 Stunden.
- Die FIS-Seiten abgeschlossener Saisons werden in
  `%LOCALAPPDATA%\SJpdf\cache` zwischengespeichert, die der aktuellen Saison
  immer frisch gelesen. Die PDFs selbst kommen nicht in den Cache.
- `archiv_log.csv` im Zielordner protokolliert pro Ordner, wann wie viele
  PDFs neu dazukamen, und Fehler einzelner PDFs.

Tests: `python -m unittest test_fis_pdf_archive.py` (aus diesem Ordner).
