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

## xlsx pro Event-Ordner

`archive_xlsx.py` schreibt in jeden Event-Ordner die Excel-Datei von
pdf2xlsx (dieselben Sheets wie `latest event.xlsx`), benannt nach dem
Wettkampf:

```
C:\downloaded_FIS_pdf\2026-09-13 FCM Einsiedeln LH\2026-09-13 FCM Einsiedeln HS117.xlsx
```

Datum, Kategorie und Gender kommen aus dem Ordnernamen, die SchanzenID
(`Ort HSxxx`) aus den PDFs; als Ort gilt der Ort des PDFs mit Ländercode
(`Einsiedeln (SUI)`), sonst der Ort des FIS-Kalenders. `JumperID` und
`Pts / Gate DGW` werden wie bei pdf2xlsx in `SJ-Listen.xlsm` gesucht.

Start mit Doppelklick auf `start_archive_xlsx.bat` (fragt nach den Saisons,
leer = alle Ordner; öffnet am Schluss `xlsx_log.csv`) oder:

```powershell
python archive_xlsx.py --season 2024-2026
```

- Ein Ordner, dessen xlsx neuer ist als seine PDFs und `SJ-Listen.xlsm`,
  wird übersprungen; `--force` schreibt alle neu. Nach dem Ergänzen von
  `SJ-Listen.xlsm` werden also alle xlsx neu geschrieben.
- Die Saison eines Ordners steht in `archiv_log.csv`.
- Ordner ohne Resultat-PDF (Gesamtwertungen wie Four Hills oder Raw Air,
  abgesagte Events mit nur dem Programm) bekommen keine xlsx.
- Events auf Schanzen bis HS30 werden nicht ausgewertet (Status
  `kleine Schanze`, eine ältere xlsx im Ordner wird gelöscht).
- `xlsx_log.csv` im Archiv protokolliert pro Ordner Status, fehlende
  JumperID und SchanzenID, Zahl der Warnungen und die Warnungen selbst.

Tests: `python -m unittest test_fis_pdf_archive.py test_archive_xlsx.py`
(aus diesem Ordner).
