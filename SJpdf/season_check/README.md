# Saison-Check: fehlende JumperID und SchanzenID

`fis_season_check.py` prüft eine ganze FIS-Saison gegen `SJ-Listen.xlsm` im
SJpdf-Datenordner und schreibt die fehlenden Einträge nach
`Fehlende_JumperID.xlsx` im selben Ordner.

Start mit Doppelklick auf `start_season_check.bat` (fragt nach der Saison
und öffnet am Schluss `Fehlende_JumperID.xlsx`) oder:

```powershell
python fis_season_check.py --season 2026
```

`--season 2026` ist die Saison 2025/26. Mehrere Saisons in einer Tabelle:
`--season 2024-2026` (oder `--season 2024 2025 2026`); die Datei heisst dann
`Fehlende_JumperID 2024-2026.xlsx`. Mit `--output` und `--listen` lassen
sich andere Dateien angeben. `Event_ID` in `SJ-Listen.xlsm` wird nicht
verändert; das Skript liest die Saison nur.

PDFs, die schon im PDF-Archiv `C:\downloaded_FIS_pdf` liegen (siehe
`pdf_archive`, `--archive`), werden von dort gelesen und nicht neu geladen.

Wettkämpfe auf Schanzen bis HS30 werden nicht berücksichtigt (weder ihre
Springer noch die Schanze).

## Woher die Daten kommen

Für jeden Wettkampf der Saison (FIS-Kalender, wie beim Race-ID-Crawler):

- **FIS-Resultatseite:** FIS-Code, Name, Jahrgang und NSA jedes Athleten.
- **Ein PDF:** Startliste 1. Durchgang, Qualifikation oder Training (sonst
  ein Resultat-PDF; die Startliste Final Round hat kein Geburtsdatum).
  Daraus kommen Geburtsdatum und Club der Athleten und die Schanzendaten
  (Ort, HS, 95 %, K-Punkt, Meterwert, Windfaktoren, Gate Factor).
- **FIS-Athletenseite:** nur für fehlende Springer ohne Geburtsdatum im PDF
  oder ohne eindeutiges Gender (z. B. nur Mixed-Team).

Teamwettkämpfe: Die Teamzeilen (z. B. `JAPAN`) werden übersprungen, die
Athleten darunter erhalten die Nation ihres Teams. PDFs lokaler Organisatoren
(z. B. FESA) haben kein FIS-Format; daraus werden nur die Schanzendaten
gelesen, und als Ort gilt der Ort des FIS-Kalenders.

Gender kommt vom Wettkampf (Männer `M`, Frauen `F` wie in der Liste).

Alle Downloads werden in `%LOCALAPPDATA%\SJpdf\cache` gespeichert (nicht im
Nextcloud-Ordner). Ein zweiter Lauf lädt nichts neu; den Ordner löschen, um
alles frisch zu laden. Zwischen zwei FIS-Anfragen wartet das Skript 2
Sekunden; eine ganze Saison dauert beim ersten Lauf etwa eine halbe Stunde.

## Abgleich

- Ein Springer ist vorhanden, wenn Name und NSA in `Liste_Springer_ID`
  vorkommen oder ein Eintrag mit seinem FIS-Code die JumperID nach der Regel
  als `JumperID` oder `JumperID_pdf` hat.
- Steht der FIS-Code unter einem anderen Namen oder einer anderen NSA in der
  Liste (z. B. nach Heirat oder Nationenwechsel), kommt er ins Sheet
  `Prüfen`, nicht zu den fehlenden.
- Jede Schreibweise und NSA eines Springers wird geprüft, nicht nur die des
  ersten Wettkampfs: Wechselt ein Springer später Namen oder Nation, erscheint
  er mit dem ersten Wettkampf unter der neuen Schreibweise.
- Die SchanzenID ist `Ort HSxxx` (Ort aus dem PDF ohne Land, z. B.
  `Einsiedeln HS117`) und wird mit den Spalten `SchanzenID_pdf` und
  `SchanzenID` von `Liste_Schanzen_ID` verglichen. `Fehlende_SchanzenID`
  hat dieselben Spalten (A `SchanzenID_pdf`, B `SchanzenID`, beide mit der
  Schreibweise des PDFs); ist es eine bekannte Schanze in anderer
  Schreibweise, in B deren SchanzenID eintragen.

## Fehlende_JumperID.xlsx

- `Übersicht`: Zahlen und Meldungen (z. B. Wettkämpfe ohne PDF oder ohne
  Athleten auf der FIS-Seite).
- `Fehlende_JumperID`: Spalten A:H wie `Liste_Springer_ID` (JumperID_pdf,
  JumperID, Name, NSA, Gender, DoB, FIS_Code, JG), zum Einfügen ab Spalte A;
  `Age` berechnet die Tabelle. Die JumperID ist ein Vorschlag nach der Regel der
  Liste (NSA, je sechs Buchstaben des ersten und zweiten Namensworts). Rot
  markiert: JumperID schon vergeben, Gender oder Geburtsdatum unbekannt.
  Gelb: FIS kennt nur den Jahrgang (meist Nachwuchs); das Geburtsdatum ist
  dann wie in der Liste üblich der 1. Januar.
  Rechts daneben Club, Quelle des Geburtsdatums, erster Wettkampf, Hinweis.
- `Prüfen`: FIS-Code vorhanden, aber Name oder NSA weichen ab. Gleiches
  Layout wie `Liste_Springer_ID`: `JumperID_pdf` so, wie sie beim Lesen der
  PDFs entsteht (z. B. `HAYBOE MICHAE`), der Rest aus der Zeile mit dem
  FIS-Code (`AUT HAYBOE MICHAE`, `HAYBOECK Michael`, ...). A:H als zusätzliche
  Zeile in die Liste einfügen; rechts daneben Name und NSA bei der FIS und
  erster Wettkampf.
- `Fehlende_SchanzenID`: Spalten wie `Liste_Schanzen_ID`; `Pts / Gate für
  DGW` und die UTC-Spalten stehen nicht im PDF und sind von Hand zu ergänzen.

Tests: `python -m unittest test_fis_season_check.py` (aus diesem Ordner).
