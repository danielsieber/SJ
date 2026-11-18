# Saison-Check: fehlende JumperID und SchanzenID

`fis_season_check.py` prüft eine ganze FIS-Saison gegen `SJ-Listen.xlsm` im
SJpdf-Datenordner und schreibt die fehlenden Einträge nach
`Fehlende_JumperID.xlsx` im selben Ordner.

Start mit Doppelklick auf `start_season_check.bat` (fragt nach der Saison
und öffnet am Schluss `Fehlende_JumperID.xlsx`) oder:

```powershell
python fis_season_check.py --season 2026
```

`--season 2026` ist die Saison 2025/26. Mit `--output` und `--listen` lassen
sich andere Dateien angeben. `Event_ID` in `SJ-Listen.xlsm` wird nicht
verändert; das Skript liest die Saison nur.

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

- Ein Springer ist vorhanden, wenn sein FIS-Code in `Liste_Springer_ID`
  steht oder Name und NSA dort vorkommen.
- Steht der FIS-Code unter einem anderen Namen oder einer anderen NSA in der
  Liste (z. B. nach Heirat oder Nationenwechsel), kommt er ins Sheet
  `Prüfen`, nicht zu den fehlenden.
- Die SchanzenID ist `Ort HSxxx` (Ort aus dem PDF ohne Land, z. B.
  `Einsiedeln HS117`) und wird mit `Liste_Schanzen_ID` verglichen.

## Fehlende_JumperID.xlsx

- `Übersicht`: Zahlen und Meldungen (z. B. Wettkämpfe ohne PDF oder ohne
  Athleten auf der FIS-Seite).
- `Fehlende_JumperID`: Spalten A:G wie `Liste_Springer_ID` (JumperID,
  Name, NSA, Gender, DoB, FIS_Code, JG), zum Einfügen ab Spalte A; `Age`
  berechnet die Tabelle. Die JumperID ist ein Vorschlag nach der Regel der
  Liste (NSA, je fünf Buchstaben des ersten und zweiten Namensworts). Rot
  markiert: JumperID schon vergeben, Gender oder Geburtsdatum unbekannt.
  Gelb: FIS kennt nur den Jahrgang (meist Nachwuchs); das Geburtsdatum ist
  dann wie in der Liste üblich der 1. Januar.
  Rechts daneben Club, Quelle des Geburtsdatums, erster Wettkampf, Hinweis.
- `Prüfen`: FIS-Code vorhanden, aber Name oder NSA weichen ab.
- `Fehlende_SchanzenID`: Spalten wie `Liste_Schanzen_ID`; `Pts / Gate für
  DGW` und die UTC-Spalten stehen nicht im PDF und sind von Hand zu ergänzen.

Tests: `python -m unittest test_fis_season_check.py` (aus diesem Ordner).
