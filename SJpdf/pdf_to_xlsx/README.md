# FIS PDF to Excel

`pdf2xlsx.bat` reads matching FIS result PDFs directly from `pdf` in the
SJpdf data folder (see the main `SJpdf/README.md`). Subfolders, including
archived event folders, are not searched. It creates `latest event.xlsx` in the
same data folder.
Only files named like
`2027JP0372RL`, `2027JP0372RLT` (also `RLT1`, `RLT2`, ... for several
official trainings), `2027JP0372RLQ`, or `2027JP0372RTRIA`
(with a `.pdf` extension) are processed. Other PDF types are ignored.

Supported layouts: Swiss Timing (World Cup, Grand Prix, COC, also the
older layout of 2023-2025 without bold print), EWOXX and the layouts of
local organisers with mixed-case headers (`Rank`, `Bib`, `NOC Code`), tables
without row lines, two tables side by side or one below the other, and
team events (team rows `GER - Germany` with their jumpers `8-1`, `8-2`, ...;
the jumpers get the nation of their team, `(W)`/`(M)` of mixed teams is
removed). PDFs without a `FIS ...` title (Alpen Cup, FESA games) are read
with their first title line.

- KO competitions (Four Hills): the qualification `RLQ` with its total column
  `Points`; in `RL`/`RL1` the pair winners and lucky losers (`*`, remark
  `Lucky Loser`) are ranked apart from the other losers (`Not Qualified`).
  The Four Hills standings (`...4H.pdf`) are not read.
- Olympic Games / Youth Olympic Games: header labels at slightly different
  heights, names without bold print (`DESCHWANDEN` / `Gregor` on two lines,
  initials such as `SUNDAL KE`).
- WinSpringen lists of DSV organisers (Alpen Cup, FESA, OPA in Germany,
  printed with "Microsoft Print to PDF"): German columns (`Rang`, `StNr`,
  `Weite`, `W-N.`, `H-N.`, `Luke`, `G-N.`, `T-N.`), two lines per athlete
  (round 1 with name, round 2 with club and nation), training lists with
  speed/distance ranks (equal values share a rank without gap), age groups
  (`Boys, 64.0 m Schanze`) ranked apart, DNS/DSQ lists, empty distance or
  wind points counted as 0. Date from `15.03.2026`, `26th October 2024` or
  the print date `Erstellt am`. Team lists of WinSpringen are not supported.
- Swiss Timing until 2023: labels `Rk`, `Speed`, `Dist`, `DP`, `JP`, `Gate`,
  `Pts`, `m/s` in up to three header lines, round total and round rank below
  `Round` / `Points`; small hills without gate / wind columns (training:
  speed, distance and gate per round); `Prequalified` athletes of a
  qualification (remark `Prequalified`, no judges and total); start lists
  with `Name` one line above `Bib`. The gender is taken from the title lines
  (`Ladies`, `Girls`, `Boys`, `Mixed`); when a PDF names none, from the
  athletes found in `Liste_Springer_ID` (column `Gender`: M men, F women;
  mixed when at least a quarter belongs to the other gender), noted in
  `Notes`.
- A round never counts below 0 points (negative sum printed as `0.0`).
- Team events of Olympic Games, Youth Olympic Games and Junior World
  Championships: round lines above and below the name, name above the bib.
  A team start list without start order for the later groups only gives a
  warning; the rounds are then ordered by the team standings.
- JumperID: a double first name in the PDF (`CHERVET Jules-Lucien`) finds the
  listed `FRA CHERVE JULES`; initials of Olympic lists (`SUNDAL KE`) find the
  only listed jumper of the nation with this surname and these initials.

The workbook contains the sheet `Daten`. Each supported PDF has its own column.
The first column contains general competition data, competition/weather
information for up to three rounds (including length), and the last-page
footer date, place, and codex. The competition information is written from
row 21; footer fields follow it at rows 51-53. The event date is stored as an
Excel date and displayed as `yyyy-mm-dd`. Unavailable individual fields are
left blank with a light red background. Each run removes all other sheets and
clears every column after column A in `Daten` before writing the current
results.

In addition, every processed PDF gets its own sheet named like the PDF
(for example `2027JP0281RLT`) with the result table (`pdf2xlsx_Resultate.py`).
There is one row per athlete, and every sheet has the same fixed columns:
rank, bib, name, NSA, U23, club, date of birth, status (Not Qualified, DSQ,
DNS, NPS), then per round speed (rank), distance (fall `*`, rank), distance
points, judge marks A-E, judge points, gate (gate before a change and the
coach-request mark `©`), gate points, wind, wind points, round total and round
rank, then total and remarks. Fields a PDF does not contain stay empty.
Shared ranks (`=5.`) are stored as plain numbers.

Values are read with their position on the page, so empty cells never shift
values into another column. Each PDF is checked arithmetically (distance and
judge points, round totals, total, and all ranks). If a check fails or the
layout is unknown, the sheet only contains the reason (red tab), and the
converter prints a warning; no unchecked values are written. A swap of two
judges (A-E) cannot be detected by these checks.
An unreadable PDF does not stop the others: `Ranglisten` is built from the
readable PDFs (warning naming the others; without a readable RL1 or start
list the rounds are ordered as described below), and `Notes` shows below
`Anzahl PDFs` which PDFs were read and, in red, which were not. Row 12 of `Daten`
(`PDF readable`, below `Rangliste`) says `yes` or `no` for each PDF column. A PDF without
text (scan) is left out of `Daten` as well.

For a competition (`...RL.pdf`), the intermediate results after the first
round (`...RL1.pdf`, same folder) are read as well; they get their own sheet
and a column in `Daten`, both directly before the RL. The start lists of the
first round (`...SLR1.pdf`) and of the final round (`...SLRF.pdf`) get a sheet
with start order, bib and name (sheet order SLR1, RL1, SLRF, RL); they have
no column in `Daten`.

The sheet `Ranglisten` (`pdf2xlsx_Ranglisten.py`) combines all rounds in jump
order. Its columns are the same on every run (it is loaded into a database
later): A:H athlete and status, I:AB the values of the round (Speed ...
Round Rank), AC `Total` (only in the second competition round), AD `JumpID`,
AE `JumperID`, AF `Date` (yyyy-mm-dd) and AG `Time` (hh:mm:ss) of the jump,
AH `PDF` (source PDF of the row).

- Rounds follow the `Daten` order (date, start time); a PDF with two rounds
  gives two blocks. `JumpID` is the competition name, the round number and
  the type: `T` training (RLT), `P` trial round (RTRIA), `Q` qualification
  (RLQ), `W` competition (RL), e.g. `2026-09-13 FCM Einsiedeln HS117 01T`.
  Without a competition name, `Ranglisten` is not created (warning).
- Within a round the rows are in start order. Training, trial round and
  qualification: bib order. Competition: order of the start lists SLR1 and
  SLRF (this also covers the zip order 50., 1., 49., 2., ... of World Cup
  competitions around New Year, and Grand Prix lists starting with the
  highest bib). An athlete missing in a start list is inserted with a
  warning: round 1 at his bib position, the second round at his position in
  the reverse RL1 ranking. Without a start list, round 1 is in bib order and
  the second round in the reverse order of RL1, with a warning.
- The first competition round is taken from RL1 (an athlete disqualified in
  the second round has no first-round values in RL anymore); the second
  round from RL. Only athletes who started are listed in a second round.
- Status: `DSQ`, `DNS`, `NPS` for the round concerned; `Not Qualified` when
  the athlete may not start in the next round.
- If RL1 is missing, the first round is taken from RL and a warning is shown.
- Rounds 3 and 4 (ski flying championships): reverse order of the standings
  after the round before, with a warning (there is no start list).
- Team events: the jumpers (not the team rows), round by round in the order
  of the start lists; without a start list group by group (first jumpers of
  all teams, then the second, ...), the teams in bib order in round 1 and in
  the reverse order of the team standings later (warning). Each team total
  is checked against the points of its jumpers.
- `Date` is the date of the round's PDF. `Time`: the first jumper of a round
  gets the start time, the last one the end time of the round (`Start Time n`,
  `End Time n` of DG n in `Daten`; RL1 holds round 1, RL round 2), the jumpers
  in between are spaced evenly. A PDF without weather block (e.g. a
  qualification) uses `Start Time` and `Finish Time`. Rows without a jump
  (DNS, Not Qualified, ...) get no time.

Jumpers and hills are checked against `SJ-Listen.xlsm` in the SJpdf data
folder (`pdf2xlsx_Listen.py`; the file is only read, it may be open in Excel):

- `JumperID` in `Ranglisten` is built from NSA and name by the rule of
  `Liste_Springer_ID` (NSA, first six letters of the first and of the second
  word of the name, e.g. `ACHLEITNER Alexander` -> `AUT ACHLEI ALEXAN`).
  When this ID is in the column `JumperID_pdf` (spelling of the PDFs, e.g.
  `HAYBOE MICHAE` for `Hayboeck Michael` without NSA), the column `JumperID`
  of that row is used (`AUT HAYBOE MICHAE`). Otherwise it must be listed there for a jumper with the same surname (first word), so
  `AUT STEINA SIMON` of STEINACHER Simon is not taken for another jumper.
  When the rule ID is not listed, name and NSA are looked up (special IDs,
  e.g. `ITA DEL LUCREZ` for DEL BIANCO Lucrezia). Otherwise the JumperID cell
  stays empty (light red) and `Notes` shows a red warning; when all are
  found, `Notes` says so.
- `Pts / Gate DGW` in `Daten` (row below `TWC`) is the value
  `Pts / Gate für DGW` of the hill in `Liste_Schanzen_ID`, found by place
  and hill size (`Einsiedeln HS117`) in the column `SchanzenID_pdf`
  (spelling of the PDFs, e.g. `Lake Placid/ORDA HS128`). The column
  `SchanzenID` of that row is the name of the hill in the competition copy
  and in `JumpID`. `Notes` shows whether the
  SchanzenID was found (a missing one is a red warning), above the JumperID
  line.
- If `SJ-Listen.xlsm` cannot be read, there is a red warning and the
  conversion goes on without JumperID and Pts / Gate DGW.

Besides `latest event.xlsx`, the same workbook is saved as
`<last date> <type><gender> <place> HS<hill size>.xlsx`, for example
`2026-09-13 FCM Einsiedeln HS117.xlsx` (types FC, WC, GP, COC, ICOC, AC;
gender M, W, A for mixed).

The sheet `Notes` (first sheet, active when the workbook is opened) shows the
time of the run, the folder and file name of the workbook (Excel formulas,
so they are correct for both files), the name of the competition copy, the
PDF folder, all warnings and
information of the run, and every PDF with its location and original FIS
link. The FIS links come from `pdf_links.csv`, which the PDF downloader
writes next to the PDFs; the PDF names in row 1 of `Daten` link to them.
PDFs downloaded before this file existed have no link.

Install dependencies from this directory with:

```powershell
python -m pip install -r requirements.txt
```

Run `pdf2xlsx.bat`, or call `python pdf2xlsx.py` directly. Optional
`--pdf-dir` and `--output` arguments can be used to select another input
folder or output workbook. After a successful run, the batch file opens the
generated workbook in its associated application.
Close the output workbook in Excel before running the converter; Excel locks
open workbooks and prevents them from being overwritten.

Run the tests from the `SJpdf` directory with:

```powershell
python -m unittest discover -s pdf_to_xlsx\tests -t . -v
```
