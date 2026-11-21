"""Extract the result tables (middle part) from FIS ski jumping result PDFs.

Every value is read with its position on the page. Columns are taken from
the table header on each page, so empty cells never shift values into a
neighbouring column. Each athlete row is bounded by the table lines drawn
in the PDF. After reading, the values are checked arithmetically (judge
points, round totals, total, ranks); a PDF that fails a check raises
``ResultsParseError`` instead of producing wrong data.

Supported layouts (EWOXX and Swiss Timing):
- competition (RL, RLQ): one line per round, judge marks, total
- training / trial round (RLT, RTRIA): one or two rounds, speed/distance
  ranks on the second line, two tables side by side (EWOXX) or two rounds
  side by side (Swiss Timing)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import pymupdf
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook import Workbook

if __package__:
    from .pdf2xlsx_Daten import MONTHS, PdfParseError
else:
    from pdf2xlsx_Daten import MONTHS, PdfParseError


def looks_like_athlete(text: str) -> bool:
    """'SURNAME Firstname' (used when the PDF does not mark names bold)."""
    words = text.split()
    for index, word in enumerate(words):
        if word != word.upper() or not any(char.isalpha() for char in word):
            return index > 0 and word[0].isupper() and any(char.islower() for char in word)
    return False


ROUNDS = 2  # rounds in the fixed result columns
MAX_ROUNDS = 4  # super team: 3, ski flying championships: 4 (extra columns when used)
ROUND_FIELDS = (
    "Speed",
    "Speed Rank",
    "Distance",
    "Fall",
    "Distance Rank",
    "Distance Points",
    "A",
    "B",
    "C",
    "D",
    "E",
    "Judge Points",
    "Gate",
    "Gate From",
    "Coach Gate",
    "Gate Points",
    "Wind",
    "Wind Points",
    "Round Total",
    "Round Rank",
)
RESULT_COLUMNS = (
    "Rank",
    "Bib",
    "Name",
    "NSA",
    "U23",
    "Club",
    "Date of Birth",
    "Status",
    *(f"{name} {number}" for number in range(1, ROUNDS + 1) for name in ROUND_FIELDS),
    "Total",
    "Remarks",
)
NUMBER_FORMATS = {
    "Speed": "0.0",
    "Distance": "0.0",
    "Distance Points": "0.0",
    "A": "0.0",
    "B": "0.0",
    "C": "0.0",
    "D": "0.0",
    "E": "0.0",
    "Judge Points": "0.0",
    "Gate Points": "0.0",
    "Wind": "0.00",
    "Wind Points": "0.0",
    "Round Total": "0.0",
    "Total": "0.0",
    "Date of Birth": "yyyy-mm-dd",
}

COMPETITION_LABELS = (
    ("[km/h]", "Speed"),
    ("[m]", "Distance"),
    ("POINTS", "Distance Points"),
    ("A", "A"),
    ("B", "B"),
    ("C", "C"),
    ("D", "D"),
    ("E", "E"),
    ("POINTS", "Judge Points"),
    ("GATE", "Gate"),
    ("POINTS", "Gate Points"),
    ("[m/s]", "Wind"),
    ("POINTS", "Wind Points"),
)
COMPETITION_ROUND_LABELS = (("TOTAL", "Round Total"), ("RANK", "Round Rank"))
TRAINING_LINE1_LABELS = (
    ("[km/h]", "Speed"),
    ("[m]", "Distance"),
    ("GATE", "Gate"),
    ("[m/s]", "Wind"),
    ("POINTS", "Round Total"),
)
TRAINING_LINE2_LABELS = (
    ("RANK", "Speed Rank"),
    ("RANK", "Distance Rank"),
    ("POINTS", "Gate Points"),
    ("POINTS", "Wind Points"),
    ("RANK", "Round Rank"),
)
# Small hills without wind / gate compensation (Swiss Timing until 2023).
COMPETITION_SHORT_LABELS = COMPETITION_LABELS[:9]
TRAINING_SHORT_LINE1_LABELS = TRAINING_LINE1_LABELS[:3]
SKIPPED = "_skip"  # column without a result field (rank below the gate)
TRAINING_SHORT_LINE2_LABELS = (("RANK", "Speed Rank"), ("RANK", "Distance Rank"), ("RANK", SKIPPED))

# Swiss Timing until 2023: 'Speed', 'Dist', 'DP', ... in up to three header
# lines; renamed to the current labels on one line before the tables are read.
LEGACY_LABELS = {
    "Speed": "[km/h]", "Dist": "[m]", "DP": "POINTS", "JP": "POINTS", "Gate": "GATE",
    "Pts": "POINTS", "m/s": "[m/s]", "Points": "POINTS", "Rank": "RANK",
    **{judge: judge for judge in "ABCDE"},
}
LEGACY_LEFT_LABELS = {
    "Rk": "RANK", "Rank": "RANK", "Bib": "BIB", "Name": "NAME", "Nation": "NSA", "NSA": "NSA",
    "Club": "Club", "Birthdate": "Birthdate",
}

STATUS_CODES = ("DNS", "DSQ", "DNF", "NPS", "DQ")
SECTION_HEADINGS = (
    (re.compile(r"^Disqualified\b", re.I), "DSQ"),
    (re.compile(r"^Did not start\b", re.I), "DNS"),
    (re.compile(r"^Не стартовал", re.I), "DNS"),  # Chaikovsky
    (re.compile(r"^Дисквалифицирован", re.I), "DSQ"),
    (re.compile(r"^Did not finish\b", re.I), "DNF"),
    (re.compile(r"^Not permitted to start\b", re.I), "NPS"),
)
END_HEADINGS = re.compile(
    r"^(?:Jury\b|Competition Management\b|Competition\s*/\s*Weather\b|"
    r"Statistics\b|Legend\b|Weather Information\b|"
    r"TD$|Chief of Competition$|FIS Technical Delegate$)",  # Swiss Timing until 2023: signatures below the table
    re.I,
)
FOOTER_PATTERN = re.compile(
    r"^(?:\d{1,2} [A-Z]{3} \d{4} / |(?:Timing and )?Data Service by\b|"
    r"(?:www\.)?fis-ski\.com\b|Report Created\b|(?:Data & )?Timing Sponsor\b|(?:FIS )?Data Provider\b|Page \d+/\d+$|Page \d+ of \d+$|"
    r"data processing by\b|printed \d{1,2} [A-Z]{3} \d{4})",
    re.I,
)
# Text inside a table that belongs to no athlete: group headings of team
# lists, explanations of disqualification rules.
PREQUALIFIED = "Prequalified"
PREQUALIFIED_PATTERN = re.compile(r"^Prequalified$", re.I)  # qualification until 2016: no rank
NO_ROW_PATTERN = re.compile(
    r"^(?:Women|Men|Mixed|-{3,}|Prequalified|Reason for disqualification:?)$|^(?:SCE|ICR)\s?\d[\d.]*(?:\s*\([\d.]+\))?\s*:", re.I
)
NOT_QUALIFIED_PATTERN = re.compile(r"^Not Qualified\b", re.I)
CLASS_HEADING_PATTERN = re.compile(r"^(?:U\d{2}|OPEN) (?:GIRLS|BOYS|WOMEN|MEN|LADIES)$")  # age class
RANK_PATTERN = re.compile(r"^=?(\d+)\.$")
RANK_VALUE_PATTERN = re.compile(r"^=?(\d+)\.?$")
BIB_PATTERN = re.compile(r"^\d+$")
NSA_PATTERN = re.compile(r"^[A-Z]{3}$")
DOB_PATTERN = re.compile(r"^(\d{1,2}) ([A-Z]{3}) (\d{4})$")
NUMBER_PATTERN = re.compile(r"^-?\d+(?:\.\d+)?$")
# '*106.5': fall; '{117.0' (Swiss Timing 2024, meaning not printed): the distance
DISTANCE_PATTERN = re.compile(r"^(\*)?\{?(\d+(?:\.\d+)?)$")
GATE_PATTERN = re.compile(r"^(©)?(\d+)(?:(?:[➔→]|->)(©)?(\d+))?$")
# '©11/13': gate 11 (requested by the coach), gate before 13
GATE_SLASH_PATTERN = re.compile(r"^(©)?(\d+)/(©)?(\d+)$")
TEAM_BIB_PATTERN = re.compile(r"^\d+ ?- ?\d$")  # '8-1', '11 - 1'
ROUND_NOTE_PATTERN = re.compile(r"^(?:(?:\d(?:st|nd|rd|th)|Final) Rnd|Rnd \d)$")  # super team, total column (2023: 'Rnd 1')
TEAM_NAME_PATTERN = re.compile(r"^([A-Z]{3})?\s*-\s+[A-Z]")  # 'GER - Germany'; mixed team '- SUI/ITA'
GENDER_SUFFIX = re.compile(r"\s*\((?:W|M)\)$")


def _is_team_name(text: str) -> bool:
    """'GER - Germany', '- SUI/ITA' or the country alone in capitals ('AUSTRIA')."""
    return bool(TEAM_NAME_PATTERN.match(text)) or _team_country(text) is not None


def _team_country(text: str) -> str | None:
    """Nation of a team printed as its country in capitals: 'AUSTRIA', 'SLOVENIA II'."""
    country = re.sub(r" (?:I{1,3}|IV|\d)$", "", text)
    if len(country) > 3 and country.isupper():  # not the nation code 'USA' of a jumper
        return COUNTRY_CODES.get(country.casefold())
    return None


BIB_NAME_PATTERN = re.compile(r"^(\d{1,3})\s+(?=\D)")
THOUSANDS_PATTERN = re.compile(r"^\d{1,3}(?:\.\d{3})+\.\d$")  # team total '1.038.5'
NAME_TEXT_PATTERN = re.compile(r"^[^\W\d_][^\W\d_'’.\- ]*(?:['’.\- ]+[^\W\d_]+)*\.?$")  # letters only
TEAM_KEY = "_team"  # team row (not a column): its members follow below
# Numbers of neighbouring columns in one span; Olympic training: distance and gate.
SPLIT_NUMBER_PATTERN = re.compile(r"^-?\d+\.\d+$|^-?\d{1,3}$|^©?\d{1,2}[➔→]©?\d{1,2}$")
SCE_PATTERN = re.compile(r"^(?:SCE|ICR)\s?\d")  # disqualification (rule code, ICR since 2025) inside a row
LUCKY_LOSER = "Lucky Loser"
# KO competition (Four Hills): round 1 in pairs, the winners and the best
# losers are ranked before the other losers.
KO_PATTERN = re.compile(r"^(?=.*\bHill\b).*\bKO\b", re.M)  # 'Men Individual KO Large Hill'
TOLERANCE = 0.051


class ResultsParseError(PdfParseError):
    """Raised when the result table cannot be read without doubt."""


@dataclass(frozen=True)
class Span:
    text: str
    x0: float
    x1: float
    y0: float
    y1: float
    bold: bool

    @property
    def x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def y(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Column:
    field: str
    low: float
    high: float


@dataclass
class Table:
    layout: str  # "competition" or "training"
    left: float
    right: float
    header_y: float
    data_low: float
    line_columns: list[list[Column]]  # competition: one list for all lines
    total_low: float | None = None
    remarks_low: float | None = None
    bottom: float | None = None  # header of the next table below on the page


@dataclass
class _State:
    pdf_name: str
    not_qualified: bool = False
    prequalified: bool = False  # rows below 'Prequalified' (qualification until 2016)
    rows: list[dict[str, Any]] = field(default_factory=list)
    section_rows: list[dict[str, Any]] = field(default_factory=list)
    layouts: set[str] = field(default_factory=set)
    team: dict[str, Any] | None = None  # last team row (team events)
    wind_column: bool = True  # False: small hill lists without wind / gate
    blank_is_zero: bool = False  # WinSpringen: no distance / wind points printed means 0
    fastreport: bool = False  # Chaikovsky: decimal comma, labels without brackets
    gate_points_missing: bool = False  # gate printed without its points: round totals not recalculated
    class_group: str | None = None  # class heading above the rows ('U15 GIRLS'), continued on the next page
    dense_ranks: bool = False  # WinSpringen training: equal values share a rank, no gap
    rank_groups: dict[str, list[dict[str, Any]]] = field(default_factory=dict)  # ranked apart


def _fail(state: _State, page: int, message: str) -> ResultsParseError:
    return ResultsParseError(f"{state.pdf_name} page {page}: {message}")


def _page_spans(page: pymupdf.Page) -> list[Span]:
    spans = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                text = " ".join(span["text"].split())
                if not text:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                bold = bool(span["flags"] & 16)
                words = text.split(" ")
                if len(words) > 1 and all(SPLIT_NUMBER_PATTERN.match(word) for word in words):
                    # Numbers of neighbouring columns in one span ('5.3 -0.01'):
                    # one span per number, placed by its share of the characters.
                    width = (x1 - x0) / len(text)
                    position = 0
                    for word in words:
                        start = x0 + position * width
                        spans.append(Span(word, start, start + len(word) * width, y0, y1, bold))
                        position += len(word) + 1
                    continue
                spans.append(Span(text, x0, x1, y0, y1, bold))
    # Some PDFs print a value twice at the same place (shadow text).
    unique: dict[tuple, Span] = {}
    for span in spans:
        key = (span.text, round(span.x0), round(span.y0))
        if key not in unique or span.bold:
            unique[key] = span
    return list(unique.values())


HEADER_WORDS = {"[KM/H]", "[M]", "POINTS", "GATE", "[M/S]", "TOTAL", "RANK"}


def _split_labels(spans: list[Span]) -> list[Span]:
    """Header labels printed as one text ('[m/s] Points', 'Gate Points',
    EYOF 2023, Universiade 2017): one span per label, placed by its letters."""
    result = []
    for span in spans:
        if span.text == "GatePoints":  # EYOF 2017: without space
            span = Span("Gate Points", span.x0, span.x1, span.y0, span.y1, span.bold)
        words = span.text.split()
        mixed = re.fullmatch(r"(\d+)\s+(- [A-Z].*)", span.text)
        if mixed:
            # Bib and mixed team in one text: '1 - SUI/ITA' (Tschagguns 2025).
            width = (span.x1 - span.x0) / max(len(span.text), 1)
            split = span.x0 + mixed.start(2) * width
            result.append(Span(mixed.group(1), span.x0, span.x0 + len(mixed.group(1)) * width,
                               span.y0, span.y1, span.bold))
            result.append(Span(mixed.group(2), split, span.x1, span.y0, span.y1, span.bold))
            continue
        if len(words) < 2 or not all(word.upper() in HEADER_WORDS for word in words):
            result.append(span)
            continue
        width = (span.x1 - span.x0) / max(len(span.text), 1)
        position = 0
        for word in words:
            start = span.text.index(word, position)
            position = start + len(word)
            result.append(Span(word, span.x0 + start * width, span.x0 + position * width,
                               span.y0, span.y1, span.bold))
    return result


def _legacy_labels(spans: list[Span]) -> list[Span]:
    """Swiss Timing until 2023: rename the header labels to the current ones.

    Competition: 'Speed Dist DP A B C D E JP' (gate / wind 'Gate Pts m/s Pts'
    a few points lower), round total and round rank below the label
    'Round' / 'Points', 'Total'. Training: 'Speed Dist Gate m/s Points' per
    round, the ranks one line lower. The labels of a table are moved onto
    the line of its 'Speed' label; 'Total' stays a little higher (own column
    outside the round columns).
    """
    if any(s.text == "[km/h]" for s in spans):
        return spans
    speeds = sorted(
        (s for s in spans if s.text == "Speed" and any(
            other.text == "Dist" and abs(other.y - s.y) < 3 and 0 < other.x0 - s.x1 < 30 for other in spans
        )),
        key=lambda s: (s.y, s.x0),
    )
    bands: list[list[Span]] = []
    for speed in speeds:
        if bands and abs(speed.y - bands[-1][0].y) < 3:
            bands[-1].append(speed)
        else:
            bands.append([speed])
    if not bands:
        return spans
    replaced: dict[int, list[Span]] = {}

    def put(span: Span, text: str, y: float, x0: float | None = None, x1: float | None = None) -> None:
        moved = Span(text, span.x0 if x0 is None else x0, span.x1 if x1 is None else x1,
                     y - (span.y1 - span.y0) / 2, y + (span.y1 - span.y0) / 2, True)
        replaced.setdefault(id(span), []).append(moved)

    for band in bands:
        band_y = band[0].y
        first = band[0].x0
        near = [s for s in spans if band_y - 13 < s.y < band_y + 14 and id(s) not in replaced]
        training = any(
            s.text == "Rank" and s.x0 >= first - 5 and band_y + 8 < s.y < band_y + 14 for s in near
        )
        if not training:
            for span in near:
                below = next((s for s in near if s.text == "Points" and 0 < s.y - span.y < 12
                              and abs(s.x0 - span.x0) < 5), None) if span.text == "Round" else None
                if below is not None:
                    # One label above two columns: round total, round rank.
                    put(span, "POINTS", band_y, below.x0 - 4, below.x1 - 10)
                    put(below, "RANK", band_y, below.x1 - 4, below.x1 + 6)
        for span in near:
            if id(span) in replaced:
                continue
            if span.x1 <= first:
                if span.text in LEGACY_LEFT_LABELS and span.y < band_y + 14:
                    put(span, LEGACY_LEFT_LABELS[span.text], band_y)
                continue
            if training:
                if span.text in LEGACY_LABELS and abs(span.y - band_y) < 3:
                    put(span, LEGACY_LABELS[span.text], band_y)
                elif span.text in ("Rank", "Points") and band_y + 8 < span.y < band_y + 14:
                    put(span, LEGACY_LABELS[span.text], span.y)
                continue
            if span.text == "Grp Rk" and abs(span.y - band_y) < 6:
                replaced[id(span)] = []  # team after round 1: rank in the group, below the points
            elif span.text == "Total" and abs(span.y - band_y) < 6:
                put(span, "TOTAL", band_y - 5)
            elif span.text in LEGACY_LABELS and abs(span.y - band_y) < 6:
                put(span, LEGACY_LABELS[span.text], band_y)
    result = []
    for span in spans:
        result.extend(replaced.get(id(span), [span]))
    return result


def _vertical_segments(page: pymupdf.Page) -> list[tuple[float, float, float]]:
    segments = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                start, end = item[1], item[2]
                if abs(start.x - end.x) < 0.6:
                    segments.append((start.x, min(start.y, end.y), max(start.y, end.y)))
            elif item[0] == "re":
                rect = item[1]
                if rect.width < 1.6:
                    segments.append((rect.x0, rect.y0, rect.y1))
    return segments


def _columns(labels: list[tuple[Span, str]], low: float, high: float) -> list[Column]:
    columns = []
    for index, (span, field_name) in enumerate(labels):
        column_low = low if index == 0 else (labels[index - 1][0].x1 + span.x0) / 2
        column_high = (
            high if index == len(labels) - 1 else (span.x1 + labels[index + 1][0].x0) / 2
        )
        columns.append(Column(field_name, column_low, column_high))
    return columns


def _same_row(
    spans: Iterable[Span], y: float, left: float, right: float, tolerance: float = 1.6
) -> list[Span]:
    return sorted(
        (span for span in spans if abs(span.y - y) < tolerance and left <= span.x0 < right),
        key=lambda span: span.x0,
    )


def _match_labels(
    row: list[Span], expected: tuple[tuple[str, str], ...], repeat: int
) -> list[tuple[Span, str]] | None:
    texts = ["POINTS" if span.text.upper() == "PTS" else span.text.upper() for span in row]  # Rasnov 2023
    expected_texts = [label.upper() for label, _ in expected] * repeat
    if texts != expected_texts:
        return None
    result = []
    for index, span in enumerate(row):
        field_name = expected[index % len(expected)][1]
        group = index // len(expected) + 1
        result.append((span, f"{field_name}#{group}"))
    return result


def _find_tables(spans: list[Span], page_width: float, state: _State, page: int) -> list[Table]:
    """Tables of a page; several header rows one below the other (team
    trial round: one block per group) give tables one below the other."""
    speed_labels = sorted((s for s in spans if s.text == "[km/h]"), key=lambda s: (s.y, s.x0))
    bands: list[list[Span]] = []
    for label in speed_labels:
        if bands and abs(label.y - bands[-1][0].y) <= 1.6:
            bands[-1].append(label)
        elif bands and label.y - bands[-1][0].y < 60:
            raise _fail(state, page, "table headers on different heights.")
        else:
            bands.append([label])
    tables = []
    for index, band in enumerate(bands):
        band_tables = _find_band_tables(spans, sorted(band, key=lambda s: s.x0), page_width, state, page)
        if index + 1 < len(bands):
            for table in band_tables:
                table.bottom = bands[index + 1][0].y - 15
        tables.extend(band_tables)
    return tables


def _find_band_tables(
    spans: list[Span], speed_labels: list[Span], page_width: float, state: _State, page: int
) -> list[Table]:
    header_y = speed_labels[0].y

    # Each table starts at its RANK/BIB header left of the first speed column.
    labels = [
        s for s in spans
        if re.match(r"^(RANK|BIB)\b", s.text, re.I) and header_y - 25 < s.y < header_y + 3
    ]
    starts = [s for s in labels if s.bold]
    if not starts:
        # Header not printed bold: RANK/BIB labels next to a NAME label.
        names = [s for s in spans if s.text.upper() == "NAME" and header_y - 25 < s.y < header_y + 3]
        starts = [
            s for s in labels
            if any(abs(name.y - s.y) < 6 and 0 < name.x0 - s.x1 < 60 for name in names)
            or any(abs(other.y - s.y) < 6 and 0 < other.x0 - s.x1 < 30 for other in labels)
        ]
    starts.sort(key=lambda s: s.x0)
    # Several header labels of one table may start with RANK/BIB; a new
    # table only starts where a RANK/BIB label follows a speed column.
    lefts: list[float] = []
    for start in starts:
        if not lefts or (
            any(lefts[-1] < speed.x0 < start.x0 for speed in speed_labels)
            and any(speed.x0 > start.x0 for speed in speed_labels)
        ):
            lefts.append(start.x0)
    if not lefts or lefts[0] > speed_labels[0].x0:
        raise _fail(state, page, "cannot find the left edge of a result table.")
    tables = []
    for index, left in enumerate(lefts):
        # Right-aligned bibs of the next table start a little left of its label.
        right = lefts[index + 1] - 10 if index + 1 < len(lefts) else page_width
        speed = next((s for s in speed_labels if left < s.x0 < right), None)
        if speed is None:
            raise _fail(state, page, "result table without speed column.")
        tables.append(_build_table(spans, speed, left, right, state, page))
    return tables


def _build_table(
    spans: list[Span], speed: Span, left: float, right: float, state: _State, page: int
) -> Table:
    header_y = speed.y
    # Olympic layout: the labels of one header row differ by 2pt in height.
    row = [s for s in _same_row(spans, header_y, left, right, 2.6) if s.x0 >= speed.x0 - 1]
    # Swiss Timing training: the POINTS labels sit one header row higher.
    for span in spans:
        if (
            span.text.upper() == "POINTS"
            and header_y - 12 < span.y < header_y - 3
            and left <= span.x0 < right
            and not any(other.x0 < span.x1 and span.x0 < other.x1 for other in row)
        ):
            row.append(span)
    row.sort(key=lambda s: s.x0)
    texts = [s.text for s in row]
    data_low = speed.x0 - 12
    # The nation column may end close to the speed column (two tables side by side).
    nation = [
        label for label in spans
        if label.text.upper() in ("NSA", "NOC", "CODE", "NAT") and header_y - 25 < label.y < header_y + 3
        and left <= label.x0 and label.x1 < speed.x0
    ]
    if nation:
        data_low = max(data_low, max(label.x1 for label in nation) + 1)

    labels_count = len(COMPETITION_LABELS)
    competition = _match_labels(row[:labels_count], COMPETITION_LABELS, 1)
    if competition is None and not any(s.text.upper() == "GATE" for s in row):
        # Small hill without gate / wind compensation.
        labels_count = len(COMPETITION_SHORT_LABELS)
        competition = _match_labels(row[:labels_count], COMPETITION_SHORT_LABELS, 1)
        if competition is not None:
            state.wind_column = False
    if competition is not None:
        rest = row[labels_count:]
        if rest and rest[-1].text == "Time Diff.":
            # Nordic mixed team (YOG 2016): time difference for the cross-country start.
            rest = rest[:-1]
            row = row[:-1]
        has_total = any(
            s.text.upper() in ("TOTAL", "TOTAL /") and header_y - 25 < s.y < header_y + 3 and left <= s.x0 < right
            for s in spans
        )
        if not has_total and rest and rest[-1].text.upper() == "POINTS" and rest[-1].y < header_y - 1:
            # KO qualification (Four Hills): the total column is called POINTS,
            # printed one header row higher.
            total = [rest.pop()]
            row = row[: labels_count + len(rest)]
        else:
            total = []
        if rest:
            # Round total with or without round rank; team events: POINTS.
            labels = COMPETITION_ROUND_LABELS[: len(rest)]
            if rest[0].text.upper() == "POINTS":
                labels = (("POINTS", "Round Total"), *labels[1:])
            elif len(rest) == 1 and rest[0].text.upper() == "RANK":
                labels = (("RANK", "Round Rank"),)  # small hill team: 'Group Rank', points in the total column
            extra = _match_labels(rest, labels, 1)
            if extra is None:
                raise _fail(state, page, f"unknown result table header {texts}.")
            competition += extra
        total = total or [
            s for s in spans
            if s.text.upper() in ("TOTAL", "TOTAL /")  # YOG 2016: "Total / Time Diff."
            and header_y - 25 < s.y < header_y + 3
            and s.x0 > row[-1].x1 + 2
            and left <= s.x0 < right
        ]
        if len(total) != 1:
            raise _fail(state, page, "cannot find the TOTAL column.")
        total_low = (row[-1].x1 + total[0].x0) / 2
        columns = _columns(competition, data_low, total_low)
        return Table("competition", left, right, header_y, data_low, [columns], total_low=total_low)

    line1_labels, line2_labels = TRAINING_LINE1_LABELS, TRAINING_LINE2_LABELS
    groups = len(row) // len(line1_labels)
    line1 = _match_labels(row, line1_labels, groups) if groups else None
    if line1 is None and len(row) % len(TRAINING_SHORT_LINE1_LABELS) == 0:
        # Small hill without wind / gate points: speed, distance, gate.
        line1_labels, line2_labels = TRAINING_SHORT_LINE1_LABELS, TRAINING_SHORT_LINE2_LABELS
        groups = len(row) // len(line1_labels)
        line1 = _match_labels(row, line1_labels, groups)
    if line1 is None:
        raise _fail(state, page, f"unknown result table header {texts}.")
    rank_row_y = next(
        (s.y for s in spans if s.text.upper() == "RANK" and header_y + 3 < s.y < header_y + 14
         and speed.x0 - 5 <= s.x0 < right),
        None,
    )
    rank_row = (
        [s for s in _same_row(spans, rank_row_y, left, right) if s.x0 >= speed.x0 - 5]
        if rank_row_y is not None
        else []
    )
    line2 = _match_labels(rank_row, line2_labels, groups)
    if line2 is None and rank_row:
        # Labels of the second line misplaced in the PDF: its values stand
        # below the first line labels (speed rank below speed, ...).
        line2 = [
            (span, f"{line2_labels[index % len(line2_labels)][1]}#{group}")
            for index, (span, field_group) in enumerate(line1)
            for group in [field_group.split("#")[1]]
        ]
    if line2 is None:
        raise _fail(state, page, f"unknown training header {[s.text for s in rank_row]}.")
    remarks = [
        s for s in spans
        if s.text.upper() == "REMARKS" and header_y - 25 < s.y < header_y + 3 and left <= s.x0 < right
    ]
    high = remarks[0].x0 - 3 if remarks else right
    return Table(
        "training",
        left,
        right,
        rank_row_y,
        data_low,
        [_columns(line1, data_low, high), _columns(line2, data_low, high)],
        remarks_low=remarks[0].x0 - 3 if remarks else None,
    )


def _row_cells(
    page: pymupdf.Page, table: Table, y_end: float, team: bool = False
) -> list[tuple[float, float]]:
    """Return athlete rows from the vertical lines at the table's left edge.

    Team events: the team row starts at the left edge, the rows of its
    jumpers a little further right (at the bib column); a row ends where the
    next one starts.
    """
    tops: list[tuple[float, float]] = []
    segments = [
        (x, top, bottom) for x, top, bottom in _vertical_segments(page)
        if table.left - 16 <= x <= table.left + (30 if team else 1) and table.header_y + 2 < top < y_end
    ]
    # Rows are more than 10 high; in small print (YOG 2020) the rows are 9.4.
    heights = sorted(bottom - top for x, top, bottom in segments if x <= table.left + 1)
    # (Trial rounds: half rows of 9.9 between rows of 19.8 are no rows.)
    row_height = 8 if heights and heights[len(heights) // 2] < 9.6 else 10
    for x, top, bottom in segments:
        # Team jumper without jump (DNS): a row of one line only.
        if bottom - top > (row_height if x <= table.left + 1 else 6):
            tops.append((top, min(bottom, y_end)))
    tops.sort()
    cells: list[tuple[float, float]] = []
    for top, bottom in tops:
        if cells and abs(cells[-1][0] - top) < 2:
            cells[-1] = (cells[-1][0], max(cells[-1][1], bottom))
        else:
            cells.append((top, bottom))
    return [
        (top, min(bottom, cells[index + 1][0]) if index + 1 < len(cells) else bottom)
        for index, (top, bottom) in enumerate(cells)
    ]


def _team_cells(
    cells: list[tuple[float, float]], spans: list[Span], table: Table, super_team: bool = False
) -> list[tuple[float, float]]:
    """Team events: one row per team and per jumper, starting at the row
    line just above its team name or bib (super team: a jumper has one line
    per round, each with its own row lines)."""
    anchors = sorted(
        span.y for span in spans
        if span.x < table.data_low
        and (_is_team_name(span.text) or TEAM_BIB_PATTERN.match(span.text))
    )
    if not anchors:
        return cells
    firsts = sorted(s.y for s in spans if re.fullmatch(r"1st Rnd|Rnd 1", s.text, re.I))
    value_ys = [s.y for s in spans if s.x >= table.data_low and NUMBER_PATTERN.match(s.text)]
    members = [s.y for s in spans if s.x < table.data_low and TEAM_BIB_PATTERN.match(s.text)]
    # Round 1 line 8 above the name (teams out of the final round: two lines
    # 4 above and below the name).
    three_lines = super_team and any(
        any(-11 < y - member < -5 for y in value_ys) and any(abs(y - member) <= 2 for y in value_ys)
        and any(5 < y - member < 11 for y in value_ys)
        for member in members
    ) and not any(
        re.search(r"-\s*[3-9]$", s.text) for s in spans if s.x < table.data_low and TEAM_BIB_PATTERN.match(s.text)
    )  # Super Team: two jumpers per team ('8-1', '8-2'); team events have four
    if super_team and (firsts or three_lines) and cells:
        # Super Team 2026: lines '1st Rnd' (above the name), '2nd Rnd', '3rd Rnd'
        # per jumper; the team line has its round points above and below.
        tops = []
        for span in sorted((s for s in spans if s.x < table.data_low
                            and (_is_team_name(s.text) or TEAM_BIB_PATTERN.match(s.text))),
                           key=lambda s: s.y):
            first = max((y for y in firsts if span.y - 12 < y <= span.y + 1), default=None)
            if TEAM_BIB_PATTERN.match(span.text) and first is not None:
                tops.append(first - 2)
            elif TEAM_BIB_PATTERN.match(span.text):
                above = any(-11 < y - span.y < -5 for y in value_ys)
                tops.append(span.y - (12 if above else 6))
            else:
                tops.append(span.y - 16)  # team: round points 10 (2025: 14) above and below
        bottom = max(bottom for _top, bottom in cells)
        if len(tops) > 1 and bottom < tops[-1]:
            # Row lines around the team rows only (Rasnov 2022): the last jumper
            # ends one row height below its top.
            bottom = tops[-1] + max(below - above for above, below in zip(tops, tops[1:]))
        return list(zip(tops, [*tops[1:], bottom]))
    first_value = min(
        (s.y for s in spans if s.x >= table.data_low and NUMBER_PATTERN.match(s.text)), default=None
    )
    names_above = any(
        s.x < table.data_low and looks_like_athlete(s.text)
        and any(2 < anchor - s.y < 8 for anchor in anchors)
        and not any(abs(anchor - s.y) <= 2 for anchor in anchors)
        and not any(t.x >= table.data_low and abs(t.y - s.y) < 2 for t in spans)  # own line
        for s in spans
    )
    # (a page that starts with the rounds of a jumper of the page before has
    # values far above the first name: not this layout)
    # Mixed team 2023 (Zakopane): round 1 line just above each jumper's name.
    # (the first jumper of a team: in other layouts the team line is above it)
    member_ys = [s.y for s in spans if s.x < table.data_low and re.fullmatch(r"\d+ ?- ?1", s.text)]
    # (teams out of round 2: one line only)
    rounds_above = bool(member_ys) and sum(
        any(-8 < s.y - member < -3 for s in spans if s.x >= table.data_low and NUMBER_PATTERN.match(s.text))
        for member in member_ys
    ) * 2 >= len(member_ys)
    if (first_value is not None and anchors[0] - 8 < first_value < anchors[0] - 2) or names_above or rounds_above:
        # Olympic layout: the round lines stand above and below the name, the
        # row border lies halfway between two names.
        borders = [(above + below) / 2 for above, below in zip(anchors, anchors[1:])]
        top = min(top for top, _bottom in cells)
        bottom = max(bottom for _top, bottom in cells)
        if borders and bottom < borders[-1]:
            # Row lines around the team rows only (Rasnov 2022): the last jumper
            # ends one row height below its name.
            bottom = anchors[-1] + max(below - above for above, below in zip(anchors, anchors[1:]))
        return list(zip([top, *borders], [*borders, bottom]))
    tops = sorted({top for top, _bottom in cells})
    kept: list[float] = []
    for anchor in anchors:
        top = max((top for top in tops if top <= anchor + 1), default=None)
        if top is not None and top not in kept:
            kept.append(top)
    if not kept:
        return cells
    bottom = max(bottom for _top, bottom in cells)
    return list(zip(kept, [*kept[1:], bottom]))


def _name_cells(spans: list[Span], table: Table, y_end: float) -> list[tuple[float, float]]:
    """Athlete rows of a table without row lines: one row per athlete name,
    the border halfway between two names."""
    names: list[float] = []
    for span in sorted(spans, key=lambda s: s.y):
        if span.x < table.data_low and (looks_like_athlete(span.text) or (
                BIB_NAME_PATTERN.match(span.text) and _looks_like_short_name(BIB_NAME_PATTERN.sub("", span.text)))
                                        or _is_team_name(span.text)
                                        # 'BROWN A J' with the nation next to it
                                        or (_looks_like_short_name(span.text) and any(
                                            abs(o.y - span.y) < 2 and NSA_PATTERN.match(o.text) for o in spans))):
            if any(abs(other.y - span.y) < 2 and other.x < table.data_low and _date(other.text) for other in spans):
                continue  # club with the date of birth ('TJ Roznov pod Radhostem', '30 MAY 1997')
            if not names or span.y - names[-1] > 3:
                names.append(span.y)
    borders = [(above + below) / 2 for above, below in zip(names, names[1:])]
    tops = [table.header_y + 3, *borders]
    bottoms = [*borders, y_end]
    return list(zip(tops, bottoms))


def _date(text: str) -> date | None:
    match = DOB_PATTERN.match(text)
    if not match:
        return None
    month = MONTHS.get(match.group(2))
    if month is None:
        return None
    return date(int(match.group(3)), month, int(match.group(1)))


def _looks_like_short_name(text: str) -> bool:
    """'KOBAYASHI Ryoyu', or with initials only (Olympics 2022: 'TROFIMOV RS')."""
    return looks_like_athlete(text) or bool(re.fullmatch(r"[A-Z][A-Z'’\-]+(?: [A-Z][A-Z'’\-]+)*(?: [A-Z]{1,3})+", text))


def _split_bib_name(spans: list[Span]) -> list[Span]:
    """Olympics 2022: bib and name in one text ('49 KOBAYASHI Ryoyu')."""
    result = []
    for span in spans:
        match = BIB_NAME_PATTERN.match(span.text)
        if match and _looks_like_short_name(span.text[match.end():]):
            width = (span.x1 - span.x0) / max(len(span.text), 1)
            split = span.x0 + width * len(match.group(0))
            result.append(Span(match.group(1), span.x0, split - width, span.y0, span.y1, span.bold))
            result.append(Span(span.text[match.end():], split, span.x1, span.y0, span.y1, span.bold))
        else:
            result.append(span)
    return result


def _identity(spans: list[Span]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    spans = _split_bib_name(spans)
    team = next((s for s in spans if _is_team_name(s.text)), None)
    if team is not None:
        # Team row: 'GER - Germany', rank, team bib; the jumpers follow.
        values.update(Name=team.text.lstrip("- "), **{TEAM_KEY: True})
        match = TEAM_NAME_PATTERN.match(team.text)
        if match is None:
            values["NSA"] = _team_country(team.text)  # 'AUSTRIA' (Zakopane 2022)
        elif match.group(1):
            values["NSA"] = match.group(1)  # mixed team: none
        for span in spans:
            if RANK_PATTERN.match(span.text):
                values["Rank"] = int(RANK_PATTERN.match(span.text).group(1))
            elif BIB_PATTERN.match(span.text):
                values["Bib"] = int(span.text)
        return values
    # Team events: 'not qualified for 2nd round' next to the jumper.
    spans = [s for s in spans if not NOT_QUALIFIED_PATTERN.match(s.text)]
    name_parts = []
    club_parts = []
    nsa = [s for s in spans if s.bold and NSA_PATTERN.match(s.text)]
    # Without bold print: the name is 'SURNAME Firstname', the nation the
    # rightmost three-letter code. A bold rank alone does not count.
    plain = not any(s.bold and not RANK_PATTERN.match(s.text) for s in spans)
    if plain:
        nsa = [s for s in spans if NSA_PATTERN.match(s.text)]
    nsa_span = max(nsa, key=lambda s: s.x0) if nsa else None
    name_spans: list[Span] = []
    if plain:
        ordered = sorted(spans, key=lambda s: (s.y, s.x0))
        name_spans = [s for s in ordered if looks_like_athlete(s.text)][:1]
        numbered = any(
            BIB_PATTERN.match(s.text) or TEAM_BIB_PATTERN.match(s.text) or RANK_PATTERN.match(s.text)
            for s in spans
        )
        if not name_spans and numbered:
            # Olympic layout: 'DESCHWANDEN' 'Gregor' in two spans, initials
            # only ('SUNDAL KE'): the text spans on the line of the first one.
            words = [
                s for s in ordered
                if NAME_TEXT_PATTERN.match(s.text) and s is not nsa_span and len(s.text) > 1
            ]
            name_spans = [s for s in words if words and abs(s.y - words[0].y) < 2]
            following = [s for s in words if s not in name_spans]
            if (
                len(name_spans) == 1 and name_spans[0].text.isupper() and following
                and looks_like_athlete(f"{name_spans[0].text} {following[0].text}")
            ):
                name_spans.append(following[0])  # first name on the next line
    numbers: list[Span] = []
    for span in sorted(spans, key=lambda s: (round(s.y), s.x0)):
        if plain and span in name_spans:
            name_parts.append(span.text)
            continue
        if span is nsa_span:
            values["NSA"] = span.text
        elif span.text == "U23":
            values["U23"] = "U23"
        elif span.text == "*":
            values["Remarks"] = LUCKY_LOSER  # KO competition: best loser of a pair
        elif re.match(r"^\*\s?\d+$", span.text):
            values["Remarks"] = LUCKY_LOSER  # 2026: '*20', star and bib in one span
            values["Bib"] = int(span.text[1:])
        elif RANK_PATTERN.match(span.text):
            values["Rank"] = int(RANK_PATTERN.match(span.text).group(1))
        elif BIB_PATTERN.match(span.text):
            values["Bib"] = int(span.text)
            numbers.append(span)
        elif TEAM_BIB_PATTERN.match(span.text):
            values["Bib"] = span.text.replace(" ", "")  # team member '8-1': team 8, first jumper
        elif _date(span.text):
            values["Date of Birth"] = _date(span.text)
        elif span.bold and not plain:
            name_parts.append(span.text)
        else:
            club_parts.append(span.text)
    if len(numbers) == 2 and "Rank" not in values and abs(numbers[0].y - numbers[1].y) < 3:
        # GP Women 2023: rank printed without point, left of the bib.
        rank, bib = sorted(numbers, key=lambda s: s.x0)
        values["Rank"], values["Bib"] = int(rank.text), int(bib.text)
    if name_parts:
        # Mixed team: 'FREITAG Selina (W)'
        values["Name"] = GENDER_SUFFIX.sub("", " ".join(name_parts))
        nation = re.fullmatch(r"(.*[a-z].*) ([A-Z]{3})", values["Name"])
        if nation and "NSA" not in values:
            values["Name"], values["NSA"] = nation.groups()  # 'KALINICHENKO Vitaliy UKR' in one text
    if club_parts:
        values["Club"] = " ".join(club_parts)
    return values


def _parse_value(field_name: str, text: str) -> dict[str, Any] | None:
    if field_name == "Distance":
        match = DISTANCE_PATTERN.match(text)
        if not match:
            return None
        values: dict[str, Any] = {"Distance": float(match.group(2))}
        if match.group(1):
            values["Fall"] = "*"
        return values
    if field_name == "Gate":
        match = GATE_PATTERN.match(text.replace(" ", ""))
        slash = GATE_SLASH_PATTERN.match(text.replace(" ", ""))
        if slash:
            coach_after, after, coach_before, before = slash.groups()
        elif match:
            coach_before, before, coach_after, after = match.groups()
        else:
            return None
        if after is None:
            return {"Gate": int(before), "Coach Gate": "©" if coach_before else None}
        return {
            "Gate": int(after),
            "Gate From": int(before),
            "Coach Gate": "©" if (coach_before or coach_after) else None,
        }
    if field_name.endswith("Rank"):
        if text in ("=.", "."):
            return {field_name: None}  # not jumped (trial round, Sapporo 2026)
        if re.fullmatch(r"f\d+\.", text):
            # Swiss Timing until 2020: 'f33.' qualified for round 2 by rule 4.2 (fall)
            return {field_name: int(text[1:-1]), "Remarks": "f: add. qualified by rule 4.2 (Fall)"}
        match = RANK_VALUE_PATTERN.match(text)
        return {field_name: int(match.group(1))} if match else None
    if NUMBER_PATTERN.match(text):
        return {field_name: float(text)}
    return None


def _lines(spans: list[Span]) -> list[list[Span]]:
    lines: list[list[Span]] = []
    for span in sorted(spans, key=lambda s: s.y):
        if lines and span.y - lines[-1][-1].y <= 2.5:
            lines[-1].append(span)
        else:
            lines.append([span])
    return lines


def _parse_block(
    spans: list[Span], table: Table, state: _State, page: int
) -> dict[str, Any] | None:
    spans = [s for s in spans if not re.fullmatch(r"-{3,}", s.text)]  # separator between teams
    identity_spans = [s for s in spans if s.x < table.data_low]
    if table.total_low is not None:
        right_spans = [s for s in spans if s.x >= table.total_low]
    else:
        right_spans = []
    remark_spans = (
        [s for s in spans if s.x0 >= table.remarks_low] if table.remarks_low is not None else []
    )
    disqualified = [s for s in spans if SCE_PATTERN.match(s.text) and s not in identity_spans]
    # Separator 'Not Qualified for Final Round' printed across the data columns.
    spans = [s for s in spans if s in identity_spans or not NOT_QUALIFIED_PATTERN.match(s.text)]
    data_spans = [
        s for s in spans
        if s not in identity_spans and s not in right_spans and s not in remark_spans
        and s not in disqualified
    ]
    row = _identity(identity_spans)
    if NOT_QUALIFIED_PATTERN.match(row.get("Name", "")) and "Bib" not in row:
        return None  # separator line printed bold inside a row cell
    if "Name" not in row:
        if any(NOT_QUALIFIED_PATTERN.match(s.text) for s in spans):
            return None
        if not spans:
            return None
        if state.rows and all(s.x < table.data_low and not NUMBER_PATTERN.match(s.text) for s in spans):
            # Club line below the jumper, cut off by a row line (team lists).
            previous = state.rows[-1]
            previous["Club"] = " ".join(filter(None, [previous.get("Club"), *(s.text for s in spans)]))
            return None
        raise _fail(state, page, f"row without athlete name: {[s.text for s in spans]}.")
    label = row["Name"]

    if right_spans and row.get(TEAM_KEY):
        # Team row: the rank may be printed next to the total ('1.', '376.3').
        right_spans = [s for s in right_spans if not RANK_PATTERN.match(s.text)]
    elif right_spans and isinstance(row.get("Bib"), str) and len(right_spans) == 2:
        # Team member after round 1 (Swiss Timing until 2023): group rank below the points.
        right_spans = [s for s in right_spans if not RANK_PATTERN.match(s.text)]
    right_spans = [s for s in right_spans if not ROUND_NOTE_PATTERN.match(s.text)
                   and not re.fullmatch(r"\+?\d+:\d\d", s.text)]  # YOG 2016: time difference '0:00'
    right_spans = [
        Span(s.text.replace(".", "", s.text.count(".") - 1), s.x0, s.x1, s.y0, s.y1, s.bold)
        if THOUSANDS_PATTERN.match(s.text) else s for s in right_spans
    ]
    if right_spans:
        if len(right_spans) != 1 or not NUMBER_PATTERN.match(right_spans[0].text):
            raise _fail(state, page, f"{label}: unreadable total {[s.text for s in right_spans]}.")
        row["Total"] = float(right_spans[0].text)
    if remark_spans or disqualified:
        row["Remarks"] = " ".join(
            filter(None, [row.get("Remarks")] + [
                s.text for s in sorted(remark_spans + disqualified, key=lambda s: (s.y, s.x0))
            ])
        )
    if disqualified:
        row["Status"] = "DSQ"

    lines = _lines(data_spans)
    if table.layout == "competition":
        if len(lines) > MAX_ROUNDS:
            raise _fail(state, page, f"{label}: more than {MAX_ROUNDS} rounds.")
        line_columns = [table.line_columns[0]] * len(lines)
        round_numbers = list(range(1, len(lines) + 1))
    else:
        if len(lines) > 2:
            raise _fail(state, page, f"{label}: unexpected training row {[s.text for s in data_spans]}.")
        line_columns = table.line_columns[: len(lines)]
        if (len(lines) == 1 and len(table.line_columns) == 2
                and all(RANK_PATTERN.match(s.text) for s in lines[0])):
            line_columns = table.line_columns[1:]  # no jump, a rank line only (Notodden 2018)
        round_numbers = [None] * len(lines)

    for line, columns, round_number in zip(lines, line_columns, round_numbers):
        cells: dict[str, list[Span]] = {}
        for span in line:
            column = next((c for c in columns if c.low <= span.x < c.high), None)
            if column is None:
                raise _fail(state, page, f"{label}: value '{span.text}' outside all columns.")
            cells.setdefault(column.field, []).append(span)
        # GP Women 2023: the gate (integer) printed left of its column, in the
        # judge points column; judge points always have a decimal.
        for column_field in list(cells):
            if column_field.startswith("Judge Points#") and len(cells[column_field]) == 2:
                gate_field = "Gate#" + column_field.split("#")[1]
                gates = [s for s in cells[column_field] if re.fullmatch(r"©?\d{1,2}", s.text)]
                if len(gates) == 1 and gate_field not in cells and any(c.field == gate_field for c in columns):
                    cells[column_field].remove(gates[0])
                    cells[gate_field] = gates
        for column_field, cell_spans in cells.items():
            field_name, group = column_field.split("#")
            number = round_number if round_number is not None else int(group)
            cell_spans.sort(key=lambda s: s.x0)
            if len(cell_spans) > 1 and field_name != "Gate":
                raise _fail(
                    state, page,
                    f"{label}: several values in {field_name}: {[s.text for s in cell_spans]}.",
                )
            if field_name == SKIPPED:
                continue
            text = "".join(s.text for s in cell_spans)
            if field_name == "Round Total" and RANK_PATTERN.match(text) and isinstance(row.get("Bib"), str):
                field_name = "Round Rank"  # team member out after round 1: rank left of its column
            if text in STATUS_CODES:
                row["Status"] = text
                continue
            parsed = _parse_value(field_name, text)
            if parsed is None:
                raise _fail(state, page, f"{label}: unreadable {field_name} '{text}'.")
            for key, value in parsed.items():
                if key == "Remarks":
                    row["Remarks"] = " ".join(filter(None, [row.get("Remarks"), value]))
                elif value is not None:
                    row[f"{key} {number}"] = value
    if (table.layout == "competition" and isinstance(row.get("Bib"), str) and len(lines) == 1
            and row.get("Total") is not None and row.get("Round Total 1") is None
            and row.get("Distance 1") is not None):
        # Team member with one round: its points printed in the total column only.
        row["Round Total 1"] = row["Total"]
    return row


def _section_rows(
    spans: list[Span], start_y: float, end_y: float, status: str, state: _State, page: int
) -> list[dict[str, Any]]:
    rows = []
    explanations: dict[str, str] = {}
    for line in _lines([s for s in spans if start_y < s.y < end_y]):
        line.sort(key=lambda s: s.x0)
        texts = [s.text for s in line]
        match = re.match(r"^(\d+)(?:\s+(.+))?$", texts[0])
        if match is None and re.match(r"^(?:SCE|ICR)\b", texts[0]):
            # Swiss Timing explains the rule codes below the list.
            explanation = " ".join(texts)
            code = explanation.split(" - ")[0].strip()
            explanations[code] = explanation
            continue
        if match is None:
            break  # end of the list (legend, statistics, ...): no bib
        rest = ([match.group(2)] if match.group(2) else []) + texts[1:]
        if (len(rest) > 1 and re.fullmatch(r"\d+", rest[0]) and NAME_TEXT_PATTERN.match(rest[1])
                and not NSA_PATTERN.match(rest[1])):
            rest = [f"{rest[0]} {rest[1]}", *rest[2:]]  # rank, bib, name in three texts
        bib_name = re.match(r"^(\d+)\s+(\D.+)$", rest[0]) if rest else None
        if bib_name:
            # Olympics 2022: rank, then bib and name in one text ('30', '48 GRANERUD HE').
            match = bib_name
            rest = [bib_name.group(2), *rest[1:]]
        # Nation and rule in one span: 'NOR SCE4 - Suit' (Olympics 2022: 'NOR Final')
        rest = [
            part for index, text in enumerate(rest)
            for part in (text.split(" ", 1) if re.match(r"^[A-Z]{3} (?:SCE|ICR)", text)
                         or (index and re.match(r"^[A-Z]{3} [A-Z][a-z]+$", text)) else [text])
        ]
        nsa_index = next((i for i, text in enumerate(rest) if NSA_PATTERN.match(text)), None)
        if nsa_index is None and rest and TEAM_NAME_PATTERN.match(rest[0]):
            # Team event: 'USA - UNITED STATES OF AMERICA'
            rows.append({"Bib": int(match.group(1)), "Name": rest[0], "NSA": rest[0][:3],
                         "Status": status, **({"Remarks": " ".join(rest[1:])} if rest[1:] else {})})
            continue
        if nsa_index is None and rest and looks_like_athlete(rest[0]):
            # Local lists: bib, name, club without nation.
            row = {"Bib": int(match.group(1)), "Name": rest[0], "Status": status}
            if len(rest) > 1:
                row["Club"] = " ".join(rest[1:])
            rows.append(row)
            continue
        if nsa_index is None or nsa_index == 0:
            raise _fail(state, page, f"unreadable {status} row {texts}.")
        row: dict[str, Any] = {
            "Bib": int(match.group(1)),
            "Name": " ".join(rest[:nsa_index]),
            "NSA": rest[nsa_index],
            "Status": status,
        }
        remarks = " ".join(rest[nsa_index + 1 :])
        if remarks:
            row["Remarks"] = remarks
        rows.append(row)
    for row in rows:
        remarks = row.get("Remarks", "")
        for code, explanation in explanations.items():
            if re.search(rf"{re.escape(code)}(?![\w.])", remarks) and explanation not in remarks:
                row["Remarks"] = re.sub(
                    rf"{re.escape(code)}(?![\w.])", explanation, remarks, count=1
                )
                break
    return rows


FASTREPORT_LABELS = {  # Chaikovsky (Russian championships until 2022)
    "km/h": "[km/h]", "m": "[m]", "m/s": "[m/s]", "Points": "POINTS", "Gate": "GATE", "Rank": "RANK",
    # 2020: labels cut off by the column width
    "POIN": "POINTS", "GAT": "GATE", "TOTA": "POINTS", "RAN": "RANK",
}


def _fastreport_spans(spans: list[Span]) -> list[Span]:
    """FastReport lists: decimal comma ('79,8'), labels without brackets."""
    header_y = next((s.y for s in spans if s.text == "km/h"), None)
    result = []
    for span in spans:
        text = span.text
        if re.fullmatch(r"-?\d+,\d*", text):
            text = text.replace(",", ".") + ("0" if text.endswith(",") else "")  # 2020: '18,' cut off
        elif header_y is not None and abs(span.y - header_y) < 3 and text in FASTREPORT_LABELS:
            text = FASTREPORT_LABELS[text]
        elif header_y is not None and text in ("Total", "TOTAL") and abs(span.y - header_y) < 9:
            text = "TOTAL"
        elif re.fullmatch(r"(?:19|20)\d\d", text):
            continue  # year of birth only
        bold = span.bold
        if bold and " " not in text and "-" in text and span.x0 < 200:
            bold = False  # club below the name, printed bold too ('Saxalinskaya-oblast')
        result.append(Span(text, span.x0, span.x1, span.y0, span.y1, bold))
    return result


def _parse_page(page: pymupdf.Page, page_number: int, state: _State) -> None:
    spans = _legacy_labels(_split_labels(_page_spans(page)))
    if state.fastreport:
        spans = _fastreport_spans(spans)
    tables = _find_tables(spans, page.rect.width, state, page_number)
    # Pages without any bold print (embedded fonts): headings are plain text.
    plain = not any(s.bold for s in spans)
    headings = sorted(
        (
            (s.y, status)
            for s in spans
            if s.bold or plain
            for pattern, status in SECTION_HEADINGS
            if pattern.match(s.text)
        ),
    )
    end_marks = sorted(
        [s.y for s in spans if (s.bold or plain) and END_HEADINGS.match(s.text)]
        # signatures below the table, not bold (Zakopane 2023)
        + [s.y for s in spans if s.text in ("FIS Technical Delegate", "Chief of Competition")]
        + [s.y for s in spans if FOOTER_PATTERN.match(s.text)]
        + [y for y, _ in headings]
    )

    for table in tables:
        state.layouts.add(table.layout)
        table_end = min((y for y in end_marks if y > table.header_y), default=page.rect.height)
        if table.bottom is not None:
            table_end = min(table_end, table.bottom)
        table_spans = [
            s for s in spans
            if table.left - 10 <= s.x0 < table.right and table.header_y + 3 < s.y < table_end
        ]
        # Swiss Timing starts this separator row left of the table edge.
        not_qualified_y = min(
            (
                s.y for s in spans
                if NOT_QUALIFIED_PATTERN.match(s.text)
                and table.left - 16 <= s.x0 < table.right
                and table.header_y + 3 < s.y < table_end
            ),
            default=None,
        )
        prequalified_y = min(
            (s.y for s in table_spans if PREQUALIFIED_PATTERN.match(s.text)), default=None
        )
        used: set[int] = set()
        team = any(_is_team_name(s.text) for s in table_spans) or (
            # Predazzo 2022: a page that continues the jumpers of the last team
            state.team is not None
            and any(s.x < table.data_low and TEAM_BIB_PATTERN.match(s.text) for s in table_spans)
        )
        cells = _row_cells(page, table, table_end, team)
        name_cells = _name_cells(table_spans, table, table_end)
        covered = sum(
            any(top <= (name_top + name_bottom) / 2 < bottom for top, bottom in cells)
            for name_top, name_bottom in name_cells
        )
        if not cells or (not team and len(name_cells) > 1 and (
            len(cells) > 1.5 * len(name_cells)
            or covered * 2 < len(name_cells)  # a stray line only (YOG 2016)
        )):
            # No row lines, or lines of a neighbouring table (half rows).
            cells = name_cells
        elif team:
            cells = _team_cells(cells, table_spans, table, super_team=bool(
                re.search(r"Super\s*Team", page.get_text()[:3000], re.I)
                or any(ROUND_NOTE_PATTERN.match(s.text) for s in table_spans)  # '1st Rnd', 'Final Rnd'
            ))
        if cells and team:
            # Super Team: the round 1 line of the first jumper of a page stands
            # just above its name, above the first row line.
            above = [s for s in table_spans if s.y < cells[0][0] and s.x >= table.data_low]
            names = [s.y for s in table_spans if s.x < table.data_low and cells[0][0] <= s.y < cells[0][1]
                     and TEAM_BIB_PATTERN.match(s.text)]
            if above and names and min(names) - max(s.y for s in above) < 12:
                cells[0] = (min(s.y for s in above) - 1, cells[0][1])
        # Classes ranked apart (Rasnov 2023: 'U15 GIRLS', 'U15 BOYS').
        class_headings = [
            s for s in table_spans if CLASS_HEADING_PATTERN.match(s.text)
            # mixed team trial round (Zakopane 2023): 'Women' and 'Men' tables side by side
            or (not team and s.text in ("Women", "Men") and not any(t.x >= table.data_low and abs(t.y - s.y) < 2
                                                                      for t in table_spans))
        ]
        used.update(id(s) for s in class_headings)
        for top, bottom in cells:
            block = [s for s in table_spans if top <= s.y < bottom and s not in class_headings]
            used.update(id(s) for s in block)
            row = _parse_block(block, table, state, page_number)
            if row is None:
                continue
            heading = max((s for s in class_headings if s.y < bottom - 2), key=lambda s: s.y, default=None)
            if heading is not None:
                state.class_group = heading.text
            if state.class_group is not None:
                state.rank_groups.setdefault(state.class_group, []).append(row)
            if state.prequalified or (prequalified_y is not None and top > prequalified_y):
                # Below the not qualified athletes: in the next round without qualifying.
                row.setdefault("Remarks", PREQUALIFIED)
            elif state.not_qualified or (not_qualified_y is not None and top > not_qualified_y):
                row.setdefault("Status", "Not Qualified")
            if row.get(TEAM_KEY):
                state.team = row
            elif isinstance(row.get("Bib"), str) and not row.get("NSA"):
                # Jumper of the team above: nation of the team.
                if state.team is not None and state.team.get("Bib") is None and row["Bib"].endswith("-1"):
                    state.team["Bib"] = int(row["Bib"].split("-")[0])  # 2023: team bib not printed
                if state.team is None or not row["Bib"].startswith(f"{state.team.get('Bib')}-"):
                    raise _fail(state, page_number, f"{row['Name']}: team of bib {row['Bib']} not found.")
                if state.team.get("NSA"):  # mixed team: nation of the jumper unknown
                    row.setdefault("NSA", state.team["NSA"])
            state.rows.append(row)
        if not_qualified_y is not None:
            state.not_qualified = True
        if prequalified_y is not None:
            state.prequalified = True
        lost = [
            s.text for s in table_spans
            if id(s) not in used and not NOT_QUALIFIED_PATTERN.match(s.text)
            and not NO_ROW_PATTERN.match(s.text) and not FOOTER_PATTERN.match(s.text)
        ]
        if lost:
            raise _fail(state, page_number, f"values outside athlete rows: {lost[:8]}.")

    for y, status in headings:
        end = min((mark for mark in end_marks if mark > y + 1), default=page.rect.height)
        state.section_rows.extend(_section_rows(spans, y + 1, end, status, state, page_number))


def _hill_values(text: str) -> tuple[float | None, float | None]:
    normalized = " ".join(text.split())
    k_point = re.search(r"\bK-POINT\s*:?\s*(\d+(?:\.\d+)?)\s*m?\b", normalized, re.I)
    meter = re.search(r"\bMET(?:ER|RE) (?:VALUE|WERT)\s*:?\s*(\d+(?:\.\d+)?)\b", normalized, re.I)
    return (
        float(k_point.group(1)) if k_point else None,
        float(meter.group(1)) if meter else None,
    )


def _fit_hill(rows: list[dict[str, Any]]) -> tuple[float, float] | None:
    """(K-point, meter value) from distances and distance points when the PDF
    does not print them (Ruhpolding: one hill per group, '40.0 m Schanze').
    All jumps of the rows must fit 60 + (distance - K) * meter value."""
    pairs = [
        (row[f"Distance {number}"], row[f"Distance Points {number}"])
        for row in rows for number in range(1, MAX_ROUNDS + 1)
        if row.get(f"Distance {number}") is not None and row.get(f"Distance Points {number}") is not None
    ]
    if len({distance for distance, _ in pairs}) < 2:
        return None
    (low, low_points), (high, high_points) = min(pairs), max(pairs)
    meter_value = round((high_points - low_points) / (high - low), 1)
    if meter_value <= 0:
        return None
    k_point = round((high - (high_points - 60.0) / meter_value) * 2) / 2
    if all(abs(60.0 + (distance - k_point) * meter_value - points) <= TOLERANCE for distance, points in pairs):
        return k_point, meter_value
    return None


def _check_ranks(
    rows: list[dict[str, Any]], value_key: str, rank_key: str, errors: list[str], dense: bool = False
) -> None:
    """Ranks follow the values: higher value, better rank; same value, same rank.

    Ranks may have gaps: an athlete ranked after round 1 and disqualified
    later is no longer listed with the round 1 result.
    """
    ranked = [
        row for row in rows if row.get(value_key) is not None and row.get(rank_key) is not None
        # Training: jumps on 0.0 points (never below) share the rank of the last with points.
        and not ("Total" in value_key and row[value_key] == 0.0)
    ]
    for row in ranked:
        better = [other[value_key] for other in ranked if other[value_key] > row[value_key] + 1e-9]
        lowest = 1 + (len(set(better)) if dense else len(better))
        consistent = all(
            other[rank_key] == row[rank_key]
            if abs(other[value_key] - row[value_key]) <= 1e-9
            else (other[rank_key] < row[rank_key]) == (other[value_key] > row[value_key])
            for other in ranked
        )
        if row[rank_key] < lowest or not consistent:
            errors.append(f"{row.get('Name')}: {rank_key} {row[rank_key]} expected {lowest}")


def _validate(state: _State, text: str) -> None:
    errors: list[str] = []
    rows = state.rows
    if not rows:
        errors.append("no athletes found")
    k_point, meter_value = _hill_values(text)
    hill_of: dict[int, tuple[float, float]] = {}  # row -> (K-point, meter value) found from its group
    if (k_point is None or meter_value is None) and rows:
        for group in list(state.rank_groups.values()) or [rows]:
            fitted = _fit_hill(group)
            if fitted is None:
                break
            hill_of.update((id(row), fitted) for row in group)
    if (k_point is None or meter_value is None) and not (rows and len(hill_of) == len(rows)):
        errors.append("K-point or meter value not found")
    if state.wind_column and not any(
        row.get(f"Wind Points {number}") is not None or row.get(f"Wind {number}") is not None
        for row in rows for number in range(1, MAX_ROUNDS + 1)
    ):
        state.wind_column = False  # no wind measured (Notodden 2025): columns empty
    if "competition" in state.layouts:
        distance_errors: list[tuple[dict[str, Any], int, float]] = []
        factors = re.search(r"Wind\s*Fa[ck]tor[^\d\n]*\n?\s*:?\s*(\d+[.,]\d+)\s*/\s*(\d+[.,]\d+)", text, re.I)
        gate_match = re.search(r"Gate\s*Fa[ck]tor[^\d\n]*\n?\s*:?\s*(\d+[.,]\d+)", text, re.I)
        gate_factor = float(gate_match.group(1).replace(",", ".")) if gate_match else None
        wind_factors = (
            (float(factors.group(1).replace(",", ".")), float(factors.group(2).replace(",", ".")))
            if factors and not any(row.get(f"Wind Points {n}") is not None
                                   for row in rows for n in range(1, MAX_ROUNDS + 1))
            else None
        )
        for row in rows:
            row_k, row_meter = hill_of.get(id(row), (k_point, meter_value))
            base = 120.0 if (row_k or 0) >= 170 else 60.0
            for number in range(1, MAX_ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                distance_points = row.get(f"Distance Points {number}")
                if distance is None or distance_points is None or row_k is None or row_meter is None:
                    continue
                calculated = base + (distance - row_k) * row_meter
                if abs(calculated - distance_points) > TOLERANCE:
                    distance_errors.append((row, number, abs(calculated - distance_points) - row_meter / 2))
        # Garmisch 2016: one distance printed half a metre off its points (misprint of the PDF).
        if len(distance_errors) == 1 and abs(distance_errors[0][2]) <= TOLERANCE:
            row, number, _difference = distance_errors[0]
            row["Remarks"] = " ".join(filter(None, [
                row.get("Remarks"), f"distance points of round {number} do not fit the distance (PDF)"]))
        else:
            errors.extend(f"{row['Name']}: distance points round {number}" for row, number, _d in distance_errors)
        # Grand Prix 2025: the total is the points of the final round only.
        finals = [
            (row["Total"], row[f"Round Total {last}"]) for row in rows
            if row.get("Total") is not None and not row.get(TEAM_KEY)
            and (last := max((n for n in range(1, MAX_ROUNDS + 1) if row.get(f"Round Total {n}") is not None),
                             default=0)) >= 2
        ]
        final_only = bool(finals) and all(abs(total - final) <= TOLERANCE for total, final in finals)
        # Zao 2023: rounds without any gate points printed although the gates
        # differ; the difference must then be the same for all on one gate.
        unprinted_gate_points = {
            number for number in range(1, MAX_ROUNDS + 1)
            if not any(row.get(f"Gate Points {number}") is not None for row in rows)
            and len({row.get(f"Gate {number}") for row in rows if row.get(f"Gate {number}") is not None}) != 1
            # (Klingenthal 2019: no gate printed at all, the same compensation for all)
            and any(row.get(f"Round Total {number}") is not None for row in rows)
        }
        gate_differences: dict[tuple[int, int | None, str | None], list[tuple[str, float]]] = {}
        group_of = {id(row): name for name, group in state.rank_groups.items() for row in group}
        unmeasured: list[str] = []
        for row in rows:
            if row.get("Remarks") == PREQUALIFIED:
                continue  # jumps without judges and total
            round_sum = 0.0
            rounds = 0
            for number in range(1, MAX_ROUNDS + 1):
                if row.get(f"Distance {number}") is None:
                    if row.get(f"Round Total {number}") is not None and row.get(f"Distance Points {number}") is not None:
                        # Deduction without a jump (Oberstdorf 2025: -70.0 below round 1).
                        round_sum += row[f"Round Total {number}"]
                        rounds += 1
                    continue
                rounds += 1
                marks = [row.get(f"{judge} {number}") for judge in "ABCDE"]
                judge_points = row.get(f"Judge Points {number}")
                if all(mark is not None for mark in marks) and judge_points is not None:
                    if abs(sum(sorted(marks)[1:4]) - judge_points) > TOLERANCE:
                        errors.append(f"{row['Name']}: judge points round {number}")
                parts = [
                    row.get(f"Distance Points {number}"),
                    judge_points,
                    row.get(f"Gate Points {number}") or 0.0,
                    row.get(f"Wind Points {number}") if state.wind_column else 0.0,
                ]
                wind = row.get(f"Wind {number}")
                if parts[3] is None and wind is not None and wind_factors is not None:
                    # Klingenthal 2019: wind in m/s only, its points from the wind factors.
                    parts[3] = round(-wind * (wind_factors[0] if wind > 0 else wind_factors[1]), 1)
                if state.blank_is_zero:
                    parts = [0.0 if part is None and index in (0, 3) else part
                             for index, part in enumerate(parts)]
                if (parts[3] is None and row.get(f"Wind {number}") is None and None not in parts[:2]
                        and row.get(f"Round Total {number}") is not None):
                    # Lahti 2023: no wind measured for a few jumpers, their round
                    # points include a compensation that is not printed.
                    unmeasured.append(f"{row['Name']}: incomplete round {number}")
                    round_sum += row[f"Round Total {number}"]
                    continue
                if any(part is None for part in parts):
                    errors.append(f"{row['Name']}: incomplete round {number}")
                    continue
                round_total = row.get(f"Round Total {number}")
                calculated = max(0.0, sum(parts))  # a round never counts below 0 points
                if round_total is not None and abs(calculated - round_total) > TOLERANCE and not (
                    state.gate_points_missing and row.get(f"Gate {number}") is not None
                ):
                    if number in unprinted_gate_points:
                        gate_differences.setdefault((number, row.get(f"Gate {number}"), group_of.get(id(row))), []).append(
                            (f"{row['Name']}: round {number} total {calculated:.1f} != {round_total}",
                             round_total - calculated))
                        round_sum = round_total if final_only else round_sum + round_total
                        continue
                    errors.append(f"{row['Name']}: round {number} total {calculated:.1f} != {round_total}")
                round_points = round_total if round_total is not None else calculated
                round_sum = round_points if final_only else round_sum + round_points
            if final_only:
                rounds = 1
            if row.get(TEAM_KEY) or isinstance(row.get("Bib"), str):
                continue  # team events: checked below
            # Planica 2017: negative distance points count as 0 in the total.
            floored = sum(
                max(0.0, row[f"Distance Points {n}"]) + (row.get(f"Judge Points {n}") or 0.0)
                for n in range(1, MAX_ROUNDS + 1) if row.get(f"Distance Points {n}") is not None
            ) if any((row.get(f"Distance Points {n}") or 0.0) < 0 for n in range(1, MAX_ROUNDS + 1)) else None
            if floored is not None and row.get("Total") is not None and abs(floored - row["Total"]) <= TOLERANCE * max(rounds, 1):
                pass
            elif row.get("Total") is not None and abs(round_sum - row["Total"]) > TOLERANCE * max(rounds, 1):
                errors.append(f"{row['Name']}: total {round_sum:.1f} != {row['Total']}")
            if row.get("Total") is None and row.get("Status") not in STATUS_CODES:
                errors.append(f"{row['Name']}: missing total")
        jumps = sum(1 for row in rows for number in range(1, MAX_ROUNDS + 1) if row.get(f"Distance {number}") is not None)
        if len(unmeasured) * 4 > jumps:
            errors.extend(unmeasured)  # wind missing for most: not read correctly
        for (number, _gate, group), differences in gate_differences.items():
            # (per class: the juniors may start from another gate, Klingenthal 2019)
            same_gate = [row for row in rows if row.get(f"Gate {number}") == _gate
                         and group_of.get(id(row)) == group
                         and row.get(f"Round Total {number}") is not None]
            if _gate is None:
                # No gate printed (Klingenthal 2019): the gate may have changed
                # during the round; each difference must be a gate compensation
                # of whole half metres.
                step = gate_factor / 2 if gate_factor else None
                if step is None or not all(
                    abs(d - round(d / step) * step) <= TOLERANCE and 1 <= abs(round(d / step)) <= 8
                    for _m, d in differences
                ):
                    errors.extend(message for message, _d in differences)
            elif len(differences) != len(same_gate) or max(d for _m, d in differences) - min(
                    d for _m, d in differences) > TOLERANCE:
                errors.extend(message for message, _d in differences)
        teams = [row for row in rows if row.get(TEAM_KEY)]
        if teams:
            # Team total = points of all its jumpers in all rounds.
            for team in teams:
                members = [
                    row for row in rows
                    if isinstance(row.get("Bib"), str) and row["Bib"].startswith(f"{team.get('Bib')}-")
                ]
                points = [
                    row.get(f"Round Total {number}") or 0.0
                    for row in members for number in range(1, MAX_ROUNDS + 1)
                ]
                if not members:
                    errors.append(f"{team['Name']}: no jumpers")
                elif team.get("Total") is not None and abs(sum(points) - team["Total"]) > TOLERANCE * len(points):
                    errors.append(f"{team['Name']}: total {sum(points):.1f} != {team['Total']}")
            for group in list(state.rank_groups.values()) or [rows]:
                _check_ranks([row for row in group if row.get(TEAM_KEY)], "Total", "Rank", errors)
        else:
            # KO: winners and lucky losers ranked apart from the other losers.
            groups = (
                [
                    [row for row in rows if row.get("Status") != "Not Qualified"],
                    [row for row in rows if row.get("Status") == "Not Qualified"],
                ]
                # KO (Four Hills), Grand Prix 2025: not qualified for the final round
                if KO_PATTERN.search(text[:2000]) or any(row.get("Status") == "Not Qualified" for row in rows)
                else list(state.rank_groups.values()) or [rows]
            )
            for group in groups:
                _check_ranks(group, "Total", "Rank", errors)
                for number in range(1, MAX_ROUNDS + 1):
                    round_errors: list[str] = []
                    _check_ranks(group, f"Round Total {number}", f"Round Rank {number}", round_errors)
                    if round_errors:
                        # Swiss Timing until 2023: who did not start or was disqualified in
                        # the next round is left out of the ranks of this round (its own
                        # rank is meaningless); the others must be ranked without them.
                        completed = [row for row in group if row.get(f"Round Total {number + 1}") is not None]
                        if len(completed) < len(group) and any(
                            row.get(f"Round Total {number}") is not None for row in completed
                        ):
                            retry: list[str] = []
                            _check_ranks(completed, f"Round Total {number}", f"Round Rank {number}", retry)
                            # (Erzurum 2017: ranked after the jumpers of the next round)
                            dropped = [
                                row for row in group if row not in completed
                                and row.get(f"Round Rank {number}") is not None
                            ]
                            if not retry and dropped:
                                round_errors = []
                    errors.extend(round_errors)
    if "training" in state.layouts:
        point_errors: list[str] = []
        checked = 0
        for row in rows:
            for number in range(1, MAX_ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                points = row.get(f"Round Total {number}")
                if distance is None or points is None or k_point is None or meter_value is None:
                    continue
                checked += 1
                calculated = max(
                    0.0,
                    (120.0 if k_point >= 170 else 60.0)
                    + (distance - k_point) * meter_value
                    + (row.get(f"Gate Points {number}") or 0.0)
                    + (row.get(f"Wind Points {number}") or 0.0),
                )
                if abs(calculated - points) > TOLERANCE:
                    point_errors.append(f"{row['Name']}: round {number} points {calculated:.1f} != {points}")
        # Hinterzarten 2024: training points on another base and with gate
        # points not printed: when most jumps differ, only the ranks are checked.
        # A single jump off (Liberec 2023: 79.5 m printed for 55.4 points) is a
        # misprint of the PDF, not of the reading.
        if 1 < len(point_errors) <= checked / 2:
            errors.extend(point_errors)
        for group in list(state.rank_groups.values()) or [rows]:
            for number in range(1, MAX_ROUNDS + 1):
                for value_key, rank_key in (("Speed", "Speed Rank"), ("Distance", "Distance Rank"),
                                            ("Round Total", "Round Rank")):
                    _check_ranks(group, f"{value_key} {number}", f"{rank_key} {number}", errors,
                                 dense=state.dense_ranks)
    # Groups ranked apart may start their bibs again at 1 (ladies, then men).
    duplicates = []
    # (NPS: bib of the qualification, Kulm 2024)
    for group in list(state.rank_groups.values()) or [
        rows + [row for row in state.section_rows if row.get("Status") != "NPS"
                # Olympics 2022: disqualified in the final, listed in the table too
                and not any(other.get("Bib") == row.get("Bib") and other.get("Status") == row.get("Status")
                            for other in rows)]
    ]:
        bibs = [row["Bib"] for row in group if row.get("Bib") is not None]
        duplicates.extend(bib for bib in bibs if bibs.count(bib) > 1)
    duplicates = sorted(set(duplicates), key=str)
    if duplicates:
        errors.append(f"duplicate bibs {duplicates}")
    if errors:
        raise ResultsParseError(
            f"{state.pdf_name}: result check failed: " + "; ".join(errors[:5])
            + (f" (+{len(errors) - 5} more)" if len(errors) > 5 else "")
        )


# --- WinSpringen (German timing software, DSV events printed with
# 'Microsoft Print to PDF'): German column labels, competition with two
# lines per athlete (line 1: rank, bib, name, round 1, total; line 2: club,
# nation, round 2), training with one line per athlete.
WINSPRINGEN_PATTERN = re.compile(r"\bWinSpringen\b")
WINSPRINGEN_LABELS = {
    "RANG": "Rank",
    "STNR": "Bib",
    "FIS-CODE": "FIS Code",
    "NAME": "Name",
    "JG": "Year",
    "VBD": "Association",
    "VEREIN": "Club",
    "STNR-P": "Bib",  # team training: '1 - 1'
    "TEAM": "Club",
    "KM/H": "Speed",
    "WEITE": "Distance",
    "W-N.": "Distance Points",
    "-A-": "A",
    "-B-": "B",
    "-C-": "C",
    "-D-": "D",
    "-E-": "E",
    "H-N.": "Judge Points",
    "WIND": "Wind",
    "LUKE": "Gate",
    "G-N.": "Round Total",
    "T-N.": "Total",
    # Ruhpolding until 2019: labels written out.
    "W-NOTE": "Distance Points",
    "H-NOTE": "Judge Points",
    "G-NOTE": "Round Total",
    "T-NOTE": "Total",
    "PP": "Cup Points",  # cup points of the place: not read
}
WINSPRINGEN_FOLLOWERS = {  # 'pts' / 'Rg' take the meaning of the column before
    ("PTS", "Wind"): "Wind Points",
    ("PTS", "Gate"): "Gate Points",
    ("PTS", "Judge Points"): "Wind Points",  # Klingenthal 2019: wind points without m/s
    ("RG", "Speed"): "Speed Rank",
    ("RG", "Distance"): "Distance Rank",
}
WINSPRINGEN_IDENTITY = ("Rank", "Bib", "FIS Code", "Name", "Year", "Association", "Club")
WINSPRINGEN_END = re.compile(
    r"^(?:From the referee|Aus dem Kampfrichterprotokoll|Statisti[ck]s?\b|Jury\b|WinSpringen\b|Erstellt am\b|Seite \d|"
    r"For the correctness|Weather\b|Wetter\b)",
    re.I,
)
WINSPRINGEN_SECTIONS = (
    (re.compile(r"^(?:Did not start|Nicht gestartet|Nicht am Start)\b", re.I), "DNS"),
    (re.compile(r"^(?:Disqualified|Disqualifiziert)\b", re.I), "DSQ"),
    (re.compile(r"^(?:Did not finish|Nicht beendet)\b", re.I), "DNF"),
)
WINSPRINGEN_GROUPS = re.compile(
    r"^(?:Men|Women|Ladies|Mixed|Youth|Herren|Damen|Schüler(?:innen)?\b.*)$", re.I  # 'Schüler S12 m'
)


@dataclass
class _WinColumn:
    field: str
    x0: float
    x1: float

    @property
    def x(self) -> float:
        return (self.x0 + self.x1) / 2


def _winspringen_header(spans: list[Span], state: _State, page: int) -> list[_WinColumn] | None:
    bib_labels = [s for s in spans if s.text.upper() in ("STNR", "STNR-P")]  # StNr-P: team training
    if not bib_labels:
        return None
    header_y = bib_labels[0].y
    columns: list[_WinColumn] = []
    for span in sorted((s for s in spans if abs(s.y - header_y) < 3), key=lambda s: s.x0):
        label = span.text.upper()
        if label in WINSPRINGEN_LABELS:
            field_name = WINSPRINGEN_LABELS[label]
        else:
            previous = columns[-1].field if columns else ""
            field_name = WINSPRINGEN_FOLLOWERS.get((label, previous))
            if field_name is None:
                raise _fail(state, page, f"unknown WinSpringen column '{span.text}'.")
        columns.append(_WinColumn(field_name, span.x0, span.x1))
    if not any(c.field == "Name" for c in columns) or not any(c.field == "Speed" for c in columns):
        raise _fail(state, page, f"unknown WinSpringen header {[c.field for c in columns]}.")
    return columns


def _winspringen_rounds(columns: list[_WinColumn]) -> list[tuple[_WinColumn, int | None]]:
    """Data columns with their round: training prints the rounds side by
    side (a new round at each speed column), competition one line per round."""
    data = [c for c in columns if c.field not in WINSPRINGEN_IDENTITY]
    training = not any(c.field == "Total" or c.field == "Round Total" for c in data)
    result = []
    number = 0
    for column in data:
        if training and column.field == "Speed":
            number += 1
        result.append((column, number if training else None))
    return result


def _winspringen_value(row: dict[str, Any], field_name: str, number: int, text: str,
                       state: _State, page: int) -> None:
    if field_name == "Cup Points":
        return
    if text.upper() in STATUS_CODES:
        row["Status"] = text.upper()
        return
    if field_name == "Total":
        if not NUMBER_PATTERN.match(text):
            raise _fail(state, page, f"{row.get('Name')}: unreadable total '{text}'.")
        row["Total"] = float(text)
        return
    parsed = _parse_value(field_name, text)
    if parsed is None:
        raise _fail(state, page, f"{row.get('Name')}: unreadable {field_name} '{text}'.")
    for key, value in parsed.items():
        if value is not None:
            row[f"{key} {number}"] = value


@dataclass
class _WinPage:
    """Column layout of a WinSpringen page (pages without header: the one before)."""

    number: int
    columns: list[_WinColumn]
    rounds: list[tuple[_WinColumn, int | None]]
    starts: list[tuple[_WinColumn, float]]
    data_low: float
    training: bool
    start_field: str

    def identity_field(self, span: Span) -> str | None:
        if span.x >= self.data_low:
            return None
        return [column.field for column, start in self.starts if span.x0 >= start - 2][-1]


def _winspringen_layout(columns: list[_WinColumn], page_number: int) -> _WinPage:
    rounds = _winspringen_rounds(columns)
    training = any(number is not None for _, number in rounds)
    identity = [c for c in columns if c.field in WINSPRINGEN_IDENTITY]
    # Identity columns by their left edge (text left aligned, numbers right
    # aligned below the label): a column starts halfway after the label before.
    starts = [
        (c, -1.0 if index == 0 else (identity[index - 1].x1 + c.x0) / 2)
        for index, c in enumerate(identity)
    ]
    start_field = "Bib" if training or not any(c.field == "Rank" for c in identity) else "Rank"
    return _WinPage(page_number, columns, rounds, starts, min(c.x0 for c, _ in rounds) - 8,
                    training, start_field)


def _winspringen_document(document: pymupdf.Document, state: _State) -> None:
    """All pages of a WinSpringen list: an athlete may continue on the next
    page (second line), later pages may have no column header."""
    layout: _WinPage | None = None
    lines: list[tuple[_WinPage, list[Span]]] = []
    for page_number, page in enumerate(document, start=1):
        spans = _page_spans(page)
        columns = _winspringen_header(spans, state, page_number)
        if columns is not None:
            layout = _winspringen_layout(columns, page_number)
            top = next(s.y for s in spans if s.text.upper() in ("STNR", "STNR-P")) + 9
        elif layout is not None:
            # Continued list: from the first value above the end of the list.
            layout = _winspringen_layout(layout.columns, page_number)
            end = min((s.y for s in spans if WINSPRINGEN_END.match(s.text)), default=page.rect.height)
            values = [
                s.y for s in spans
                if s.x >= layout.data_low and NUMBER_PATTERN.match(s.text) and s.y < end
            ]
            top = min(values, default=page.rect.height) - 3
        if layout is not None:
            end = min(
                (s.y for s in spans if s.y > top and WINSPRINGEN_END.match(s.text)),
                default=page.rect.height,
            )
            for line in _lines([s for s in spans if top < s.y < end]):
                lines.append((layout, sorted(line, key=lambda s: s.x0)))
        _winspringen_sections(spans, state, page_number)
    if layout is None:
        return
    state.layouts.add("training" if layout.training else "competition")
    state.blank_is_zero = True
    state.dense_ranks = layout.training
    if not any(c.field in ("Wind", "Wind Points") for c in layout.columns):
        state.wind_column = False

    blocks: list[tuple[str, list[tuple[_WinPage, list[Span]]]]] = []
    group = ""
    for page_layout, line in lines:
        first = next((s for s in line if page_layout.identity_field(s) == page_layout.start_field), None)
        if first is not None and (BIB_PATTERN.match(first.text) or TEAM_BIB_PATTERN.match(first.text)):
            blocks.append((group, [(page_layout, line)]))
        elif any(s.x >= page_layout.data_low for s in line):
            if not blocks:
                raise _fail(state, page_layout.number, f"values outside athlete rows: {[s.text for s in line]}.")
            blocks[-1][1].append((page_layout, line))
        elif len(line) == 1 and line[0].text.upper() in WINSPRINGEN_LABELS:
            continue  # second header line of a later page ('Vbd')
        elif len(line) == 1 and (line[0].bold or re.search(r"\bSchanze\b", line[0].text)):
            group = line[0].text  # 'Boys, 64.0 m Schanze': ranked apart
        elif blocks:
            blocks[-1][1].append((page_layout, line))
    for group, block in blocks:
        row: dict[str, Any] = {}
        names: list[str] = []
        clubs: list[str] = []
        for line_index, (page_layout, line) in enumerate(block):
            page_number = page_layout.number
            if line_index > 1 and any(s.x >= page_layout.data_low for s in line):
                raise _fail(state, page_number, f"unexpected WinSpringen line {[s.text for s in line]}.")
            for span in line:
                field_name = page_layout.identity_field(span)
                if field_name is None:
                    column, number = min(page_layout.rounds, key=lambda item: abs(item[0].x - span.x))
                    if abs(column.x - span.x) > 14:
                        raise _fail(state, page_number, f"value '{span.text}' outside all columns.")
                    if number is None:
                        number = line_index + 1  # competition: one line per round
                    _winspringen_value(row, column.field, number, span.text, state, page_number)
                elif field_name == "FIS Code" and span.text.casefold() in ("ohne", "-"):
                    continue  # athlete without FIS code (Ruhpolding)
                elif field_name == "Bib" and line_index == 0 and TEAM_BIB_PATTERN.match(span.text):
                    row["Bib"] = re.sub(r"\s+", "", span.text)  # '1 - 1' -> '1-1'
                elif field_name in ("Rank", "Bib", "FIS Code") and line_index == 0:
                    match = RANK_VALUE_PATTERN.match(span.text)
                    if not match:
                        raise _fail(state, page_number, f"unreadable {field_name} '{span.text}'.")
                    if field_name != "FIS Code":
                        row[field_name] = int(match.group(1))
                elif field_name == "Name" and line_index == 0:
                    names.append(span.text)
                elif field_name in ("Year", "Association") and NSA_PATTERN.match(span.text) and line_index:
                    row["NSA"] = span.text  # line 2: nation below the year
                elif field_name in ("Year", "Association") or WINSPRINGEN_GROUPS.match(span.text):
                    continue
                else:
                    clubs.append(span.text)
        if not names:
            raise _fail(state, block[0][0].number,
                        f"row without athlete name: {[s.text for _, line in block for s in line]}.")
        row["Name"] = GENDER_SUFFIX.sub("", " ".join(names)).lstrip("*")  # Klingenthal 2019: '*NOWAK Jenny'
        if clubs:
            row["Club"] = " ".join(clubs)
        if (not layout.training and row.get("Total") is None
                and not any(c.field == "Total" for c in layout.columns)):
            row["Total"] = row.get("Round Total 1")  # after round 1: no total column
        state.rows.append(row)
        state.rank_groups.setdefault(group, []).append(row)


# --- Kilipa (EYOF Lahti 2022) -------------------------------------------------
# Results of the organiser (www.kilipa.com): nation written out below the name
# ('Slovenia'), both rounds on two lines, no wind; rounds not jumped as 0.0.
KILIPA_PATTERN = re.compile(r"\bKilipa\b|www\.kilipa\.com", re.I)
KILIPA_LABELS = {
    "M": "Distance", "DIST [M]": "Distance",
    "KM/H": "Speed", "SPEED [KM/H]": "Speed",
    "DP": "Distance Points", "DIST. POINTS": "Distance Points",
    "A": "A", "B": "B", "C": "C", "D": "D", "E": "E",
    "JP": "Judge Points", "JUDGE POINTS": "Judge Points",
    "GATE": "Gate", "GPTS": "Gate Points",
    "PTS": "Round Total", "ROUND POINTS": "Round Total",
    "RANK": "Round Rank", "TOTAL": "Total", "BEHIND": None,
}
COUNTRY_CODES = {
    "andorra": "AND", "argentina": "ARG", "armenia": "ARM", "australia": "AUS", "austria": "AUT",
    "belarus": "BLR", "belgium": "BEL", "bosnia and herzegovina": "BIH", "bulgaria": "BUL",
    "canada": "CAN", "china": "CHN", "croatia": "CRO", "czech republic": "CZE", "czechia": "CZE",
    "denmark": "DEN", "estonia": "EST", "finland": "FIN", "france": "FRA", "georgia": "GEO",
    "germany": "GER", "great britain": "GBR", "greece": "GRE", "hungary": "HUN", "iceland": "ISL",
    "israel": "ISR", "italy": "ITA", "japan": "JPN", "kazakhstan": "KAZ", "korea": "KOR",
    "kyrgyzstan": "KGZ", "latvia": "LAT", "liechtenstein": "LIE", "lithuania": "LTU",
    "montenegro": "MNE", "netherlands": "NED", "north macedonia": "MKD", "norway": "NOR",
    "poland": "POL", "romania": "ROU", "russia": "RUS", "serbia": "SRB", "slovakia": "SVK",
    "slovenia": "SLO", "spain": "ESP", "sweden": "SWE", "switzerland": "SUI", "turkey": "TUR",
    "turkiye": "TUR", "ukraine": "UKR", "united states": "USA", "united states of america": "USA", "usa": "USA",
}


def _kilipa_columns(spans: list[Span], state: _State, page: int) -> tuple[list[_WinColumn], Span, Span] | None:
    """Columns by the right edge of their (one or two line) labels; the
    first 'Rank' and the 'Name' / 'Name/Club' label give the identity part."""
    ranks = sorted((s for s in spans if s.text == "Rank" and any(
        # (the Rank label of the table: Name next to it; Whistler 2019 has a jury list above)
        other.text in ("Name/Club", "Name") and abs(other.y - s.y) < 8 for other in spans
    )), key=lambda s: (s.y, s.x0))
    if not ranks:
        return None
    rank = ranks[0]
    name = next((s for s in spans if s.text in ("Name/Club", "Name") and abs(s.y - rank.y) < 8), None)
    if name is None:
        return None
    words = {word for label in KILIPA_LABELS for word in label.split()}
    header = [
        s for s in spans if abs(s.y - rank.y) < 8 and s.x0 > name.x1 and s.text not in ("Name", "Team")
        # (Whistler 2019: the jury 'E Steve Sydow' just above the labels)
        and (abs(s.y - rank.y) < 3 or (len(s.text) > 1 and all(w in words for w in s.text.upper().split())))
    ]
    groups: list[list[Span]] = []
    for span in sorted(header, key=lambda s: s.x1):
        if groups and abs(groups[-1][-1].x1 - span.x1) < 3:
            groups[-1].append(span)
        else:
            groups.append([span])
    columns = []
    for group in groups:
        label = " ".join(s.text for s in sorted(group, key=lambda s: s.y)).upper()
        if label not in KILIPA_LABELS:
            raise _fail(state, page, f"unknown Kilipa column '{label}'.")
        if KILIPA_LABELS[label] is not None:
            columns.append(_WinColumn(KILIPA_LABELS[label], min(s.x0 for s in group), max(s.x1 for s in group)))
    return columns, rank, name


def _kilipa_document(document: pymupdf.Document, state: _State) -> None:
    columns: list[_WinColumn] | None = None
    rank_label = name_label = None
    current: dict[str, Any] | None = None
    round_number = 0
    team: dict[str, Any] | None = None
    for page_number, page in enumerate(document, start=1):
        spans = _page_spans(page)
        found = _kilipa_columns(spans, state, page_number)
        if found is not None:
            columns, rank_label, name_label = found
            top = max(s.y for s in spans if abs(s.y - rank_label.y) < 8) + 2
        elif columns is None:
            continue
        else:
            # Page without header: from the first value in the columns (below the title).
            low = min(c.x0 for c in columns) - 8
            top = min((s.y for s in spans if s.x0 >= low and NUMBER_PATTERN.match(s.text)),
                      default=page.rect.height) - 4
        end = min((s.y for s in spans if s.y > top and re.match(r"^(?:Did not start|Weather|Temperature|Printed)\b", s.text)),
                  default=page.rect.height)
        data_low = min(c.x0 for c in columns) - 8
        for line in _lines([s for s in spans if top < s.y < end]):
            line.sort(key=lambda s: s.x0)
            # Values: numbers right aligned in a column; the rest is rank, bib,
            # name or nation (a long name may reach into the values).
            values = [
                s for s in line
                if s.x0 >= data_low - 12 and re.fullmatch(r"-?\d+(?:\.\d*)?", s.text)  # '18.': round rank
                and min(abs(c.x1 - s.x1) for c in columns) <= 8
            ]
            left = [s for s in line if s not in values]
            texts = [s.text for s in line]
            fields = [min(columns, key=lambda c: abs(c.x1 - s.x1)) for s in values]
            if not left and values and all(f.field in ("Round Total", "Round Rank") for f in fields):
                continue  # team: points and rank of the team per round
            bib = next((s for s in left if re.fullmatch(r"\d+(?:-\d+)?", s.text) and s.x1 < name_label.x0 + 8), None)
            text = " ".join(s.text for s in left if s.x0 >= name_label.x0 - 6)
            joined = next((s for s in left if re.fullmatch(r"\d+ \D.*", s.text) and s.x0 < name_label.x0 + 8), None)
            rank_bib = next((s for s in left if re.fullmatch(r"\d+\.\s+\d+", s.text)), None)
            if bib is None and rank_bib is not None:
                # Whistler 2019: rank and bib in one text ('1. 9')
                rank_text, bib_text = rank_bib.text.split()
                bib = Span(bib_text, rank_bib.x1, rank_bib.x1, rank_bib.y0, rank_bib.y1, rank_bib.bold)
                left = [Span(rank_text, rank_bib.x0, rank_bib.x0 + 1, rank_bib.y0, rank_bib.y1, rank_bib.bold)
                        if s is rank_bib else s for s in left]
            if bib is None and joined is not None:
                # Whistler 2019: bib and name in one text ('33 Nathaniel Mah')
                bib = Span(joined.text.split(" ", 1)[0], joined.x0, joined.x0, joined.y0, joined.y1, joined.bold)
                text = " ".join([joined.text.split(" ", 1)[1], *(s.text for s in left if s.x0 > joined.x1)])
            if bib is not None and re.fullmatch(r"[A-Z][a-z]+(?:[ -][A-Z][a-z'’]+)+", text) and " - " not in text:
                # 'Anna Hoffman' (first name first): 'HOFFMAN Anna'
                first, last = text.rsplit(" ", 1)
                text = f"{last.upper()} {first}"
            rank = next((s for s in left if RANK_PATTERN.match(s.text) and s.x1 <= rank_label.x1 + 4), None)
            if bib is not None and "-" not in bib.text and " - " in text:
                # Team: 'SLO - SLOVENIA', total of the team.
                team = {"Bib": int(bib.text), "Name": text, "NSA": text[:3], TEAM_KEY: True}
                if rank is not None:
                    team["Rank"] = int(RANK_PATTERN.match(rank.text).group(1))
                for span, column in zip(values, fields):
                    if column.field == "Total":
                        _winspringen_value(team, "Total", 1, span.text, state, page_number)
                state.rows.append(team)
                current = None
                continue
            if bib is not None and text:
                current = {"Bib": bib.text if "-" in bib.text else int(bib.text), "Name": text}
                if rank is not None:
                    current["Rank"] = int(RANK_PATTERN.match(rank.text).group(1))
                if team is not None and "-" in bib.text:
                    current["NSA"] = team["NSA"]
                state.rows.append(current)
                round_number = 1
            elif current is not None and (values or text):
                round_number += 1
                if text:
                    nsa = COUNTRY_CODES.get(text.casefold())
                    if nsa is None:
                        # Whistler 2019: nation and club in one text ('USA Blackhawk', 'Canada Altius Nordic ...')
                        nsa = next((code for country, code in sorted(COUNTRY_CODES.items(), key=lambda item: -len(item[0]))
                                    if text.casefold().startswith(country + " ")), None)
                        if nsa is not None:
                            club = text[len(next(c for c, v in sorted(COUNTRY_CODES.items(), key=lambda item: -len(item[0]))
                                                 if text.casefold().startswith(c + " "))) + 1:]
                            current["Club"] = club
                    code = re.match(r"([A-Z]{3}) (.+)", text)
                    if nsa is None and code and code.group(1) in COUNTRY_CODES.values():
                        nsa, current["Club"] = code.groups()  # 'GBR Altius Nordic Ski Club'
                    if nsa is None:
                        raise _fail(state, page_number, f"{current['Name']}: unknown nation '{text}'.")
                    current["NSA"] = nsa
            elif not values:
                continue
            else:
                raise _fail(state, page_number, f"values outside athlete rows: {texts}.")
            jump: dict[str, Any] = {}
            for span, column in zip(values, fields):
                if abs(column.x1 - span.x1) > 8:
                    raise _fail(state, page_number, f"value '{span.text}' outside all columns.")
                if column.field == "Total":
                    if round_number == 1 and not isinstance(current["Bib"], str):
                        _winspringen_value(current, "Total", 1, span.text, state, page_number)
                    continue
                _winspringen_value(jump, column.field, round_number, span.text, state, page_number)
            if not jump.get(f"Distance {round_number}"):
                continue  # round not jumped: printed as 0.0
            marks = [jump.get(f"{judge} {round_number}") for judge in "ABCDE"]
            if all(not mark for mark in marks):
                # Training and trial round: no judges (marks 0.0).
                for judge in "ABCDE":
                    jump.pop(f"{judge} {round_number}", None)
                jump.setdefault(f"Judge Points {round_number}", 0.0)
            current.update(jump)
        _kilipa_sections(spans, end, state)
    state.layouts.add("competition")
    state.blank_is_zero = True
    state.wind_column = False


def _kilipa_sections(spans: list[Span], start: float, state: _State) -> None:
    """'Did not start:' bib, 'NAME First, Nation (SJ)'; '(2nd round)' only
    skipped the second trial round and is ranked above."""
    for line in _lines([s for s in spans if s.y > start]):
        line.sort(key=lambda s: s.x0)
        texts = [s.text for s in line]
        if not texts or not BIB_PATTERN.match(texts[0]) or len(texts) < 2:
            continue
        if any("2nd round" in text for text in texts):
            continue
        name, _, nation = texts[1].rpartition(",")
        row: dict[str, Any] = {"Bib": int(texts[0]), "Status": "DNS", "Name": name.strip() or texts[1]}
        if COUNTRY_CODES.get(nation.strip().casefold()):
            row["NSA"] = COUNTRY_CODES[nation.strip().casefold()]
        state.section_rows.append(row)


CARPATH_PATTERN = re.compile(r"Carpath Coordinator")  # Rasnov FIS Cup 2017, 2019 (Excel)


def _carpath_document(document: pymupdf.Document, state: _State) -> None:
    """Carpath Cup lists (Rasnov): rank, bib, name, country, year on one
    line, the rounds on the lines above and below; columns 'M', 'Points',
    'A'-'E', 'Points', 'Total' per round, 'Total' of the competition."""
    state.layouts.add("competition")
    state.wind_column = False
    layout = None  # columns of the page before: the next pages have no header
    for page_number, page in enumerate(document, start=1):
        spans = _page_spans(page)
        judges = next(
            (sorted(group, key=lambda s: s.x0) for letter in spans
             if letter.text in ("A", "B", "C", "D", "E") and letter.bold
             and len(group := [s for s in spans if s.bold and len(s.text) == 1 and s.text in "ABCDE"
                               and abs(s.y - letter.y) < 2]) == 5),
            None,
        )
        if judges is not None:
            header_y = judges[0].y
            header = sorted((s for s in spans if s.bold and -8 < s.y - header_y < 8), key=lambda s: s.x0)
            metre = next((s for s in header if s.text == "M"), None)
            points = [s for s in header if s.text == "Points"]
            totals = [s for s in header if s.text == "Total"]
            country = next((s for s in header if s.text in ("Country", "Nation")), None)
            bib_label = next((s for s in header if s.text == "Bib"), None)
            if metre is None or len(points) < 2 or len(totals) < 2 or country is None or bib_label is None:
                raise _fail(state, page_number, f"unknown header {[s.text for s in header]}.")
            columns = {
                "Distance": metre.x, "Distance Points": points[0].x, **{s.text: s.x for s in judges},
                "Judge Points": points[-1].x, "Round Total": totals[0].x,
            }
            layout = (columns, totals[-1], metre, country, bib_label)
            below = [s for s in spans if s.y > header_y + 7]
        elif layout is not None:
            below = spans
        else:
            continue
        columns, final_total, metre, country, bib_label = layout
        sections = sorted((s for s in below if s.text in ("DNS", "DSQ", "DNF")), key=lambda s: s.y)
        end_y = sections[0].y if sections else page.rect.height
        names = sorted(
            (s for s in below if s.y < end_y and bib_label.x1 < s.x0 < country.x0 and re.search(r"[A-Za-z]{2}", s.text)
             and not NSA_PATTERN.match(s.text)),
            key=lambda s: s.y,
        )
        for index, name in enumerate(names):
            top = (names[index - 1].y + name.y) / 2 if index else name.y - 9
            bottom = (name.y + names[index + 1].y) / 2 if index + 1 < len(names) else min(name.y + 9, end_y)
            block = [s for s in below if top <= s.y < bottom and s is not name]
            row: dict[str, Any] = {"Name": name.text}
            numbers = sorted((s for s in block if s.x1 <= name.x0 + 1 and BIB_PATTERN.match(s.text)), key=lambda s: s.x0)
            if len(numbers) == 2:
                row["Rank"], row["Bib"] = int(numbers[0].text), int(numbers[1].text)
            elif len(numbers) == 1:
                row["Bib"] = int(numbers[0].text)
            nation = next((s for s in block if NSA_PATTERN.match(s.text) and abs(s.y - name.y) < 3), None)
            if nation is not None:
                row["NSA"] = nation.text
            total = next((s for s in block if abs(s.x - final_total.x) < 12 and abs(s.y - name.y) < 3
                          and NUMBER_PATTERN.match(s.text)), None)
            if total is not None:
                row["Total"] = float(total.text)
            values = [s for s in block if s.x0 > metre.x0 - 12 and s is not total]
            for round_number, line in enumerate(_lines(values), start=1):
                for span in line:
                    field_name = min(columns, key=lambda key: abs(columns[key] - span.x))
                    if abs(columns[field_name] - span.x) > 14:
                        raise _fail(state, page_number, f"{row['Name']}: value '{span.text}' outside all columns.")
                    parsed = _parse_value(field_name, span.text)
                    if parsed is None:
                        raise _fail(state, page_number, f"{row['Name']}: unreadable {field_name} '{span.text}'.")
                    for key, value in parsed.items():
                        if value is not None:
                            row[f"{key} {round_number}"] = value
            state.rows.append(row)
        for section in sections:
            following = [s for s in sections if s.y > section.y]
            stop = following[0].y if following else section.y + 60
            for line in _lines([s for s in below if section.y + 3 < s.y < stop]):
                texts = [s.text for s in sorted(line, key=lambda s: s.x0)]
                if len(texts) >= 2 and BIB_PATTERN.match(texts[0]) and re.search(r"[A-Za-z]{2}", texts[1]):
                    state.section_rows.append({
                        "Bib": int(texts[0]), "Name": texts[1], "Status": section.text,
                        **({"NSA": texts[2]} if len(texts) > 2 and NSA_PATTERN.match(texts[2]) else {}),
                    })
                elif texts[:1] != ["Bib"]:
                    break



CRYSTAL_PATTERN = re.compile(r"^DGN$", re.M)  # Planica 2017 (Crystal Reports, German labels)
CRYSTAL_STATUS = {"NAS": "DNS", "DIS": "DSQ", "DNS": "DNS", "DSQ": "DSQ"}


def _crystal_document(document: pymupdf.Document, state: _State) -> None:
    """German club lists (Planica FIS Cup 2017): two lines per jumper, one per round.

    Line 1: rank, name, FIS code, round 1; line 2: bib, nation, year, round 2;
    the total between both lines. 'WN' distance points, 'HN' judge points,
    'DGN' round points, the round rank right of it without label.
    """
    state.layouts.add("competition")
    state.wind_column = False
    for page_number, page in enumerate(document, start=1):
        spans = [
            Span(" ".join(s.text.split()).replace(",", ".") if re.fullmatch(r"-?\d+,\d+", s.text.strip()) else s.text.strip(),
                 s.x0, s.x1, s.y0, s.y1, s.bold)
            for s in _page_spans(page)
        ]
        dgn = next((s for s in spans if s.text == "DGN"), None)
        if dgn is None:
            continue
        header = [s for s in spans if abs(s.y - dgn.y) < 14 and s.bold]
        label = {s.text: s for s in header}
        needed = ("Speed", "Weite", "WN", "A", "B", "C", "D", "E", "HN", "DGN", "Total", "StNr")
        if any(name not in label for name in needed):
            raise _fail(state, page_number, f"unknown header {sorted(label)}.")
        columns = {
            "Speed": label["Speed"].x, "Distance": label["Weite"].x, "Distance Points": label["WN"].x,
            **{judge: label[judge].x for judge in "ABCDE"}, "Judge Points": label["HN"].x,
            "Round Total": dgn.x, "Round Rank": (dgn.x1 + label["Total"].x0) / 2,
        }
        data_low = label["Speed"].x0 - 8
        below = [s for s in spans if s.y > dgn.y + 6]
        names = sorted(
            (s for s in below if s.bold and s.x0 < data_low and looks_like_athlete(s.text)), key=lambda s: s.y
        )
        footer = min((s.y for s in below if re.match(r"Seite \d|Gedruckt:", s.text)), default=page.rect.height)
        headings = [s for s in below if s.bold and re.match(r"^(?:Ladies|Men|Women|Damen|Herren) \(", s.text)]
        for index, name in enumerate(names):
            heading = max((h for h in headings if h.y < name.y), key=lambda h: h.y, default=None)
            if heading is not None:
                state.class_group = heading.text  # ladies and men ranked apart
            if state.class_group is not None:
                state.rank_groups.setdefault(state.class_group, [])
            end = names[index + 1].y - 3 if index + 1 < len(names) else min(name.y + 22, footer - 2)
            block = [s for s in below if name.y - 3 <= s.y < end and s is not name]
            row: dict[str, Any] = {"Name": name.text}
            values = []
            for span in block:
                if span.x0 < label["StNr"].x1 + 3:
                    if RANK_PATTERN.match(span.text):
                        row["Rank"] = int(RANK_PATTERN.match(span.text).group(1))
                    elif BIB_PATTERN.match(span.text):
                        row["Bib"] = int(span.text)
                elif span.x < data_low:
                    if NSA_PATTERN.match(span.text):
                        row["NSA"] = span.text
                elif span.text in CRYSTAL_STATUS:
                    row["Status"] = CRYSTAL_STATUS[span.text]
                elif span.bold and abs(span.x - label["Total"].x) < 20 and NUMBER_PATTERN.match(span.text):
                    row["Total"] = float(span.text)
                elif span.bold and RANK_PATTERN.match(span.text) and span.x > label["Total"].x0 - 5:
                    row["Rank"] = int(RANK_PATTERN.match(span.text).group(1))
                else:
                    values.append(span)
            for round_number, line in enumerate(_lines(values), start=1):
                for span in line:
                    field_name = min(columns, key=lambda key: abs(columns[key] - span.x))
                    if abs(columns[field_name] - span.x) > 14:
                        raise _fail(state, page_number, f"{row['Name']}: value '{span.text}' outside all columns.")
                    parsed = _parse_value(field_name, span.text)
                    if parsed is None:
                        raise _fail(state, page_number, f"{row['Name']}: unreadable {field_name} '{span.text}'.")
                    for key, value in parsed.items():
                        if value is not None:
                            row[f"{key} {round_number}"] = value
            state.rows.append(row)
            if state.class_group is not None:
                state.rank_groups[state.class_group].append(row)
        if headings:
            state.class_group = max(headings, key=lambda h: h.y).text  # heading at the page end



CZECH_PATTERN = re.compile(r"stylové hodnocení")  # Harrachov 2023 (Excel, Czech labels)


def _czech_document(document: pymupdf.Document, state: _State) -> None:
    """Czech club lists (Harrachov): two lines per jumper, one per round.

    Header 'po ř.', 'st. č.', 'jméno klub', 'roč.', 'skok m', 'A'-'E',
    'm' (distance points), 'st.' (judge points), 'celk.' (round total),
    '1.k / 2.k' (round rank), 'body celk.' (total), 'žebř.' (cup points).
    """
    state.layouts.add("competition")
    state.wind_column = False
    for page_number, page in enumerate(document, start=1):
        spans = [
            Span(s.text.replace(",", ".") if re.fullmatch(r"-?\d+,\d+", s.text) else s.text,
                 s.x0, s.x1, s.y0, s.y1, s.bold)
            for s in _page_spans(page)
        ]
        letters = [s for s in spans if s.bold and len(s.text) == 1 and s.text in "ABCDE"]
        # (the judges of the jury are listed above with the same letters, one per line)
        judges = next(
            (sorted(group, key=lambda s: s.x0) for letter in letters
             if len(group := [s for s in letters if abs(s.y - letter.y) < 2]) == 5),
            [],
        )
        if not judges:
            continue
        header_y = judges[0].y
        header = sorted((s for s in spans if abs(s.y - header_y) < 4), key=lambda s: s.x0)
        metres = [s for s in header if s.text == "m"]
        totals = [s for s in header if s.text == "celk."]
        number = next((s for s in header if s.text == "č."), None)
        rank_label = next((s for s in spans if re.fullmatch(r"\d\.k", s.text) and header_y - 14 < s.y < header_y + 4), None)
        style = next((s for s in header if s.text == "st." and metres and s.x0 > metres[-1].x0), None)
        if len(metres) != 2 or len(totals) != 2 or number is None or style is None or rank_label is None:
            raise _fail(state, page_number, f"unknown Czech header {[s.text for s in header]}.")
        columns = {
            "Distance": metres[0].x, **{s.text: s.x for s in judges},
            "Distance Points": metres[1].x, "Judge Points": style.x,
            "Round Total": totals[0].x, "Round Rank": rank_label.x,
        }
        below = [s for s in spans if s.y > header_y + 5]
        bibs = sorted(
            (s for s in below if not s.bold and BIB_PATTERN.match(s.text) and abs(s.x - number.x) < 8),
            key=lambda s: s.y,
        )
        for index, bib in enumerate(bibs):
            end = bibs[index + 1].y - 3 if index + 1 < len(bibs) else bib.y + 16
            block = [s for s in below if bib.y - 3 <= s.y < end]
            name = next((s for s in block if abs(s.y - bib.y) < 2 and s.x0 > bib.x1 and s.x0 < number.x1 + 30), None)
            if name is None:
                raise _fail(state, page_number, f"no name for bib {bib.text}.")
            words = name.text.split()
            row: dict[str, Any] = {"Bib": int(bib.text), "Name": " ".join([words[0].upper(), *words[1:]])}
            for span in block:
                if span is bib or span is name:
                    continue
                if span.bold and RANK_PATTERN.match(span.text + ".") and span.x1 < number.x0:
                    row["Rank"] = int(span.text)
                elif span.bold and NUMBER_PATTERN.match(span.text) and abs(span.x - totals[1].x) < 15:
                    row["Total"] = float(span.text)
                elif not span.bold and NSA_PATTERN.match(span.text):
                    row["NSA"] = span.text
            lines = _lines([s for s in block if not s.bold and s.x0 > metres[0].x0 - 10 and s.x < rank_label.x1 + 8])
            for round_number, line in enumerate(lines, start=1):
                if any(span.text == "x" for span in line):
                    continue  # no jump in this round: zeros and rank 'x'
                for span in line:
                    field_name = min(columns, key=lambda key: abs(columns[key] - span.x))
                    if abs(columns[field_name] - span.x) > 12:
                        raise _fail(state, page_number, f"{row['Name']}: value '{span.text}' outside all columns.")
                    parsed = _parse_value(field_name, span.text + ("." if field_name == "Round Rank" else ""))
                    if parsed is None:
                        raise _fail(state, page_number, f"{row['Name']}: unreadable {field_name} '{span.text}'.")
                    for key, value in parsed.items():
                        if value is not None:
                            row[f"{key} {round_number}"] = value
            club = [s.text for s in block if not s.bold and abs(s.x0 - name.x0) < 3 and s.y > bib.y + 2]
            if club:
                row["Club"] = " ".join(club)
            state.rows.append(row)



def _is_winspringen_team(document: pymupdf.Document) -> bool:
    """WinSpringen team list (Ruhpolding CHI): header 'Rang StNr Team ... Total',
    second header line 'Namen km/h Weite ...'."""
    spans = _page_spans(document[0])
    return any(s.text.upper() == "TEAM" for s in spans) and any(s.text.upper() == "NAMEN" for s in spans)


def _winspringen_team_document(document: pymupdf.Document, state: _State) -> None:
    """Team list: team line (rank, bib, team, total), then per jumper a line
    with name and round 1 and a line with round 2 (no wind, no gate points)."""
    columns: list[_WinColumn] | None = None
    bib_p = False
    total_x = 0.0
    name_x = 0.0
    teams: list[dict[str, Any]] = []
    members: list[dict[str, Any]] = []
    member: dict[str, Any] | None = None
    group = ""
    for page_number, page in enumerate(document, start=1):
        spans = _page_spans(page)
        names_label = next((s for s in spans if s.text.upper() == "NAMEN"), None)
        if names_label is not None:
            header = sorted((s for s in spans if abs(s.y - names_label.y) < 3), key=lambda s: s.x0)
            columns = []
            bib_p = any(span.text.upper() == "STNR-P" for span in header)  # member bib '9 - 1'
            for span in header:
                label = span.text.upper()
                if label in ("NAMEN", "FIS-CODE", "STNR-P", "STNR"):
                    continue
                field_name = WINSPRINGEN_LABELS.get(label) or WINSPRINGEN_FOLLOWERS.get(
                    (label, columns[-1].field if columns else "")
                )  # 'pts' after Wind / Luke
                if field_name is None:
                    raise _fail(state, page_number, f"unknown WinSpringen team column '{span.text}'.")
                columns.append(_WinColumn(field_name, span.x0, span.x1))
            total = next((s for s in spans if s.text.upper() == "TOTAL" and names_label.y - 20 < s.y < names_label.y),
                         None)
            if total is None or not columns:
                raise _fail(state, page_number, "WinSpringen team header without Total.")
            total_x = total.x
            name_x = names_label.x0
            top = names_label.y + 3
        elif columns is None:
            continue
        else:
            top = 0.0
        data_low = min(c.x0 for c in columns) - 8
        end = min((s.y for s in spans if s.y > top and WINSPRINGEN_END.match(s.text)), default=page.rect.height)
        for line in _lines([s for s in spans if top < s.y < end]):
            line.sort(key=lambda s: s.x0)
            texts = [s.text for s in line]
            values = [s for s in line if s.x >= data_low]
            left = [s for s in line if s.x < data_low]
            if not values and len(line) == 1:
                group = line[0].text  # 'Schüler Mixed, 40.0 m Schanze': ranked apart
                continue
            if (len(left) >= 3 and RANK_VALUE_PATTERN.match(left[0].text) and BIB_PATTERN.match(left[1].text)
                    and left[0].x0 < name_x - 5):
                # Team line: rank, bib, team name; total at the right.
                total_span = min(values, key=lambda s: abs(s.x - total_x), default=None)
                if total_span is None or abs(total_span.x - total_x) > 25:
                    raise _fail(state, page_number, f"team without total {texts}.")
                name = " ".join(s.text for s in left[2:] if not NSA_PATTERN.match(s.text) or len(left) == 3)
                team = {"Rank": int(RANK_VALUE_PATTERN.match(left[0].text).group(1)),
                        "Bib": int(left[1].text), "Name": name, TEAM_KEY: True}
                if NSA_PATTERN.match(name[:3]) and (not re.search(r"[/-]", name) or re.match(r"^[A-Z]{3} - [A-Z]", name)):
                    team["NSA"] = name[:3]  # 'SUI', 'AUT I', 'SLO - SLOVENIA 1': one nation
                _winspringen_value(team, "Total", 1, total_span.text, state, page_number)
                teams.append(team)
                state.rows.append(team)
                state.rank_groups.setdefault(group, []).append(team)
                member = None
                continue
            athlete = [s for s in left if re.search(r"[^\W\d_]", s.text) and not re.fullmatch(r"\d+\s*-\s*\d+", s.text)]
            member_bib = any(re.fullmatch(r"\d+\s*-\s*\d+", s.text) for s in left)
            if athlete and bib_p and not member_bib and member is not None:
                # Line 2: club of the jumper, round 2.
                member["Club"] = " ".join(s.text for s in athlete)
                round_number = 2
            elif athlete:
                if not teams:
                    raise _fail(state, page_number, f"jumper before the first team {texts}.")
                team = teams[-1]
                number = sum(1 for row in members if row["Bib"].startswith(f"{team['Bib']}-")) + 1
                member = {"Bib": f"{team['Bib']}-{number}", "Name": " ".join(s.text for s in athlete)}
                if team.get("NSA"):
                    member["NSA"] = team["NSA"]
                members.append(member)
                state.rows.append(member)
                state.rank_groups.setdefault(group, []).append(member)
                round_number = 1
            elif values and member is not None and not left:
                round_number = 2
            else:
                raise _fail(state, page_number, f"values outside athlete rows: {texts}.")
            for span in values:
                column = min(columns, key=lambda c: abs(c.x - span.x))
                if abs(column.x - span.x) > 14:
                    raise _fail(state, page_number, f"value '{span.text}' outside all columns.")
                _winspringen_value(member, column.field, round_number, span.text, state, page_number)
        _winspringen_sections(spans, state, page_number)
    state.layouts.add("competition")
    state.blank_is_zero = True
    state.wind_column = bool(columns) and any(column.field in ("Wind", "Wind Points") for column in columns)
    fields = {column.field for column in columns or []}
    state.gate_points_missing = "Gate" in fields and "Gate Points" not in fields  # FESA 2025 RL1


def _winspringen_sections(spans: list[Span], state: _State, page: int) -> None:
    """DNS/DSQ lists below 'From the referee's protocol': bib, name, nation, club."""
    lines = _lines(spans)
    status = None
    for line in lines:
        line.sort(key=lambda s: s.x0)
        heading = next(
            (code for pattern, code in WINSPRINGEN_SECTIONS for s in line if pattern.match(s.text)),
            None,
        )
        if heading is not None:
            status = heading
            continue
        if status is None:
            continue
        texts = [s.text for s in line if not WINSPRINGEN_GROUPS.match(s.text)]
        if not texts or not BIB_PATTERN.match(texts[0]):
            status = None  # 'Total ranked:', statistics
            continue
        row: dict[str, Any] = {"Bib": int(texts[0]), "Status": status}
        rest = texts[1:]
        nsa_index = next((i for i, text in enumerate(rest) if NSA_PATTERN.match(text)), None)
        if nsa_index == 0 or not rest or (nsa_index is None and not looks_like_athlete(rest[0])):
            raise _fail(state, page, f"unreadable {status} row {texts}.")
        if nsa_index is None:
            row["Name"] = rest[0]  # without nation: bib, name, club
            others = rest[1:]
        else:
            row["Name"] = " ".join(rest[:nsa_index])
            row["NSA"] = rest[nsa_index]
            others = rest[nsa_index + 1:]
        if others:
            row["Club"] = others[0]
        if len(others) > 1:
            row["Remarks"] = " ".join(others[1:])
        state.section_rows.append(row)


PAGE_DATE_PATTERN = re.compile(r"\b(\d{1,2}) ([A-Z]{3}) (\d{4})\b")


def _page_date(page: pymupdf.Page) -> date | None:
    """'SAT 24 JAN 2026' in the head of a page."""
    head = "\n".join(page.get_text(sort=True).splitlines()[:4])
    match = PAGE_DATE_PATTERN.search(head)
    month = MONTHS.get(match.group(2)) if match else None
    return date(int(match.group(3)), month, int(match.group(1))) if month else None


def _rounds_by_day(day_rows: dict[date, list[dict[str, Any]]], state: _State) -> None:
    """Trial rounds of two days in one PDF (Ski Flying WCh 2026), each with
    its own bibs and ranks: ranked apart per day."""
    days = [day for day, rows in day_rows.items() if rows]
    if len(days) > 1 and not state.rank_groups:
        state.rank_groups.update((day.isoformat(), day_rows[day]) for day in sorted(days))


PAGE_GENDER_PATTERN = re.compile(r"\b(LADIES|WOMEN|MEN)\b")


def _page_gender(page: pymupdf.Page) -> str | None:
    """'SKI JUMPING LADIES' / 'SKI JUMPING MEN' in the head of a page."""
    head = "\n".join(page.get_text(sort=True).splitlines()[:6])
    match = PAGE_GENDER_PATTERN.search(head)
    if match is None:
        return None
    return "M" if match.group(1) == "MEN" else "W"


def extract_results(pdf_path: Path) -> list[dict[str, Any]]:
    """Return one dict per athlete (keys from RESULT_COLUMNS) for a PDF."""
    state = _State(pdf_path.name)
    try:
        with pymupdf.open(pdf_path) as document:
            text = "\n".join(page.get_text() for page in document)
            if KILIPA_PATTERN.search(text):
                _kilipa_document(document, state)
            elif CZECH_PATTERN.search(text):
                _czech_document(document, state)
            elif CARPATH_PATTERN.search(text):
                _carpath_document(document, state)
            elif CRYSTAL_PATTERN.search(text) and re.search(r"^Weite$", text, re.M):
                _crystal_document(document, state)
            elif WINSPRINGEN_PATTERN.search(text) and _is_winspringen_team(document):
                _winspringen_team_document(document, state)
            elif WINSPRINGEN_PATTERN.search(text):
                _winspringen_document(document, state)
            else:
                if "FastReport" in (document.metadata or {}).get("producer", "") or (
                    # (2022: converted by LibreOffice)
                    re.search(r"^km/h$", text, re.M) and re.search(r"^\d+,\d+$", text, re.M)
                ):
                    state.fastreport = True
                    text = re.sub(r"(\d),(\d)", r"\1.\2", text)  # 'METER VALUE 2,4 points/m'
                page_groups: dict[str, list[dict[str, Any]]] = {}
                day_rows: dict[date, list[dict[str, Any]]] = {}
                parts: list[list[dict[str, Any]]] = []
                for page_number, page in enumerate(document, start=1):
                    before = len(state.rows)
                    if not parts or (state.rows and any(
                        # (the jury at the top of the page, not below the last results)
                        rect.y0 < page.rect.height * 0.3
                        for rect in page.search_for("Jury/Competition Management")
                    )):
                        parts.append([])  # Park City 2017: two lists in one PDF, each with its jury
                    _parse_page(page, page_number, state)
                    parts[-1].extend(state.rows[before:])
                    gender = _page_gender(page)
                    if gender is not None:
                        page_groups.setdefault(gender, []).extend(state.rows[before:])
                    day = _page_date(page)
                    if day is not None:
                        day_rows.setdefault(day, []).extend(state.rows[before:])
                # Mixed training (WSC Seefeld 2019): pages of the ladies, then of
                # the men, each ranked apart.
                if len(page_groups) > 1 and not state.rank_groups:
                    state.rank_groups.update(page_groups)
                if len([part for part in parts if part]) > 1 and not state.rank_groups:
                    state.rank_groups.update((f"part {index}", part) for index, part in enumerate(parts) if part)
                _rounds_by_day(day_rows, state)
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise ResultsParseError(f"{pdf_path.name}: cannot read PDF: {error}") from error
    if state.fastreport:
        # Russian championships (Chaikovsky): regions instead of nations.
        for row in state.rows + state.section_rows:
            row.setdefault("NSA", "RUS")
            if row.get("NSA") is None:
                row["NSA"] = "RUS"
    # Ranked after round 1, then DSQ/DNS in round 2: listed in both places.
    ranked = {row.get("Bib"): row for row in state.rows if row.get("Bib") is not None}
    remaining = []
    for row in state.section_rows:
        same = ranked.get(row.get("Bib"))
        if same is not None and (same.get("Name") == row.get("Name") or (
            # Olympics 2022: 'GRANERUD HE' for 'GRANERUD Halvor Egner'
            same.get("Status") == row.get("Status")
            and str(same.get("Name", "")).split(" ")[0] == str(row.get("Name", "")).split(" ")[0]
        )):
            same.setdefault("Status", row["Status"])
            if row.get("Remarks") and row["Remarks"] not in same.get("Remarks", ""):
                same["Remarks"] = " ".join(filter(None, (same.get("Remarks"), row["Remarks"])))
        else:
            remaining.append(row)
    # The same athlete in two lists (DSQ in both rounds, DNS and NPS): one row.
    merged: dict[tuple[Any, Any], dict[str, Any]] = {}
    for row in remaining:
        key = (row.get("Bib"), row.get("Name"))
        if key[0] is None or key not in merged:
            merged[key if key[0] is not None else (id(row), None)] = row
            continue
        first = merged[key]
        extra = [text for text in (
            row["Status"] if row.get("Status") != first.get("Status") else None, row.get("Remarks"),
        ) if text and text not in first.get("Remarks", "")]
        if extra:
            first["Remarks"] = " ".join(filter(None, (first.get("Remarks"), *extra)))
    state.section_rows = list(merged.values())
    _validate(state, text)
    return state.rows + state.section_rows


HEADER_FILL = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
ERROR_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")


def write_results_sheet(
    workbook: Workbook, sheet_name: str, rows: list[dict[str, Any]]
) -> None:
    """Add a sheet with one row per athlete and the fixed RESULT_COLUMNS
    (plus the columns of rounds 3 and 4 when the PDF has them)."""
    worksheet = workbook.create_sheet(sheet_name[:31])
    extra_rounds = [
        number for number in range(ROUNDS + 1, MAX_ROUNDS + 1)
        if any(f"{name} {number}" in row for row in rows for name in ROUND_FIELDS)
    ]
    total_index = RESULT_COLUMNS.index("Total")
    columns = (
        *RESULT_COLUMNS[:total_index],
        *(f"{name} {number}" for number in extra_rounds for name in ROUND_FIELDS),
        *RESULT_COLUMNS[total_index:],
    )
    for column, key in enumerate(columns, start=1):
        cell = worksheet.cell(row=1, column=column, value=key)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        letter = cell.column_letter
        worksheet.column_dimensions[letter].width = {
            "Name": 26, "Club": 30, "Remarks": 50, "Date of Birth": 12, "Status": 13,
        }.get(key, 9)
    for row_number, values in enumerate(rows, start=2):
        for column, key in enumerate(columns, start=1):
            value = values.get(key)
            if value is None:
                continue
            cell = worksheet.cell(row=row_number, column=column, value=value)
            base_key = re.sub(r" \d+$", "", key)
            number_format = NUMBER_FORMATS.get(base_key)
            if number_format:
                cell.number_format = number_format
    worksheet.freeze_panes = "D2"
    worksheet.auto_filter.ref = worksheet.dimensions


def write_error_sheet(workbook: Workbook, sheet_name: str, message: str) -> None:
    """Add a sheet that only shows why the results could not be read."""
    worksheet = workbook.create_sheet(sheet_name[:31])
    worksheet.sheet_properties.tabColor = "FFFF0000"
    cell = worksheet.cell(row=1, column=1, value=f"Results not imported: {message}")
    cell.fill = ERROR_FILL
    cell.font = Font(bold=True)
