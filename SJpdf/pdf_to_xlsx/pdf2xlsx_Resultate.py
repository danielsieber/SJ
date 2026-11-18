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

STATUS_CODES = ("DNS", "DSQ", "DNF", "NPS", "DQ")
SECTION_HEADINGS = (
    (re.compile(r"^Disqualified\b", re.I), "DSQ"),
    (re.compile(r"^Did not start\b", re.I), "DNS"),
    (re.compile(r"^Did not finish\b", re.I), "DNF"),
    (re.compile(r"^Not permitted to start\b", re.I), "NPS"),
)
END_HEADINGS = re.compile(
    r"^(?:Jury\b|Competition Management\b|Competition\s*/\s*Weather\b|"
    r"Statistics\b|Legend\b|Weather Information\b)",
    re.I,
)
FOOTER_PATTERN = re.compile(
    r"^(?:\d{1,2} [A-Z]{3} \d{4} / |(?:Timing and )?Data Service by\b|"
    r"(?:www\.)?fis-ski\.com\b|Report Created\b|Timing Sponsor\b|Data Provider\b)",
    re.I,
)
# Text inside a table that belongs to no athlete: group headings of team
# lists, explanations of disqualification rules.
NO_ROW_PATTERN = re.compile(r"^(?:Women|Men|Mixed|-{3,})$|^SCE\s?\d[\d.]*\s*:", re.I)
NOT_QUALIFIED_PATTERN = re.compile(r"^Not Qualified\b", re.I)
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
ROUND_NOTE_PATTERN = re.compile(r"^(?:\d(?:st|nd|rd|th)|Final) Rnd$")  # super team, total column
TEAM_NAME_PATTERN = re.compile(r"^([A-Z]{3}) - \S")  # 'GER - Germany'
GENDER_SUFFIX = re.compile(r"\s*\((?:W|M)\)$")
NAME_TEXT_PATTERN = re.compile(r"^[^\W\d_][^\W\d_'’.\- ]*(?:['’.\- ]+[^\W\d_]+)*\.?$")  # letters only
TEAM_KEY = "_team"  # team row (not a column): its members follow below
# Numbers of neighbouring columns in one span; Olympic training: distance and gate.
SPLIT_NUMBER_PATTERN = re.compile(r"^-?\d+\.\d+$|^-?\d{1,3}$|^©?\d{1,2}[➔→]©?\d{1,2}$")
SCE_PATTERN = re.compile(r"^SCE\s?\d")  # disqualification (rule code) inside a row
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
    rows: list[dict[str, Any]] = field(default_factory=list)
    section_rows: list[dict[str, Any]] = field(default_factory=list)
    layouts: set[str] = field(default_factory=set)
    team: dict[str, Any] | None = None  # last team row (team events)
    wind_column: bool = True  # False: small hill lists without wind / gate
    blank_is_zero: bool = False  # WinSpringen: no distance / wind points printed means 0
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
    texts = [span.text.upper() for span in row]
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
        if label.text.upper() in ("NSA", "NOC", "CODE") and header_y - 25 < label.y < header_y + 3
        and left <= label.x0 and label.x1 < speed.x0
    ]
    if nation:
        data_low = max(data_low, max(label.x1 for label in nation) + 1)

    competition = _match_labels(row[: len(COMPETITION_LABELS)], COMPETITION_LABELS, 1)
    if competition is not None:
        rest = row[len(COMPETITION_LABELS) :]
        has_total = any(
            s.text.upper() == "TOTAL" and header_y - 25 < s.y < header_y + 3 and left <= s.x0 < right
            for s in spans
        )
        if not has_total and rest and rest[-1].text.upper() == "POINTS" and rest[-1].y < header_y - 1:
            # KO qualification (Four Hills): the total column is called POINTS,
            # printed one header row higher.
            total = [rest.pop()]
            row = row[: len(COMPETITION_LABELS) + len(rest)]
        else:
            total = []
        if rest:
            # Round total with or without round rank; team events: POINTS.
            labels = COMPETITION_ROUND_LABELS[: len(rest)]
            if rest[0].text.upper() == "POINTS":
                labels = (("POINTS", "Round Total"), *labels[1:])
            extra = _match_labels(rest, labels, 1)
            if extra is None:
                raise _fail(state, page, f"unknown result table header {texts}.")
            competition += extra
        total = total or [
            s for s in spans
            if s.text.upper() == "TOTAL"
            and header_y - 25 < s.y < header_y + 3
            and s.x0 > row[-1].x1 + 2
            and left <= s.x0 < right
        ]
        if len(total) != 1:
            raise _fail(state, page, "cannot find the TOTAL column.")
        total_low = (row[-1].x1 + total[0].x0) / 2
        columns = _columns(competition, data_low, total_low)
        return Table("competition", left, right, header_y, data_low, [columns], total_low=total_low)

    groups = len(row) // len(TRAINING_LINE1_LABELS)
    line1 = _match_labels(row, TRAINING_LINE1_LABELS, groups) if groups else None
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
    line2 = _match_labels(rank_row, TRAINING_LINE2_LABELS, groups)
    if line2 is None and rank_row:
        # Labels of the second line misplaced in the PDF: its values stand
        # below the first line labels (speed rank below speed, ...).
        line2 = [
            (span, f"{TRAINING_LINE2_LABELS[index % len(TRAINING_LINE2_LABELS)][1]}#{group}")
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
    for x, top, bottom in _vertical_segments(page):
        if table.left - 16 <= x <= table.left + (30 if team else 1) and table.header_y + 2 < top < y_end:
            # Team jumper without jump (DNS): a row of one line only.
            if bottom - top > (10 if x <= table.left + 1 else 6):
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
    cells: list[tuple[float, float]], spans: list[Span], table: Table
) -> list[tuple[float, float]]:
    """Team events: one row per team and per jumper, starting at the row
    line just above its team name or bib (super team: a jumper has one line
    per round, each with its own row lines)."""
    anchors = sorted(
        span.y for span in spans
        if span.x < table.data_low
        and (TEAM_NAME_PATTERN.match(span.text) or TEAM_BIB_PATTERN.match(span.text))
    )
    if not anchors:
        return cells
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
    if (first_value is not None and anchors[0] - 8 < first_value < anchors[0] - 2) or names_above:
        # Olympic layout: the round lines stand above and below the name, the
        # row border lies halfway between two names.
        borders = [(above + below) / 2 for above, below in zip(anchors, anchors[1:])]
        top = min(top for top, _bottom in cells)
        bottom = max(bottom for _top, bottom in cells)
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
        if span.x < table.data_low and (looks_like_athlete(span.text) or TEAM_NAME_PATTERN.match(span.text)):
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


def _identity(spans: list[Span]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    team = next((s for s in spans if TEAM_NAME_PATTERN.match(s.text)), None)
    if team is not None:
        # Team row: 'GER - Germany', rank, team bib; the jumpers follow.
        values.update(Name=team.text, NSA=team.text[:3], **{TEAM_KEY: True})
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
        elif TEAM_BIB_PATTERN.match(span.text):
            values["Bib"] = span.text.replace(" ", "")  # team member '8-1': team 8, first jumper
        elif _date(span.text):
            values["Date of Birth"] = _date(span.text)
        elif span.bold and not plain:
            name_parts.append(span.text)
        else:
            club_parts.append(span.text)
    if name_parts:
        # Mixed team: 'FREITAG Selina (W)'
        values["Name"] = GENDER_SUFFIX.sub("", " ".join(name_parts))
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
    right_spans = [s for s in right_spans if not ROUND_NOTE_PATTERN.match(s.text)]
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
        round_numbers = [None] * len(lines)

    for line, columns, round_number in zip(lines, line_columns, round_numbers):
        cells: dict[str, list[Span]] = {}
        for span in line:
            column = next((c for c in columns if c.low <= span.x < c.high), None)
            if column is None:
                raise _fail(state, page, f"{label}: value '{span.text}' outside all columns.")
            cells.setdefault(column.field, []).append(span)
        for column_field, cell_spans in cells.items():
            field_name, group = column_field.split("#")
            number = round_number if round_number is not None else int(group)
            cell_spans.sort(key=lambda s: s.x0)
            if len(cell_spans) > 1 and field_name != "Gate":
                raise _fail(
                    state, page,
                    f"{label}: several values in {field_name}: {[s.text for s in cell_spans]}.",
                )
            text = "".join(s.text for s in cell_spans)
            if text in STATUS_CODES:
                row["Status"] = text
                continue
            parsed = _parse_value(field_name, text)
            if parsed is None:
                raise _fail(state, page, f"{label}: unreadable {field_name} '{text}'.")
            for key, value in parsed.items():
                if value is not None:
                    row[f"{key} {number}"] = value
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
        if match is None and re.match(r"^SCE\b", texts[0]):
            # Swiss Timing explains the rule codes below the list.
            explanation = " ".join(texts)
            code = explanation.split(" - ")[0].strip()
            explanations[code] = explanation
            continue
        if match is None:
            break  # end of the list (legend, statistics, ...): no bib
        rest = ([match.group(2)] if match.group(2) else []) + texts[1:]
        # Nation and rule in one span: 'NOR SCE4 - Suit'
        rest = [
            part for text in rest
            for part in (text.split(" ", 1) if re.match(r"^[A-Z]{3} SCE", text) else [text])
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


def _parse_page(page: pymupdf.Page, page_number: int, state: _State) -> None:
    spans = _page_spans(page)
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
        used: set[int] = set()
        team = any(TEAM_NAME_PATTERN.match(s.text) for s in table_spans)
        cells = _row_cells(page, table, table_end, team)
        name_cells = _name_cells(table_spans, table, table_end)
        if not cells or (not team and len(name_cells) > 1 and len(cells) > 1.5 * len(name_cells)):
            # No row lines, or lines of a neighbouring table (half rows).
            cells = name_cells
        elif team:
            cells = _team_cells(cells, table_spans, table)
        for top, bottom in cells:
            block = [s for s in table_spans if top <= s.y < bottom]
            used.update(id(s) for s in block)
            row = _parse_block(block, table, state, page_number)
            if row is None:
                continue
            if state.not_qualified or (not_qualified_y is not None and top > not_qualified_y):
                row.setdefault("Status", "Not Qualified")
            if row.get(TEAM_KEY):
                state.team = row
            elif isinstance(row.get("Bib"), str) and not row.get("NSA"):
                # Jumper of the team above: nation of the team.
                if state.team is None or not row["Bib"].startswith(f"{state.team.get('Bib')}-"):
                    raise _fail(state, page_number, f"{row['Name']}: team of bib {row['Bib']} not found.")
                row.setdefault("NSA", state.team["NSA"])
            state.rows.append(row)
        if not_qualified_y is not None:
            state.not_qualified = True
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
    meter = re.search(r"\bMET(?:ER|RE) VALUE\s*:?\s*(\d+(?:\.\d+)?)\b", normalized, re.I)
    return (
        float(k_point.group(1)) if k_point else None,
        float(meter.group(1)) if meter else None,
    )


def _check_ranks(
    rows: list[dict[str, Any]], value_key: str, rank_key: str, errors: list[str], dense: bool = False
) -> None:
    """Ranks follow the values: higher value, better rank; same value, same rank.

    Ranks may have gaps: an athlete ranked after round 1 and disqualified
    later is no longer listed with the round 1 result.
    """
    ranked = [row for row in rows if row.get(value_key) is not None and row.get(rank_key) is not None]
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
    if k_point is None or meter_value is None:
        errors.append("K-point or meter value not found")
    base = 120.0 if (k_point or 0) >= 170 else 60.0
    if "competition" in state.layouts:
        for row in rows:
            for number in range(1, MAX_ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                distance_points = row.get(f"Distance Points {number}")
                if distance is None or distance_points is None or k_point is None or meter_value is None:
                    continue
                calculated = base + (distance - k_point) * meter_value
                if abs(calculated - distance_points) > TOLERANCE:
                    errors.append(f"{row['Name']}: distance points round {number}")
        for row in rows:
            round_sum = 0.0
            rounds = 0
            for number in range(1, MAX_ROUNDS + 1):
                if row.get(f"Distance {number}") is None:
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
                if state.blank_is_zero:
                    parts = [0.0 if part is None and index in (0, 3) else part
                             for index, part in enumerate(parts)]
                if any(part is None for part in parts):
                    errors.append(f"{row['Name']}: incomplete round {number}")
                    continue
                round_total = row.get(f"Round Total {number}")
                calculated = max(0.0, sum(parts))  # a round never counts below 0 points
                if round_total is not None and abs(calculated - round_total) > TOLERANCE:
                    errors.append(f"{row['Name']}: round {number} total {calculated:.1f} != {round_total}")
                round_sum += round_total if round_total is not None else calculated
            if row.get(TEAM_KEY) or isinstance(row.get("Bib"), str):
                continue  # team events: checked below
            if row.get("Total") is not None and abs(round_sum - row["Total"]) > TOLERANCE * max(rounds, 1):
                errors.append(f"{row['Name']}: total {round_sum:.1f} != {row['Total']}")
            if row.get("Total") is None and row.get("Status") not in STATUS_CODES:
                errors.append(f"{row['Name']}: missing total")
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
            _check_ranks(teams, "Total", "Rank", errors)
        else:
            # KO: winners and lucky losers ranked apart from the other losers.
            groups = (
                [
                    [row for row in rows if row.get("Status") != "Not Qualified"],
                    [row for row in rows if row.get("Status") == "Not Qualified"],
                ]
                if KO_PATTERN.search(text[:2000])
                else list(state.rank_groups.values()) or [rows]
            )
            for group in groups:
                _check_ranks(group, "Total", "Rank", errors)
                for number in range(1, MAX_ROUNDS + 1):
                    _check_ranks(group, f"Round Total {number}", f"Round Rank {number}", errors)
    if "training" in state.layouts:
        for row in rows:
            for number in range(1, MAX_ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                points = row.get(f"Round Total {number}")
                if distance is None or points is None or k_point is None or meter_value is None:
                    continue
                calculated = max(
                    0.0,
                    base
                    + (distance - k_point) * meter_value
                    + (row.get(f"Gate Points {number}") or 0.0)
                    + (row.get(f"Wind Points {number}") or 0.0),
                )
                if abs(calculated - points) > TOLERANCE:
                    errors.append(f"{row['Name']}: round {number} points {calculated:.1f} != {points}")
        for group in list(state.rank_groups.values()) or [rows]:
            for number in range(1, MAX_ROUNDS + 1):
                for value_key, rank_key in (("Speed", "Speed Rank"), ("Distance", "Distance Rank"),
                                            ("Round Total", "Round Rank")):
                    _check_ranks(group, f"{value_key} {number}", f"{rank_key} {number}", errors,
                                 dense=state.dense_ranks)
    bibs = [row["Bib"] for row in rows + state.section_rows if row.get("Bib") is not None]
    duplicates = sorted({bib for bib in bibs if bibs.count(bib) > 1})
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
    "PP": "Cup Points",  # cup points of the place: not read
}
WINSPRINGEN_FOLLOWERS = {  # 'pts' / 'Rg' take the meaning of the column before
    ("PTS", "Wind"): "Wind Points",
    ("PTS", "Gate"): "Gate Points",
    ("RG", "Speed"): "Speed Rank",
    ("RG", "Distance"): "Distance Rank",
}
WINSPRINGEN_IDENTITY = ("Rank", "Bib", "FIS Code", "Name", "Year", "Association", "Club")
WINSPRINGEN_END = re.compile(
    r"^(?:From the referee|Statisti[ck]s?\b|Jury\b|WinSpringen\b|Erstellt am\b|Seite \d|"
    r"For the correctness|Weather\b|Wetter\b)",
    re.I,
)
WINSPRINGEN_SECTIONS = (
    (re.compile(r"^(?:Did not start|Nicht gestartet)\b", re.I), "DNS"),
    (re.compile(r"^(?:Disqualified|Disqualifiziert)\b", re.I), "DSQ"),
    (re.compile(r"^(?:Did not finish|Nicht beendet)\b", re.I), "DNF"),
)
WINSPRINGEN_GROUPS = re.compile(r"^(?:Men|Women|Ladies|Mixed|Youth|Herren|Damen)$", re.I)


@dataclass
class _WinColumn:
    field: str
    x0: float
    x1: float

    @property
    def x(self) -> float:
        return (self.x0 + self.x1) / 2


def _winspringen_header(spans: list[Span], state: _State, page: int) -> list[_WinColumn] | None:
    bib_labels = [s for s in spans if s.text.upper() == "STNR"]
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
            top = next(s.y for s in spans if s.text.upper() == "STNR") + 9
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
    if not any(c.field == "Wind" for c in layout.columns):
        state.wind_column = False

    blocks: list[tuple[str, list[tuple[_WinPage, list[Span]]]]] = []
    group = ""
    for page_layout, line in lines:
        first = next((s for s in line if page_layout.identity_field(s) == page_layout.start_field), None)
        if first is not None and BIB_PATTERN.match(first.text):
            blocks.append((group, [(page_layout, line)]))
        elif any(s.x >= page_layout.data_low for s in line):
            if not blocks:
                raise _fail(state, page_layout.number, f"values outside athlete rows: {[s.text for s in line]}.")
            blocks[-1][1].append((page_layout, line))
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
        row["Name"] = GENDER_SUFFIX.sub("", " ".join(names))
        if clubs:
            row["Club"] = " ".join(clubs)
        if (not layout.training and row.get("Total") is None
                and not any(c.field == "Total" for c in layout.columns)):
            row["Total"] = row.get("Round Total 1")  # after round 1: no total column
        state.rows.append(row)
        state.rank_groups.setdefault(group, []).append(row)


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


def extract_results(pdf_path: Path) -> list[dict[str, Any]]:
    """Return one dict per athlete (keys from RESULT_COLUMNS) for a PDF."""
    state = _State(pdf_path.name)
    try:
        with pymupdf.open(pdf_path) as document:
            text = "\n".join(page.get_text() for page in document)
            if WINSPRINGEN_PATTERN.search(text):
                _winspringen_document(document, state)
            else:
                for page_number, page in enumerate(document, start=1):
                    _parse_page(page, page_number, state)
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise ResultsParseError(f"{pdf_path.name}: cannot read PDF: {error}") from error
    # Ranked after round 1, then DSQ/DNS in round 2: listed in both places.
    ranked = {row.get("Bib"): row for row in state.rows if row.get("Bib") is not None}
    remaining = []
    for row in state.section_rows:
        same = ranked.get(row.get("Bib"))
        if same is not None and same.get("Name") == row.get("Name"):
            same.setdefault("Status", row["Status"])
            if row.get("Remarks") and row["Remarks"] not in same.get("Remarks", ""):
                same["Remarks"] = " ".join(filter(None, (same.get("Remarks"), row["Remarks"])))
        else:
            remaining.append(row)
    state.section_rows = remaining
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
