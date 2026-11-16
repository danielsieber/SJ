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


ROUNDS = 2
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
    r"(?:www\.)?fis-ski\.com\b|Report Created\b)",
    re.I,
)
NOT_QUALIFIED_PATTERN = re.compile(r"^Not Qualified\b", re.I)
RANK_PATTERN = re.compile(r"^=?(\d+)\.$")
RANK_VALUE_PATTERN = re.compile(r"^=?(\d+)\.?$")
BIB_PATTERN = re.compile(r"^\d+$")
NSA_PATTERN = re.compile(r"^[A-Z]{3}$")
DOB_PATTERN = re.compile(r"^(\d{1,2}) ([A-Z]{3}) (\d{4})$")
NUMBER_PATTERN = re.compile(r"^-?\d+(?:\.\d+)?$")
DISTANCE_PATTERN = re.compile(r"^(\*)?(\d+(?:\.\d+)?)$")
GATE_PATTERN = re.compile(r"^(©)?(\d+)(?:[➔→](©)?(\d+))?$")
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


@dataclass
class _State:
    pdf_name: str
    not_qualified: bool = False
    rows: list[dict[str, Any]] = field(default_factory=list)
    section_rows: list[dict[str, Any]] = field(default_factory=list)
    layouts: set[str] = field(default_factory=set)


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
                spans.append(Span(text, x0, x1, y0, y1, bool(span["flags"] & 16)))
    return spans


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


def _same_row(spans: Iterable[Span], y: float, left: float, right: float) -> list[Span]:
    return sorted(
        (span for span in spans if abs(span.y - y) < 1.6 and left <= span.x0 < right),
        key=lambda span: span.x0,
    )


def _match_labels(
    row: list[Span], expected: tuple[tuple[str, str], ...], repeat: int
) -> list[tuple[Span, str]] | None:
    texts = [span.text for span in row]
    expected_texts = [label for label, _ in expected] * repeat
    if texts != expected_texts:
        return None
    result = []
    for index, span in enumerate(row):
        field_name = expected[index % len(expected)][1]
        group = index // len(expected) + 1
        result.append((span, f"{field_name}#{group}"))
    return result


def _find_tables(spans: list[Span], page_width: float, state: _State, page: int) -> list[Table]:
    speed_labels = sorted((s for s in spans if s.text == "[km/h]"), key=lambda s: s.x0)
    if not speed_labels:
        return []
    header_y = speed_labels[0].y
    if any(abs(label.y - header_y) > 1.6 for label in speed_labels):
        raise _fail(state, page, "table headers on different heights.")

    # Each table starts at its RANK/BIB header left of the first speed column.
    starts = sorted(
        (
            s for s in spans
            if s.bold
            and re.match(r"^(RANK|BIB)\b", s.text)
            and header_y - 25 < s.y < header_y + 3
        ),
        key=lambda s: s.x0,
    )
    # Several header labels of one table may start with RANK/BIB; a new
    # table only starts where a RANK/BIB label follows a speed column.
    lefts: list[float] = []
    for start in starts:
        if not lefts or any(lefts[-1] < speed.x0 < start.x0 for speed in speed_labels):
            lefts.append(start.x0)
    if not lefts or lefts[0] > speed_labels[0].x0:
        raise _fail(state, page, "cannot find the left edge of a result table.")
    tables = []
    for index, left in enumerate(lefts):
        right = lefts[index + 1] if index + 1 < len(lefts) else page_width
        speed = next((s for s in speed_labels if left < s.x0 < right), None)
        if speed is None:
            raise _fail(state, page, "result table without speed column.")
        tables.append(_build_table(spans, speed, left, right, state, page))
    return tables


def _build_table(
    spans: list[Span], speed: Span, left: float, right: float, state: _State, page: int
) -> Table:
    header_y = speed.y
    row = [s for s in _same_row(spans, header_y, left, right) if s.x0 >= speed.x0 - 1]
    # Swiss Timing training: the POINTS labels sit one header row higher.
    for span in spans:
        if (
            span.text == "POINTS"
            and header_y - 12 < span.y < header_y - 3
            and left <= span.x0 < right
            and not any(other.x0 < span.x1 and span.x0 < other.x1 for other in row)
        ):
            row.append(span)
    row.sort(key=lambda s: s.x0)
    texts = [s.text for s in row]
    data_low = speed.x0 - 12

    competition = _match_labels(row[: len(COMPETITION_LABELS)], COMPETITION_LABELS, 1)
    if competition is not None:
        rest = row[len(COMPETITION_LABELS) :]
        if rest:
            extra = _match_labels(rest, COMPETITION_ROUND_LABELS, 1)
            if extra is None:
                raise _fail(state, page, f"unknown result table header {texts}.")
            competition += extra
        total = [
            s for s in spans
            if s.text == "TOTAL"
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
        (s.y for s in spans if s.text == "RANK" and header_y + 3 < s.y < header_y + 14
         and speed.x0 - 5 <= s.x0 < right),
        None,
    )
    rank_row = (
        [s for s in _same_row(spans, rank_row_y, left, right) if s.x0 >= speed.x0 - 5]
        if rank_row_y is not None
        else []
    )
    line2 = _match_labels(rank_row, TRAINING_LINE2_LABELS, groups)
    if line2 is None:
        raise _fail(state, page, f"unknown training header {[s.text for s in rank_row]}.")
    remarks = [
        s for s in spans
        if s.text == "REMARKS" and header_y - 25 < s.y < header_y + 3 and left <= s.x0 < right
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
    page: pymupdf.Page, table: Table, y_end: float
) -> list[tuple[float, float]]:
    """Return athlete rows from the vertical lines at the table's left edge."""
    tops: list[tuple[float, float]] = []
    for x, top, bottom in _vertical_segments(page):
        if table.left - 16 <= x <= table.left + 1 and top > table.header_y + 2 and top < y_end:
            if bottom - top > 10:
                tops.append((top, min(bottom, y_end)))
    tops.sort()
    cells: list[tuple[float, float]] = []
    for top, bottom in tops:
        if cells and abs(cells[-1][0] - top) < 2:
            cells[-1] = (cells[-1][0], max(cells[-1][1], bottom))
        else:
            cells.append((top, bottom))
    return cells


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
    name_parts = []
    club_parts = []
    nsa = [s for s in spans if s.bold and NSA_PATTERN.match(s.text)]
    nsa_span = max(nsa, key=lambda s: s.x0) if nsa else None
    for span in sorted(spans, key=lambda s: (round(s.y), s.x0)):
        if span is nsa_span:
            values["NSA"] = span.text
        elif span.text == "U23":
            values["U23"] = "U23"
        elif RANK_PATTERN.match(span.text):
            values["Rank"] = int(RANK_PATTERN.match(span.text).group(1))
        elif BIB_PATTERN.match(span.text):
            values["Bib"] = int(span.text)
        elif _date(span.text):
            values["Date of Birth"] = _date(span.text)
        elif span.bold:
            name_parts.append(span.text)
        else:
            club_parts.append(span.text)
    if name_parts:
        values["Name"] = " ".join(name_parts)
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
        if not match:
            return None
        coach_before, before, coach_after, after = match.groups()
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
    identity_spans = [s for s in spans if s.x < table.data_low]
    if table.total_low is not None:
        right_spans = [s for s in spans if s.x >= table.total_low]
    else:
        right_spans = []
    remark_spans = (
        [s for s in spans if s.x0 >= table.remarks_low] if table.remarks_low is not None else []
    )
    data_spans = [
        s for s in spans
        if s not in identity_spans and s not in right_spans and s not in remark_spans
    ]
    row = _identity(identity_spans)
    if "Name" not in row:
        if any(NOT_QUALIFIED_PATTERN.match(s.text) for s in spans):
            return None
        if not spans:
            return None
        raise _fail(state, page, f"row without athlete name: {[s.text for s in spans]}.")
    label = row["Name"]

    if right_spans:
        if len(right_spans) != 1 or not NUMBER_PATTERN.match(right_spans[0].text):
            raise _fail(state, page, f"{label}: unreadable total {[s.text for s in right_spans]}.")
        row["Total"] = float(right_spans[0].text)
    if remark_spans:
        row["Remarks"] = " ".join(s.text for s in sorted(remark_spans, key=lambda s: (s.y, s.x0)))

    lines = _lines(data_spans)
    if table.layout == "competition":
        if len(lines) > ROUNDS:
            raise _fail(state, page, f"{label}: more than {ROUNDS} rounds.")
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
            raise _fail(state, page, f"unreadable {status} row {texts}.")
        rest = ([match.group(2)] if match.group(2) else []) + texts[1:]
        nsa_index = next((i for i, text in enumerate(rest) if NSA_PATTERN.match(text)), None)
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
    headings = sorted(
        (
            (s.y, status)
            for s in spans
            if s.bold
            for pattern, status in SECTION_HEADINGS
            if pattern.match(s.text)
        ),
    )
    end_marks = sorted(
        [s.y for s in spans if s.bold and END_HEADINGS.match(s.text)]
        + [s.y for s in spans if FOOTER_PATTERN.match(s.text)]
        + [y for y, _ in headings]
    )

    for table in tables:
        state.layouts.add(table.layout)
        table_end = min((y for y in end_marks if y > table.header_y), default=page.rect.height)
        table_spans = [
            s for s in spans
            if table.left - 2 <= s.x0 < table.right and table.header_y + 3 < s.y < table_end
        ]
        not_qualified_y = min(
            (s.y for s in table_spans if NOT_QUALIFIED_PATTERN.match(s.text)), default=None
        )
        used: set[int] = set()
        for top, bottom in _row_cells(page, table, table_end):
            block = [s for s in table_spans if top <= s.y < bottom]
            used.update(id(s) for s in block)
            row = _parse_block(block, table, state, page_number)
            if row is None:
                continue
            if state.not_qualified or (not_qualified_y is not None and top > not_qualified_y):
                row.setdefault("Status", "Not Qualified")
            state.rows.append(row)
        if not_qualified_y is not None:
            state.not_qualified = True
        lost = [
            s.text for s in table_spans
            if id(s) not in used and not NOT_QUALIFIED_PATTERN.match(s.text)
        ]
        if lost:
            raise _fail(state, page_number, f"values outside athlete rows: {lost[:8]}.")

    for y, status in headings:
        end = min((mark for mark in end_marks if mark > y + 1), default=page.rect.height)
        state.section_rows.extend(_section_rows(spans, y + 1, end, status, state, page_number))


def _hill_values(text: str) -> tuple[float | None, float | None]:
    normalized = " ".join(text.split())
    k_point = re.search(r"\bK-POINT\s*(\d+(?:\.\d+)?)\s*m\b", normalized, re.I)
    meter = re.search(r"\bMETER VALUE\s*(\d+(?:\.\d+)?)\b", normalized, re.I)
    return (
        float(k_point.group(1)) if k_point else None,
        float(meter.group(1)) if meter else None,
    )


def _check_ranks(rows: list[dict[str, Any]], value_key: str, rank_key: str, errors: list[str]) -> None:
    ranked = [row for row in rows if row.get(value_key) is not None and row.get(rank_key) is not None]
    for row in ranked:
        expected = 1 + sum(1 for other in ranked if other[value_key] > row[value_key] + 1e-9)
        if row[rank_key] != expected:
            errors.append(
                f"{row.get('Name')}: {rank_key} {row[rank_key]} expected {expected}"
            )


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
            for number in range(1, ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                distance_points = row.get(f"Distance Points {number}")
                if distance is None or distance_points is None or k_point is None:
                    continue
                calculated = base + (distance - k_point) * meter_value
                if abs(calculated - distance_points) > TOLERANCE:
                    errors.append(f"{row['Name']}: distance points round {number}")
        for row in rows:
            round_sum = 0.0
            rounds = 0
            for number in range(1, ROUNDS + 1):
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
                    row.get(f"Wind Points {number}"),
                ]
                if any(part is None for part in parts):
                    errors.append(f"{row['Name']}: incomplete round {number}")
                    continue
                round_total = row.get(f"Round Total {number}")
                calculated = sum(parts)
                if round_total is not None and abs(calculated - round_total) > TOLERANCE:
                    errors.append(f"{row['Name']}: round {number} total {calculated:.1f} != {round_total}")
                round_sum += round_total if round_total is not None else calculated
            if row.get("Total") is not None and abs(round_sum - row["Total"]) > TOLERANCE * max(rounds, 1):
                errors.append(f"{row['Name']}: total {round_sum:.1f} != {row['Total']}")
            if row.get("Total") is None and row.get("Status") not in STATUS_CODES:
                errors.append(f"{row['Name']}: missing total")
        _check_ranks(rows, "Total", "Rank", errors)
        for number in range(1, ROUNDS + 1):
            _check_ranks(rows, f"Round Total {number}", f"Round Rank {number}", errors)
    if "training" in state.layouts:
        for row in rows:
            for number in range(1, ROUNDS + 1):
                distance = row.get(f"Distance {number}")
                points = row.get(f"Round Total {number}")
                if distance is None or points is None or k_point is None:
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
        for number in range(1, ROUNDS + 1):
            _check_ranks(rows, f"Speed {number}", f"Speed Rank {number}", errors)
            _check_ranks(rows, f"Distance {number}", f"Distance Rank {number}", errors)
            _check_ranks(rows, f"Round Total {number}", f"Round Rank {number}", errors)
    bibs = [row["Bib"] for row in rows + state.section_rows if row.get("Bib") is not None]
    duplicates = sorted({bib for bib in bibs if bibs.count(bib) > 1})
    if duplicates:
        errors.append(f"duplicate bibs {duplicates}")
    if errors:
        raise ResultsParseError(
            f"{state.pdf_name}: result check failed: " + "; ".join(errors[:5])
            + (f" (+{len(errors) - 5} more)" if len(errors) > 5 else "")
        )


def extract_results(pdf_path: Path) -> list[dict[str, Any]]:
    """Return one dict per athlete (keys from RESULT_COLUMNS) for a PDF."""
    state = _State(pdf_path.name)
    try:
        with pymupdf.open(pdf_path) as document:
            text = "\n".join(page.get_text() for page in document)
            for page_number, page in enumerate(document, start=1):
                _parse_page(page, page_number, state)
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as error:
        raise ResultsParseError(f"{pdf_path.name}: cannot read PDF: {error}") from error
    _validate(state, text)
    return state.rows + state.section_rows


HEADER_FILL = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
ERROR_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")


def write_results_sheet(
    workbook: Workbook, sheet_name: str, rows: list[dict[str, Any]]
) -> None:
    """Add a sheet with one row per athlete and the fixed RESULT_COLUMNS."""
    worksheet = workbook.create_sheet(sheet_name[:31])
    for column, key in enumerate(RESULT_COLUMNS, start=1):
        cell = worksheet.cell(row=1, column=column, value=key)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        letter = cell.column_letter
        worksheet.column_dimensions[letter].width = {
            "Name": 26, "Club": 30, "Remarks": 50, "Date of Birth": 12, "Status": 13,
        }.get(key, 9)
    for row_number, values in enumerate(rows, start=2):
        for column, key in enumerate(RESULT_COLUMNS, start=1):
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
