"""Post-process the exported ``fol_calc`` worksheet before it is saved."""

from __future__ import annotations

import math
import re
from copy import copy
from datetime import date, datetime

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from src.utils.fis_pdf import build_fis_pdf_url, fis_season


# Workbook colors are kept in one place so they can be adjusted easily.
COLOR_CALCULATED = "FFF2CC"
COLOR_FORMULA_TEXT = "00B050"
COLOR_DG_EVEN = "D9D9D9"
COLOR_DG_ODD = "BFBFBF"
COLOR_WARNING = "F4CCCC"
COLOR_SUCCESS = "D9EAD3"
COLOR_SUI_TEXT = "0000FF"
COLOR_HYPERLINK_TEXT = "0563C1"

CODEX_DOCUMENT_TYPES = {
    "T": "TNG",
    "Q": "QUAL",
    "P": "TRIAL",
    "W": "RL",
}

# Calculation behavior and comparison precision.
ALLOW_WPTS_EXTRAPOLATION = False
CALCULATION_DECIMAL_PLACES = 1
FLOAT_TOLERANCE = 1e-6

ONE_DECIMAL_HEADERS = {
    "km/h",
    "Dist_m",
    "Dpts",
    "A",
    "B",
    "C",
    "D",
    "E",
    "Jpts",
    "Gpts",
    "Wpts",
    "Rpts",
    "Total",
    "To_beat",
    "Dist_oben",
}
BOLD_COLUMN_HEADERS = {"BIB", "Dpts", "Gpts", "Wpts", "Jpts", "Rpts"}
JUDGE_HEADERS = ("A", "B", "C", "D", "E", "Jpts")
PRIMARY_MISSING_WARNING_HEADERS = (
    "BIB",
    "Name",
    "Vorname",
    "NSA",
    "km/h",
    "Dist_m",
    "Dpts",
)
MISSING_WARNING_HEADERS = (
    "Gate",
    "Gpts",
    "Wind",
    "Wpts",
    "Rpts",
    "Total",
    "Dist_oben",
)


def _opaque_color(rgb: str) -> str:
    """Convert configured RRGGBB colors to Excel's opaque ARGB form."""
    rgb = rgb.lstrip("#").upper()
    return rgb if len(rgb) == 8 else f"FF{rgb}"


def _solid_fill(rgb: str) -> PatternFill:
    """Create a solid fill that works for both cells and differential styles."""
    color = _opaque_color(rgb)
    return PatternFill(
        fill_type="solid",
        start_color=color,
        end_color=color,
    )


_CALCULATED_FILL = _solid_fill(COLOR_CALCULATED)
_DG_EVEN_FILL = _solid_fill(COLOR_DG_EVEN)
_DG_ODD_FILL = _solid_fill(COLOR_DG_ODD)
_WARNING_FILL = _solid_fill(COLOR_WARNING)
_SUI_FONT = Font(color=_opaque_color(COLOR_SUI_TEXT))
_FORMULA_FONT = Font(color=_opaque_color(COLOR_FORMULA_TEXT))


def _normalize_header(value) -> str:
    """Normalize a header for case-insensitive, punctuation-tolerant matching."""
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _header_columns(worksheet) -> dict[str, int]:
    """Return normalized header names mapped to their one-based column indexes."""
    columns = {}
    for cell in worksheet[1]:
        normalized = _normalize_header(cell.value)
        if normalized and normalized not in columns:
            columns[normalized] = cell.column
    return columns


def _column(columns: dict[str, int], header: str, *aliases: str) -> int | None:
    """Find a column by its header name and optional legacy aliases."""
    for candidate in (header, *aliases):
        index = columns.get(_normalize_header(candidate))
        if index is not None:
            return index
    return None


def _is_missing(value, *, zero_is_missing: bool = False) -> bool:
    """Treat None, empty strings, and optionally numeric zero as missing."""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return zero_is_missing and _to_float(value) == 0


def _to_float(value) -> float | None:
    """Return a finite numeric value, accepting decimal commas in text."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", ".")
        if not value or value.startswith("="):
            return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.lstrip().startswith("=")


def _cell_reference(row: int, column: int, *, absolute: bool = False) -> str:
    letter = get_column_letter(column)
    return f"${letter}${row}" if absolute else f"{letter}{row}"


def _resolve_simple_reference(worksheet, row: int, column: int, seen=None):
    """Resolve direct-reference formulas such as ``=W2`` recursively."""
    seen = set() if seen is None else seen
    key = (row, column)
    if key in seen:
        return None
    seen.add(key)

    value = worksheet.cell(row, column).value
    if not _is_formula(value):
        return value

    match = re.fullmatch(
        r"=\s*\$?([A-Z]{1,3})\$?(\d+)\s*",
        value.strip(),
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    referenced_column = 0
    for character in match.group(1).upper():
        referenced_column = referenced_column * 26 + ord(character) - 64
    return _resolve_simple_reference(
        worksheet,
        int(match.group(2)),
        referenced_column,
        seen,
    )


def _effective_number(worksheet, row: int, column: int) -> float | None:
    return _to_float(_resolve_simple_reference(worksheet, row, column))


def _parse_event_date(value) -> date | None:
    """Parse the date representations used in the imported FIS data."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None

    text = str(value).strip()
    if not text or text.startswith("="):
        return None
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass
    for date_format in ("%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            pass
    return None


def _normalize_codex(value) -> str | None:
    """Normalize numeric Codex values without changing digit-only text."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        return str(int(value)) if value.is_integer() else str(value)

    text = str(value).strip()
    if not text or text.startswith("="):
        return None
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _mark_calculated(cell) -> None:
    """Apply the configured fill to a value produced by the new logic."""
    cell.fill = copy(_CALCULATED_FILL)


def _fill_inherited_columns(worksheet, columns: dict[str, int]) -> None:
    """Fill blank DG and Typ cells with a reference to the preceding row."""
    if worksheet.max_row < 3:
        return
    for header in ("DG", "Typ"):
        column = _column(columns, header)
        if column is None:
            continue
        for row in range(3, worksheet.max_row + 1):
            cell = worksheet.cell(row, column)
            if _is_missing(cell.value):
                cell.value = f"={_cell_reference(row - 1, column)}"


def _gate_key(value):
    numeric = _to_float(value)
    if numeric is not None:
        return ("number", numeric)
    if _is_missing(value):
        return None
    return ("text", str(value).strip().casefold())


def _explicit_dg_starts(worksheet, dg_column: int) -> list[int]:
    """Find block starts before inherited DG formulas are considered."""
    return [
        row
        for row in range(2, worksheet.max_row + 1)
        if not _is_missing(worksheet.cell(row, dg_column).value)
        and not _is_formula(worksheet.cell(row, dg_column).value)
    ]


def _fill_gate_points(
    worksheet,
    columns: dict[str, int],
    explicit_dg_starts: list[int],
) -> None:
    """Copy Gpts within each explicit DG block and seed an unresolved first gate."""
    gate_column = _column(columns, "Gate")
    gpts_column = _column(columns, "Gpts")
    if gate_column is None or gpts_column is None:
        return

    for block_index, start_row in enumerate(explicit_dg_starts):
        end_row = (
            explicit_dg_starts[block_index + 1] - 1
            if block_index + 1 < len(explicit_dg_starts)
            else worksheet.max_row
        )

        known_by_gate = {}
        for row in range(start_row, end_row + 1):
            gate = _gate_key(worksheet.cell(row, gate_column).value)
            gpts = _effective_number(worksheet, row, gpts_column)
            if gate is not None and gpts is not None and gate not in known_by_gate:
                known_by_gate[gate] = gpts

        for row in range(start_row, end_row + 1):
            target = worksheet.cell(row, gpts_column)
            gate = _gate_key(worksheet.cell(row, gate_column).value)
            if _is_missing(target.value) and gate in known_by_gate:
                target.value = known_by_gate[gate]
                _mark_calculated(target)

        first_target = worksheet.cell(start_row, gpts_column)
        if _is_missing(first_target.value):
            first_gate = _gate_key(worksheet.cell(start_row, gate_column).value)
            first_target.value = 0
            _mark_calculated(first_target)
            if first_gate is not None:
                for row in range(start_row + 1, end_row + 1):
                    target = worksheet.cell(row, gpts_column)
                    gate = _gate_key(worksheet.cell(row, gate_column).value)
                    if gate == first_gate and _is_missing(target.value):
                        target.value = 0
                        _mark_calculated(target)


def _fill_distance_references(worksheet, columns: dict[str, int]) -> None:
    """Link a missing Dist_m cell to Dist_oben in the same row."""
    distance_column = _column(columns, "Dist_m")
    upper_distance_column = _column(columns, "Dist_oben")
    if distance_column is None or upper_distance_column is None:
        return
    for row in range(2, worksheet.max_row + 1):
        target = worksheet.cell(row, distance_column)
        source = worksheet.cell(row, upper_distance_column)
        if _is_missing(target.value) and not _is_missing(source.value):
            target.value = f"={_cell_reference(row, upper_distance_column)}"
            _mark_calculated(target)


def _find_hill_parameters(schanzen_worksheet):
    """Return K and Pts./m values plus their cells for either common layout."""
    if schanzen_worksheet is None:
        return None

    # The application currently writes SchanzenID transposed: Header | Value.
    transposed = {}
    for row in range(1, schanzen_worksheet.max_row + 1):
        key = _normalize_header(schanzen_worksheet.cell(row, 1).value)
        if key:
            transposed[key] = schanzen_worksheet.cell(row, 2)
    k_cell = transposed.get(_normalize_header("K"))
    points_cell = transposed.get(_normalize_header("Pts./m"))

    # Also support a conventional header row with the record beneath it.
    if k_cell is None or points_cell is None:
        headers = _header_columns(schanzen_worksheet)
        k_column = _column(headers, "K")
        points_column = _column(headers, "Pts./m", "Pts/m", "Pts_m")
        if k_column is not None and points_column is not None and schanzen_worksheet.max_row >= 2:
            k_cell = schanzen_worksheet.cell(2, k_column)
            points_cell = schanzen_worksheet.cell(2, points_column)

    k_value = _to_float(k_cell.value) if k_cell is not None else None
    points_value = _to_float(points_cell.value) if points_cell is not None else None
    if k_value is None or points_value is None:
        return None
    return k_value, points_value, k_cell, points_cell


def _fill_distance_points(
    worksheet,
    columns: dict[str, int],
    schanzen_worksheet,
) -> None:
    """Copy Dpts by effective distance or calculate it from SchanzenID."""
    distance_column = _column(columns, "Dist_m")
    dpts_column = _column(columns, "Dpts")
    if distance_column is None or dpts_column is None:
        return

    known_by_distance = {}
    for row in range(2, worksheet.max_row + 1):
        distance = _effective_number(worksheet, row, distance_column)
        dpts = _effective_number(worksheet, row, dpts_column)
        if distance is not None and dpts is not None and distance not in known_by_distance:
            known_by_distance[distance] = dpts

    hill_parameters = _find_hill_parameters(schanzen_worksheet)
    for row in range(2, worksheet.max_row + 1):
        target = worksheet.cell(row, dpts_column)
        if not _is_missing(target.value):
            continue
        distance = _effective_number(worksheet, row, distance_column)
        if distance is not None and distance in known_by_distance:
            target.value = known_by_distance[distance]
            _mark_calculated(target)
            continue
        if hill_parameters is None:
            continue

        k_value, points_value, k_cell, points_cell = hill_parameters
        base_value = 60 if k_value <= 160 else 120
        distance_ref = _cell_reference(row, distance_column)
        k_ref = _cell_reference(k_cell.row, k_cell.column, absolute=True)
        points_ref = _cell_reference(points_cell.row, points_cell.column, absolute=True)
        target.value = (
            f'=IF({distance_ref}="","",ROUND('
            f"IF('SchanzenID'!{k_ref}<=160,60,120)+({distance_ref}-"
            f"'SchanzenID'!{k_ref})*'SchanzenID'!{points_ref},"
            f"{CALCULATION_DECIMAL_PLACES}))"
        )
        _mark_calculated(target)
        if distance is not None:
            known_by_distance[distance] = round(
                base_value + (distance - k_value) * points_value,
                CALCULATION_DECIMAL_PLACES,
            )


def _judge_argument(judge_columns: list[int], row: int) -> str:
    """Use a compact range when the judge columns are contiguous."""
    if judge_columns == list(range(min(judge_columns), max(judge_columns) + 1)):
        return (
            f"{_cell_reference(row, min(judge_columns))}:"
            f"{_cell_reference(row, max(judge_columns))}"
        )
    return ",".join(_cell_reference(row, column) for column in judge_columns)


def _fill_judges_points(worksheet, columns: dict[str, int]) -> None:
    """Insert the middle-three judge formula using xlsx-invariant functions."""
    judge_columns = [_column(columns, header) for header in ("A", "B", "C", "D", "E")]
    jpts_column = _column(columns, "Jpts")
    if jpts_column is None or any(column is None for column in judge_columns):
        return
    judge_columns = [int(column) for column in judge_columns]
    e_column = _column(columns, "E")

    for row in range(2, worksheet.max_row + 1):
        target = worksheet.cell(row, jpts_column)
        e_value = _to_float(worksheet.cell(row, e_column).value)
        if _is_missing(target.value) and e_value is not None and e_value > 5:
            arguments = _judge_argument(judge_columns, row)
            target.value = (
                f"=SUM({arguments})-MIN({arguments})-MAX({arguments})"
            )
            _mark_calculated(target)


def _linear_value(x, lower, upper) -> float | None:
    if lower is None or upper is None or lower[0] == upper[0]:
        return None
    return lower[1] + (x - lower[0]) * (upper[1] - lower[1]) / (upper[0] - lower[0])


def _fill_wind_points(worksheet, columns: dict[str, int]) -> None:
    """Copy exact Wpts matches or interpolate between surrounding references."""
    wind_column = _column(columns, "Wind")
    wpts_column = _column(columns, "Wpts")
    if wind_column is None or wpts_column is None:
        return

    references = {}
    for row in range(2, worksheet.max_row + 1):
        wind = _effective_number(worksheet, row, wind_column)
        wpts = _effective_number(worksheet, row, wpts_column)
        if wind is not None and wpts is not None and wind not in references:
            references[wind] = wpts
    reference_points = sorted(references.items())

    for row in range(2, worksheet.max_row + 1):
        target = worksheet.cell(row, wpts_column)
        wind = _effective_number(worksheet, row, wind_column)
        if not _is_missing(target.value) or wind is None:
            continue

        result = references.get(wind)
        if result is None:
            lower = next((point for point in reversed(reference_points) if point[0] < wind), None)
            upper = next((point for point in reference_points if point[0] > wind), None)
            result = _linear_value(wind, lower, upper)

            if result is None and ALLOW_WPTS_EXTRAPOLATION and len(reference_points) >= 2:
                pair = reference_points[:2] if lower is None else reference_points[-2:]
                result = _linear_value(wind, pair[0], pair[1])

        if result is not None:
            target.value = round(result, CALCULATION_DECIMAL_PLACES)
            _mark_calculated(target)


def _fill_round_points(worksheet, columns: dict[str, int]) -> None:
    """Insert the sum of Dpts, Gpts, Wpts, and Jpts into missing/zero Rpts."""
    component_columns = [
        _column(columns, header) for header in ("Dpts", "Gpts", "Wpts", "Jpts")
    ]
    rpts_column = _column(columns, "Rpts")
    if rpts_column is None or any(column is None for column in component_columns):
        return
    for row in range(2, worksheet.max_row + 1):
        target = worksheet.cell(row, rpts_column)
        if _is_missing(target.value, zero_is_missing=True):
            references = [
                _cell_reference(row, int(column)) for column in component_columns
            ]
            target.value = "=" + "+".join(references)
            _mark_calculated(target)


def _fill_total_from_round_points(
    worksheet,
    columns: dict[str, int],
    explicit_dg_starts: list[int],
) -> None:
    """Fill missing Total from current and, for a second W round, prior Rpts."""
    total_column = _column(columns, "Total")
    rpts_column = _column(columns, "Rpts")
    typ_column = _column(columns, "Typ")
    bib_column = _column(columns, "BIB")
    if total_column is None or rpts_column is None:
        return

    for block_index, start_row in enumerate(explicit_dg_starts):
        end_row = (
            explicit_dg_starts[block_index + 1] - 1
            if block_index + 1 < len(explicit_dg_starts)
            else worksheet.max_row
        )

        comparable_pairs = []
        for row in range(start_row, end_row + 1):
            total = _effective_number(worksheet, row, total_column)
            rpts = _effective_number(worksheet, row, rpts_column)
            if total is not None and rpts is not None:
                comparable_pairs.append((total, rpts))

        differing_pairs = sum(
            abs(total - rpts) > FLOAT_TOLERANCE
            for total, rpts in comparable_pairs
        )
        mostly_different = differing_pairs > len(comparable_pairs) / 2

        current_typ = (
            _resolve_simple_reference(worksheet, start_row, typ_column)
            if typ_column is not None
            else None
        )
        previous_rows_by_bib = {}
        is_second_w_round = False
        if block_index > 0 and typ_column is not None and bib_column is not None:
            previous_start = explicit_dg_starts[block_index - 1]
            previous_typ = _resolve_simple_reference(
                worksheet, previous_start, typ_column
            )
            is_second_w_round = (
                str(current_typ).strip().upper() == "W"
                and str(previous_typ).strip().upper() == "W"
                and mostly_different
            )
            if is_second_w_round:
                previous_end = start_row - 1
                for previous_row in range(previous_start, previous_end + 1):
                    bib = _gate_key(worksheet.cell(previous_row, bib_column).value)
                    if bib is not None and bib not in previous_rows_by_bib:
                        previous_rows_by_bib[bib] = previous_row

        for row in range(start_row, end_row + 1):
            target = worksheet.cell(row, total_column)
            if _is_missing(target.value):
                current_rpts_ref = _cell_reference(row, rpts_column)
                previous_row = None
                if is_second_w_round:
                    bib = _gate_key(worksheet.cell(row, bib_column).value)
                    previous_row = previous_rows_by_bib.get(bib)

                if previous_row is None:
                    target.value = f"={current_rpts_ref}"
                else:
                    previous_rpts_ref = _cell_reference(previous_row, rpts_column)
                    target.value = f"={current_rpts_ref}+{previous_rpts_ref}"
                _mark_calculated(target)


def apply_calculations(worksheet, schanzen_worksheet=None) -> None:
    """Run all calculation passes in dependency order."""
    columns = _header_columns(worksheet)
    dg_column = _column(columns, "DG")
    explicit_dg_starts = (
        _explicit_dg_starts(worksheet, dg_column) if dg_column is not None else []
    )

    # Preserve explicit DG boundaries before blank cells become formulas.
    _fill_inherited_columns(worksheet, columns)
    _fill_gate_points(worksheet, columns, explicit_dg_starts)
    _fill_distance_references(worksheet, columns)
    _fill_distance_points(worksheet, columns, schanzen_worksheet)
    _fill_judges_points(worksheet, columns)
    _fill_wind_points(worksheet, columns)
    _fill_round_points(worksheet, columns)
    _fill_total_from_round_points(
        worksheet,
        columns,
        explicit_dg_starts,
    )


def apply_standard_formatting(worksheet) -> None:
    """Apply requested alignment and number formats to fol_calc data rows."""
    if worksheet.max_row < 2:
        return
    columns = _header_columns(worksheet)
    left_aligned_columns = {
        column
        for column in (
            _column(columns, "Name"),
            _column(columns, "Vorname", "Vormane"),
            _column(columns, "MSG"),
            _column(columns, "Jumper_ID", "SpringerID"),
            _column(columns, "Jump_ID"),
        )
        if column is not None
    }

    for column in range(1, worksheet.max_column + 1):
        header_cell = worksheet.cell(1, column)
        header_cell.fill = copy(_CALCULATED_FILL)
        header_font = copy(header_cell.font)
        header_font.bold = True
        header_cell.font = header_font
        header_alignment = copy(header_cell.alignment)
        header_alignment.horizontal = "center"
        header_cell.alignment = header_alignment

        horizontal = "left" if column in left_aligned_columns else "center"
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row, column)
            alignment = copy(cell.alignment)
            alignment.horizontal = horizontal
            cell.alignment = alignment

    for header in ONE_DECIMAL_HEADERS:
        column = _column(columns, header)
        if column is not None:
            for row in range(2, worksheet.max_row + 1):
                worksheet.cell(row, column).number_format = "0.0"

    wind_column = _column(columns, "Wind")
    if wind_column is not None:
        for row in range(2, worksheet.max_row + 1):
            worksheet.cell(row, wind_column).number_format = "0.00"

    for header in BOLD_COLUMN_HEADERS:
        column = _column(columns, header)
        if column is None:
            continue
        for row in range(1, worksheet.max_row + 1):
            cell = worksheet.cell(row, column)
            font = copy(cell.font)
            font.bold = True
            cell.font = font


def apply_codex_hyperlinks(worksheet) -> None:
    """Attach each FIS PDF link to the corresponding Typ cell."""
    if worksheet.max_row < 2:
        return
    columns = _header_columns(worksheet)
    codex_column = _column(columns, "Codex")
    date_column = _column(columns, "Date")
    typ_column = _column(columns, "Typ")
    if codex_column is None or date_column is None or typ_column is None:
        return

    for row in range(2, worksheet.max_row + 1):
        typ_cell = worksheet.cell(row, typ_column)
        codex = _normalize_codex(
            _resolve_simple_reference(worksheet, row, codex_column)
        )
        event_date = _parse_event_date(
            _resolve_simple_reference(worksheet, row, date_column)
        )
        typ_value = _resolve_simple_reference(worksheet, row, typ_column)
        typ = str(typ_value).strip().upper() if typ_value is not None else ""
        document_type = CODEX_DOCUMENT_TYPES.get(typ)
        if codex is None or event_date is None or document_type is None:
            continue

        typ_cell.hyperlink = build_fis_pdf_url(
            fis_season(event_date), codex, document_type
        )
        font = copy(typ_cell.font)
        font.color = _opaque_color(COLOR_HYPERLINK_TEXT)
        font.underline = "single"
        typ_cell.font = font


def collect_codex_download_keys(worksheet) -> set[tuple[int, str]]:
    """Collect each unique (season, Codex) pair from the exported rows."""
    columns = _header_columns(worksheet)
    codex_column = _column(columns, "Codex")
    date_column = _column(columns, "Date")
    if codex_column is None or date_column is None:
        return set()

    keys: set[tuple[int, str]] = set()
    for row in range(2, worksheet.max_row + 1):
        codex = _normalize_codex(
            _resolve_simple_reference(worksheet, row, codex_column)
        )
        event_date = _parse_event_date(
            _resolve_simple_reference(worksheet, row, date_column)
        )
        if codex is not None and event_date is not None:
            keys.add((fis_season(event_date), codex))
    return keys

def _column_range(column: int, last_row: int) -> str:
    letter = get_column_letter(column)
    return f"{letter}2:{letter}{last_row}"


def _add_formula_rule(worksheet, cell_range: str, formula: str, *, fill=None, font=None):
    worksheet.conditional_formatting.add(
        cell_range,
        FormulaRule(formula=[formula], fill=fill, font=font),
    )


def apply_conditional_formatting(worksheet) -> None:
    """Apply the currently approved conditional-formatting rules."""
    if worksheet.max_row < 2:
        return

    columns = _header_columns(worksheet)
    last_row = worksheet.max_row
    dg_column = _column(columns, "DG")
    typ_column = _column(columns, "Typ")
    bib_column = _column(columns, "BIB")

    # ISFORMULA was introduced after the original OOXML function set, so the
    # invariant xlfn name is required in the file. German Excel displays this
    # as ISTFORMEL while evaluating each relative cell reference.
    first_cell = "G2"
    last_cell = f"{get_column_letter(worksheet.max_column)}{last_row}"
    _add_formula_rule(
        worksheet,
        f"{first_cell}:{last_cell}",
        f"_xlfn.ISFORMULA({first_cell})",
        font=_FORMULA_FONT,
    )

    # Shade DG and Typ together according to the effective DG value.
    if dg_column is not None:
        dg_ref = f"${get_column_letter(dg_column)}2"
        even_formula = f'AND({dg_ref}<>"",MOD({dg_ref},2)=0)'
        odd_formula = f'AND({dg_ref}<>"",MOD({dg_ref},2)=1)'
        for target_column in (dg_column, typ_column):
            if target_column is None:
                continue
            target_range = _column_range(target_column, last_row)
            _add_formula_rule(
                worksheet,
                target_range,
                even_formula,
                fill=_DG_EVEN_FILL,
            )
            _add_formula_rule(
                worksheet,
                target_range,
                odd_formula,
                fill=_DG_ODD_FILL,
            )

    # A BIB may occur only once within the same effective DG. COUNTIFS is
    # stored with Excel's invariant xlfn name so localized (including German)
    # Excel versions can evaluate the rule. Blank DG/BIB cells are ignored.
    if dg_column is not None and bib_column is not None:
        dg_letter = get_column_letter(dg_column)
        bib_letter = get_column_letter(bib_column)
        dg_ref = f"${dg_letter}2"
        bib_ref = f"${bib_letter}2"
        dg_range = f"${dg_letter}$2:${dg_letter}${last_row}"
        bib_range = f"${bib_letter}$2:${bib_letter}${last_row}"
        _add_formula_rule(
            worksheet,
            _column_range(bib_column, last_row),
            (
                f'AND({dg_ref}<>"",{bib_ref}<>"",'
                f"_xlfn.COUNTIFS({dg_range},{dg_ref},"
                f"{bib_range},{bib_ref})>1)"
            ),
            fill=_WARNING_FILL,
        )

    # Display Swiss athlete identity fields in red.
    nsa_column = _column(columns, "NSA")
    if nsa_column is not None:
        nsa_ref = f"${get_column_letter(nsa_column)}2"
        sui_formula = f'UPPER(TRIM({nsa_ref}&""))="SUI"'
        for header, aliases in (
            ("Name", ()),
            ("Vorname", ("Vormane",)),
            ("NSA", ()),
        ):
            target_column = _column(columns, header, *aliases)
            if target_column is not None:
                _add_formula_rule(
                    worksheet,
                    _column_range(target_column, last_row),
                    sui_formula,
                    font=_SUI_FONT,
                )

    # Validate judge values individually according to the effective Typ.
    if typ_column is not None:
        typ_ref = f"${get_column_letter(typ_column)}2"
        training_type = (
            f'OR(UPPER(TRIM({typ_ref}&""))="T",'
            f'UPPER(TRIM({typ_ref}&""))="P")'
        )
        scored_type = (
            f'OR(UPPER(TRIM({typ_ref}&""))="Q",'
            f'UPPER(TRIM({typ_ref}&""))="W")'
        )
        for header in JUDGE_HEADERS:
            target_column = _column(columns, header)
            if target_column is None:
                continue
            cell_ref = f"{get_column_letter(target_column)}2"
            target_range = _column_range(target_column, last_row)
            _add_formula_rule(
                worksheet,
                target_range,
                (
                    f"AND({training_type},ISNUMBER({cell_ref}),"
                    f"{cell_ref}>0)"
                ),
                fill=_WARNING_FILL,
            )
            _add_formula_rule(
                worksheet,
                target_range,
                (
                    f"AND({scored_type},"
                    f'OR({cell_ref}="",'
                    f"AND(ISNUMBER({cell_ref}),{cell_ref}<3)))"
                ),
                fill=_WARNING_FILL,
            )

    # Highlight the primary required fields with one relative rule when they
    # form the standard contiguous BIB:Dpts block. This produces C:I in the
    # exported workbook while remaining header-driven for alternate layouts.
    primary_columns = [
        _column(columns, header)
        for header in PRIMARY_MISSING_WARNING_HEADERS
    ]
    available_primary_columns = [
        column for column in primary_columns if column is not None
    ]
    if (
        len(available_primary_columns) == len(PRIMARY_MISSING_WARNING_HEADERS)
        and available_primary_columns
        == list(range(
            available_primary_columns[0],
            available_primary_columns[-1] + 1,
        ))
    ):
        first_primary_column = available_primary_columns[0]
        first_primary_letter = get_column_letter(first_primary_column)
        last_primary_letter = get_column_letter(available_primary_columns[-1])
        _add_formula_rule(
            worksheet,
            f"{first_primary_letter}2:{last_primary_letter}{last_row}",
            f'{first_primary_letter}2=""',
            fill=_WARNING_FILL,
        )
    else:
        for target_column in available_primary_columns:
            cell_ref = f"{get_column_letter(target_column)}2"
            _add_formula_rule(
                worksheet,
                _column_range(target_column, last_row),
                f'{cell_ref}=""',
                fill=_WARNING_FILL,
            )

    # A valid distance must be numeric and use whole or half metres. Blank
    # values are handled by the required-field rule above.
    distance_column = _column(columns, "Dist_m")
    if distance_column is not None:
        distance_ref = f"{get_column_letter(distance_column)}2"
        _add_formula_rule(
            worksheet,
            _column_range(distance_column, last_row),
            (
                f'AND({distance_ref}<>"",'
                f"OR(NOT(ISNUMBER({distance_ref})),"
                f"MOD({distance_ref}*2,1)<>0))"
            ),
            fill=_WARNING_FILL,
        )

    # Highlight each remaining missing required value; numeric zero is valid.
    for header in MISSING_WARNING_HEADERS:
        target_column = _column(columns, header)
        if target_column is None:
            continue
        cell_ref = f"{get_column_letter(target_column)}2"
        _add_formula_rule(
            worksheet,
            _column_range(target_column, last_row),
            f'{cell_ref}=""',
            fill=_WARNING_FILL,
        )

    # Warn when the generated Jumper_ID is empty or absent from column A of
    # the SpringerID reference sheet. INDIRECT keeps the cross-sheet reference
    # compatible with Excel conditional-format formulas.
    springer_id_column = _column(columns, "Jumper_ID", "SpringerID")
    if springer_id_column is not None:
        cell_ref = f"{get_column_letter(springer_id_column)}2"
        _add_formula_rule(
            worksheet,
            _column_range(springer_id_column, last_row),
            (
                f'OR({cell_ref}="",'
                f'COUNTIF(INDIRECT("\'SpringerID\'!A:A"),{cell_ref})=0)'
            ),
            fill=_WARNING_FILL,
        )


def process_fol_calc_worksheet(worksheet, schanzen_worksheet=None) -> None:
    """Apply calculations and the approved formatting to fol_calc."""
    apply_calculations(worksheet, schanzen_worksheet)
    apply_standard_formatting(worksheet)
    apply_codex_hyperlinks(worksheet)
    apply_conditional_formatting(worksheet)
