"""
Data imputation engine for fol_calcdata.

Runs sequential loops to fill in missing values:
  Loop 1: Jump-level single checks (cross-row lookups)
  Loop 2: DG-grouped checks (within same Durchgang)
  Loop 3: Multi-variable calculation (formula-based solve)
  Loop 4: Sync workspace_msg with workspace_fol

Column indices reference (calc_headers, 33 columns):
  0=DG, 1=Typ, 2=BIB, 3=Name, 4=Vorname, 5=NSA, 6=km/h,
  7=Dist_m, 8=Dpts, 9=A, 10=B, 11=C, 12=D, 13=E, 14=Jpts,
  15=Gate, 16=Gpts, 17=Wind, 18=Wpts, 19=Rpts, 20=Total,
  21=To_beat, 22=Dist_oben, 23=Rk_LDG, 24=NPS_DSQ,
  25=Gate_Captain, 26=MSG, 27=Codex, 28=Date, 29=Time,
  30=Gap, 31=Sheet, 32=Row
"""
import copy
import re
from collections import Counter, defaultdict
from datetime import date, datetime, time

from src.pipeline.phase_20_preprocessing.jump_id_matching import (
    assign_msg_block_indexes,
    build_jump_blocks,
    normalize_codex,
    parse_date_time,
)


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _is_empty(val) -> bool:
    """Return True if the value is None or a blank string."""
    return val is None or str(val).strip() == ""


def _to_float(val):
    """Convert a value to float, return None on failure."""
    if val is None:
        return None
    if isinstance(val, str):
        val = re.sub(r'^[🔵🔹🔷\s]+', '', val).strip()
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _norm_text(val) -> str:
    """Return a stripped string; blank for None."""
    if val is None:
        return ""
    s = str(val).strip()
    return re.sub(r'^[🔵🔹🔷\s]+', '', s).strip()


def _norm_upper(val) -> str:
    """Return a stripped uppercase string; blank for None."""
    return _norm_text(val).upper()


def _norm_bib(val) -> str:
    """Normalize BIB values so 12 and 12.0 compare equally."""
    num = _to_float(val)
    if num is not None and num.is_integer():
        return str(int(num))
    return _norm_text(val)


def _date_key(val) -> str:
    """Normalize common date values for equality checks."""
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date):
        return val.isoformat()
    text = _norm_text(val)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).date().isoformat()
        except ValueError:
            pass
    return text


def _time_minutes(val) -> int | None:
    """Return minutes since midnight for common time values."""
    if isinstance(val, datetime):
        return val.hour * 60 + val.minute
    if isinstance(val, time):
        return val.hour * 60 + val.minute

    text = _norm_text(val)
    if not text:
        return None

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            parsed = datetime.strptime(text[:8] if fmt == "%H:%M:%S" else text[:5], fmt)
            return parsed.hour * 60 + parsed.minute
        except ValueError:
            pass

    match = re.search(r"(\d{1,2})[:.](\d{2})", text)
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2))
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return hour * 60 + minute

    return None


def _append_msg(row: list, text: str) -> bool:
    """Append message text to fol MSG column; return True if changed."""
    msg = _norm_text(text)
    if not msg:
        return False

    current = _norm_text(row[26])
    if not current:
        row[26] = msg
        return True
    if msg in [part.strip() for part in current.split(";")]:
        return False

    row[26] = f"{current}; {msg}"
    return True


def _build_majority_map(data: list[list], src_col: int, target_col: int) -> dict:
    """
    Build {src_value: most_common_target_value} from rows where both columns
    are populated.  Uses majority vote; ties broken by first occurrence.
    """
    mapping: dict[float, list] = defaultdict(list)
    for row in data:
        src = _to_float(row[src_col])
        tgt = row[target_col]
        if src is not None and not _is_empty(tgt):
            mapping[src].append(tgt)

    result: dict[float, object] = {}
    for key, vals in mapping.items():
        counter = Counter(str(v) for v in vals)
        most_common_str = counter.most_common(1)[0][0]
        # Return the original (non-stringified) value
        for v in vals:
            if str(v) == most_common_str:
                result[key] = v
                break
    return result


# ---------------------------------------------------------------------------
# Pre-processing: resolve formula strings to numeric values
# ---------------------------------------------------------------------------

def _resolve_formulas(data: list[list]) -> None:
    """
    Resolve Jpts SUM formulas to their numeric equivalents, in-place.

    This mirrors the logic already executed by the JS editor (evalGrid),
    so resolved values will match what the user sees in the table.
    """
    # Resolve Jpts formulas (col 14)
    for r, row in enumerate(data):
        val_jpts = row[14]
        if isinstance(val_jpts, str) and val_jpts.startswith("="):
            judges = [_to_float(row[i]) for i in range(9, 14)]  # A–E
            valid = [v for v in judges if v is not None]
            if len(valid) == 5:
                row[14] = round(sum(valid) - min(valid) - max(valid), 1)
            else:
                row[14] = None


# ---------------------------------------------------------------------------
# Loop 1: Jump-level single checks
# ---------------------------------------------------------------------------

def _loop1_single_checks(data: list[list], changed: set[tuple[int, int]]) -> None:
    """
    Iterate every record and fill missing values via cross-row lookups.
    Uses pre-built dict maps for O(1) lookups.
    """
    # Build lookup maps once from current data
    dpts_to_dist = _build_majority_map(data, src_col=8, target_col=7)    # Dpts → Dist_m
    dist_to_dpts = _build_majority_map(data, src_col=7, target_col=8)    # Dist_m → Dpts
    wind_to_wpts = _build_majority_map(data, src_col=17, target_col=18)  # Wind → Wpts
    wpts_to_wind = _build_majority_map(data, src_col=18, target_col=17)  # Wpts → Wind

    for r, row in enumerate(data):
        # Missing Dist_m from Dpts
        if _is_empty(row[7]) and not _is_empty(row[8]):
            val = dpts_to_dist.get(_to_float(row[8]))
            if val is not None:
                row[7] = val
                changed.add((r, 7))

        # Missing Dist_m from Dist_oben (fallback)
        if _is_empty(row[7]) and not _is_empty(row[22]):
            row[7] = row[22]
            changed.add((r, 7))

        # Missing Dpts from Dist_m
        if _is_empty(row[8]) and not _is_empty(row[7]):
            val = dist_to_dpts.get(_to_float(row[7]))
            if val is not None:
                row[8] = val
                changed.add((r, 8))

        # Missing Wpts from Wind
        if _is_empty(row[18]) and not _is_empty(row[17]):
            val = wind_to_wpts.get(_to_float(row[17]))
            if val is not None:
                row[18] = val
                changed.add((r, 18))

        # Missing Wind from Wpts
        if _is_empty(row[17]) and not _is_empty(row[18]):
            val = wpts_to_wind.get(_to_float(row[18]))
            if val is not None:
                row[17] = val
                changed.add((r, 17))

        # Missing Jpts from Judges (A–E, columns 9–13)
        if _is_empty(row[14]) and not _is_empty(row[13]):  # col E present
            judges = [_to_float(row[i]) for i in range(9, 14)]
            valid = [v for v in judges if v is not None]
            if len(valid) == 5:
                row[14] = round(sum(valid) - min(valid) - max(valid), 1)
                changed.add((r, 14))


# ---------------------------------------------------------------------------
# Loop 2: DG-grouped checks
# ---------------------------------------------------------------------------

def _loop2_dg_grouped(data: list[list], changed: set[tuple[int, int]]) -> None:
    """
    Group rows by DG and fill missing values using within-group context.
    """
    # Group row indices by DG value
    dg_groups: dict[object, list[int]] = defaultdict(list)
    for r, row in enumerate(data):
        dg_val = row[0]
        if dg_val is not None:
            dg_groups[dg_val].append(r)

    for dg_val, row_indices in dg_groups.items():
        # --- Missing Gpts from Gate (within same DG) ---
        gate_to_gpts: dict[float, object] = {}
        for r in row_indices:
            gate = data[r][15]
            gpts = data[r][16]
            if not _is_empty(gate) and not _is_empty(gpts):
                key = _to_float(gate)
                if key is not None:
                    gate_to_gpts.setdefault(key, gpts)

        for r in row_indices:
            if _is_empty(data[r][16]) and not _is_empty(data[r][15]):
                key = _to_float(data[r][15])
                if key is not None and key in gate_to_gpts:
                    data[r][16] = gate_to_gpts[key]
                    changed.add((r, 16))

        # --- Missing Rpts from Total ---
        typ = data[row_indices[0]][1]  # Typ of this DG (from first row)
        typ_str = str(typ).strip().upper() if typ else ""

        if typ_str in ("T", "Q", "P"):
            # For T, Q, P: Total and Rpts are identical
            for r in row_indices:
                if _is_empty(data[r][19]) and not _is_empty(data[r][20]):
                    data[r][19] = data[r][20]
                    changed.add((r, 19))
        elif typ_str == "W":
            # For W: check if there is a consistent offset between Total and Rpts
            offset = None
            for r in row_indices:
                rpts = _to_float(data[r][19])
                total = _to_float(data[r][20])
                if rpts is not None and total is not None:
                    offset = round(total - rpts, 2)
                    break

            for r in row_indices:
                if _is_empty(data[r][19]) and not _is_empty(data[r][20]):
                    total = _to_float(data[r][20])
                    if total is not None:
                        if offset is not None:
                            data[r][19] = round(total - offset, 1)
                        else:
                            data[r][19] = total
                        changed.add((r, 19))


# ---------------------------------------------------------------------------
# Loop 3: Multi-variable calculation & dependency resolution
# ---------------------------------------------------------------------------

def _loop3_multi_variable(data: list[list], changed: set[tuple[int, int]]) -> None:
    """
    Calculate missing values from mathematical relationships between
    Rpts, Dpts, Gpts, Wpts, and optionally Jpts.
    Also performs final cross-checks for Dist_m and Wind.
    """
    # Rebuild cross-reference maps (data may have changed after loops 1 & 2)
    dpts_to_dist = _build_majority_map(data, src_col=8, target_col=7)
    wpts_to_wind = _build_majority_map(data, src_col=18, target_col=17)

    for r, row in enumerate(data):
        typ = str(row[1]).strip().upper() if row[1] else ""
        uses_jpts = typ in ("Q", "W")

        dpts = _to_float(row[8])
        gpts = _to_float(row[16])
        wpts = _to_float(row[18])
        jpts = _to_float(row[14])
        rpts = _to_float(row[19])

        # Build component dict — only include Jpts when relevant
        components = {"Dpts": dpts, "Gpts": gpts, "Wpts": wpts, "Rpts": rpts}
        if uses_jpts:
            components["Jpts"] = jpts

        filled = {k: v for k, v in components.items() if v is not None}
        missing = [k for k, v in components.items() if v is None]

        # Solve when exactly one component is missing
        if len(missing) == 1:
            m = missing[0]

            if m == "Rpts":
                val = filled.get("Dpts", 0) + filled.get("Gpts", 0) + filled.get("Wpts", 0)
                if uses_jpts:
                    val += filled.get("Jpts", 0)
                row[19] = round(val, 1)
                changed.add((r, 19))

            elif m == "Dpts":
                val = filled["Rpts"] - filled.get("Gpts", 0) - filled.get("Wpts", 0)
                if uses_jpts:
                    val -= filled.get("Jpts", 0)
                row[8] = round(val, 1)
                changed.add((r, 8))

            elif m == "Gpts":
                val = filled["Rpts"] - filled.get("Dpts", 0) - filled.get("Wpts", 0)
                if uses_jpts:
                    val -= filled.get("Jpts", 0)
                row[16] = round(val, 1)
                changed.add((r, 16))

            elif m == "Wpts":
                val = filled["Rpts"] - filled.get("Dpts", 0) - filled.get("Gpts", 0)
                if uses_jpts:
                    val -= filled.get("Jpts", 0)
                row[18] = round(val, 1)
                changed.add((r, 18))

            elif m == "Jpts":
                val = filled["Rpts"] - filled.get("Dpts", 0) - filled.get("Gpts", 0) - filled.get("Wpts", 0)
                row[14] = round(val, 1)
                changed.add((r, 14))

        # Cross-check: Dist_m from Dpts
        if _is_empty(row[7]) and not _is_empty(row[8]):
            val = dpts_to_dist.get(_to_float(row[8]))
            if val is not None:
                row[7] = val
                changed.add((r, 7))

        # Cross-check: Wind from Wpts
        if _is_empty(row[17]) and not _is_empty(row[18]):
            val = wpts_to_wind.get(_to_float(row[18]))
            if val is not None:
                row[17] = val
                changed.add((r, 17))


# ---------------------------------------------------------------------------
# Loop 4: sync workspace_msg with workspace_fol
# ---------------------------------------------------------------------------

_STATUS_VALUES = {"DSQ", "NPS", "DNS"}
_TRIALJUMPER_STATUS = "TRIALJUMPER"
_SKIPPED_MSG_WORDS = re.compile(
    r"\b(?:progress|start|starts|completed)\b", re.IGNORECASE
)


def _find_block_by_jump_id(data: list[list], jump_id: str) -> dict | None:
    """Rebuild blocks after insertions and return the requested Jump_ID."""
    return next(
        (
            block
            for block in build_jump_blocks(data)
            if block["jump_id"] == jump_id
        ),
        None,
    )


def _nearest_row_in_block(
    data: list[list], block: dict, msg_datetime: datetime, bib=None
) -> int | None:
    """Find the nearest row in one Jump_ID, optionally restricted by BIB."""
    wanted_bib = _norm_bib(bib)
    candidates = []
    fallback = None
    for position, row_index in enumerate(block["row_indexes"]):
        row = data[row_index]
        if wanted_bib and _norm_bib(row[2]) != wanted_bib:
            continue
        if fallback is None:
            fallback = row_index
        row_datetime = parse_date_time(row[28], row[29])
        if row_datetime is not None:
            candidates.append(
                (abs(row_datetime - msg_datetime), position, row_index)
            )

    if candidates:
        return min(candidates, key=lambda candidate: candidate[:2])[2]
    return fallback


def find_trialjumper_remark_rows(
    fol_calcdata: list[list],
    msg_cleandata: list[list] | None = None,
) -> set[int]:
    """Return fol row indexes that should receive a trialjumper remark."""
    data = fol_calcdata or []
    blocks = build_jump_blocks(data)
    block_indexes = assign_msg_block_indexes(msg_cleandata or [], blocks)
    target_rows: set[int] = set()

    for msg_row, block_index in zip(msg_cleandata or [], block_indexes):
        if (
            len(msg_row) < 8
            or block_index is None
            or _norm_upper(msg_row[4]) != _TRIALJUMPER_STATUS
        ):
            continue

        msg_datetime = parse_date_time(msg_row[0], msg_row[1])
        block = blocks[block_index]
        if msg_datetime is None:
            continue

        nearest_row = _nearest_row_in_block(data, block, msg_datetime)
        if nearest_row is None:
            continue

        candidates = []
        first_row = max(0, nearest_row - 4)
        last_row = min(len(data) - 1, nearest_row + 4)
        for row_index in range(first_row, last_row + 1):
            row_distance = abs(row_index - nearest_row)
            gap = _to_float(data[row_index][30])
            if gap is None or gap <= 100:
                continue

            row_datetime = parse_date_time(
                data[row_index][28], data[row_index][29]
            )
            time_distance = (
                abs(row_datetime - msg_datetime)
                if row_datetime is not None
                else datetime.max - msg_datetime
            )
            candidates.append((row_distance, time_distance, row_index))

        if candidates:
            target_rows.add(min(candidates)[2])

    return target_rows


def _normalized_name_text(value) -> str:
    """Normalize names and messages for whole-phrase matching."""
    return " ".join(re.sub(r"[^\w]+", " ", _norm_text(value).casefold()).split())


def _message_contains_name(message: str, name, vorname) -> bool:
    """Return whether a stored athlete name is present in a message."""
    message_text = f" {_normalized_name_text(message)} "
    normalized_name = _normalized_name_text(name)
    normalized_vorname = _normalized_name_text(vorname)
    candidates = [normalized_name]
    if normalized_vorname:
        candidates.extend(
            [
                normalized_vorname,
                f"{normalized_vorname} {normalized_name}",
                f"{normalized_name} {normalized_vorname}",
            ]
        )
    return any(candidate and f" {candidate} " in message_text for candidate in candidates)


def _find_identity_in_other_block(
    data: list[list],
    target_jump_id: str,
    bib,
    codex,
    message: str,
) -> tuple[object, object, object] | None:
    """Find a message-confirmed identity in another block of the same Codex."""
    wanted_bib = _norm_bib(bib)
    wanted_codex = normalize_codex(codex)
    if not wanted_bib or not wanted_codex:
        return None

    for block in build_jump_blocks(data):
        if block["jump_id"] == target_jump_id:
            continue
        if wanted_codex not in block["codexes"]:
            continue
        for row_index in block["row_indexes"]:
            row = data[row_index]
            if normalize_codex(row[27]) != wanted_codex:
                continue
            if _norm_bib(row[2]) != wanted_bib:
                continue
            if _is_empty(row[3]) or not _message_contains_name(
                message, row[3], row[4]
            ):
                continue
            return row[3], row[4], row[5]

    return None


def _shift_changed_after_insert(
    changed: set[tuple[int, int]], insert_idx: int
) -> set[tuple[int, int]]:
    """Move changed-cell row indices after an inserted fol row."""
    return {
        (row + 1, col) if row >= insert_idx else (row, col)
        for row, col in changed
    }


def _find_block_insert_index(
    data: list[list], block: dict, msg_datetime: datetime
) -> int:
    """Find the chronological insertion point inside one Jump_ID block."""
    for row_index in block["row_indexes"]:
        row_datetime = parse_date_time(data[row_index][28], data[row_index][29])
        if row_datetime is not None and row_datetime > msg_datetime:
            return row_index
    return block["row_indexes"][-1] + 1


def _insert_missing_starter_row(
    data: list[list],
    changed: set[tuple[int, int]],
    msg_row: list,
    block: dict,
) -> tuple[int, set[tuple[int, int]]]:
    """Insert a missing status BIB into its chronological Jump_ID position."""
    status = _norm_upper(msg_row[4])
    message = _norm_text(msg_row[2])
    msg_datetime = parse_date_time(msg_row[0], msg_row[1])
    identity = _find_identity_in_other_block(
        data,
        block["jump_id"],
        msg_row[5],
        msg_row[7],
        message,
    )

    new_row = [None] * 33
    new_row[2] = msg_row[5]   # BIB
    new_row[24] = status      # NPS_DSQ
    new_row[26] = message     # MSG
    new_row[27] = msg_row[7]  # Codex
    new_row[28] = msg_row[0]  # Date
    new_row[29] = msg_row[1]  # Time
    if identity is not None:
        new_row[3], new_row[4], new_row[5] = identity

    insert_idx = _find_block_insert_index(data, block, msg_datetime)
    first_idx = block["row_indexes"][0]
    moves_block_marker = insert_idx == first_idx
    if moves_block_marker:
        new_row[0] = data[first_idx][0]
        new_row[1] = data[first_idx][1]

    changed = _shift_changed_after_insert(changed, insert_idx)
    data.insert(insert_idx, new_row)
    if moves_block_marker:
        data[insert_idx + 1][0] = None
        data[insert_idx + 1][1] = None

    for col in (2, 24, 26, 27, 28, 29):
        changed.add((insert_idx, col))
    if identity is not None:
        for col in (3, 4, 5):
            changed.add((insert_idx, col))
    if moves_block_marker:
        for col in (0, 1):
            changed.add((insert_idx, col))
            changed.add((insert_idx + 1, col))

    return insert_idx, changed


def _loop4_sync_msg_with_fol(
    data: list[list],
    msg_cleandata: list[list],
    changed: set[tuple[int, int]],
) -> set[tuple[int, int]]:
    """
    Sync useful message rows into fol_calcdata in one pass.

    msg_cleandata columns:
      0=Date, 1=Time, 2=Message, 3=Jump, 4=Status,
      5=BIB, 6=Gate, 7=Codex, 8=Sheet, 9=Row
    """
    original_blocks = build_jump_blocks(data)
    block_indexes = assign_msg_block_indexes(msg_cleandata, original_blocks)
    jump_ids = [
        original_blocks[index]["jump_id"] if index is not None else None
        for index in block_indexes
    ]

    # Status messages run first because they may insert rows. Each message
    # keeps the Jump_ID assigned from the original, unmodified fol data.
    for msg_row, jump_id in zip(msg_cleandata, jump_ids):
        if len(msg_row) < 8 or jump_id is None:
            continue
        message = _norm_text(msg_row[2])
        status = _norm_upper(msg_row[4])
        bib = msg_row[5]
        if status not in _STATUS_VALUES or _is_empty(bib):
            continue

        block = _find_block_by_jump_id(data, jump_id)
        msg_datetime = parse_date_time(msg_row[0], msg_row[1])
        if block is None or msg_datetime is None:
            continue
        target_idx = _nearest_row_in_block(data, block, msg_datetime, bib)
        if target_idx is None:
            _, changed = _insert_missing_starter_row(
                data, changed, msg_row, block
            )
            continue

        if data[target_idx][24] != status:
            data[target_idx][24] = status
            changed.add((target_idx, 24))
        if _append_msg(data[target_idx], message):
            changed.add((target_idx, 26))

    newest_messages = {}

    # Ordinary messages cannot change row count, so indexes remain stable.
    for msg_position, (msg_row, jump_id) in enumerate(
        zip(msg_cleandata, jump_ids)
    ):
        if len(msg_row) < 8 or jump_id is None:
            continue
        msg_status = _norm_upper(msg_row[4])
        if msg_status in _STATUS_VALUES or msg_status == _TRIALJUMPER_STATUS:
            continue

        message = _norm_text(msg_row[2])
        bib = msg_row[5]
        gate = msg_row[6]
        msg_datetime = parse_date_time(msg_row[0], msg_row[1])
        block = _find_block_by_jump_id(data, jump_id)
        if msg_datetime is None or block is None:
            continue

        if _is_empty(bib):
            if not message or _SKIPPED_MSG_WORDS.search(message):
                continue
            target_idx = _nearest_row_in_block(data, block, msg_datetime)
            if target_idx is None:
                continue
            candidate = (msg_datetime, msg_position, message)
            if (
                target_idx not in newest_messages
                or candidate[:2] > newest_messages[target_idx][:2]
            ):
                newest_messages[target_idx] = candidate
            continue

        target_idx = _nearest_row_in_block(data, block, msg_datetime, bib)
        if target_idx is None:
            continue
        if message and _append_msg(data[target_idx], message):
            changed.add((target_idx, 26))
        if not _is_empty(gate) and data[target_idx][25] != gate:
            data[target_idx][25] = gate
            changed.add((target_idx, 25))

    for target_idx, (_, _, message) in newest_messages.items():
        if data[target_idx][26] != message:
            data[target_idx][26] = message
            changed.add((target_idx, 26))

    return changed


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def sync_messages_with_fol(
    fol_calcdata: list[list],
    msg_cleandata: list[list] | None = None,
) -> tuple[list[list], set[tuple[int, int]]]:
    """Return a copy of fol_calcdata with msg_clean assignments applied."""
    data = copy.deepcopy(fol_calcdata)
    changed: set[tuple[int, int]] = set()
    if msg_cleandata:
        changed = _loop4_sync_msg_with_fol(data, msg_cleandata, changed)
    return data, changed


def run_calculation(
    fol_calcdata: list[list],
    msg_cleandata: list[list] | None = None,
) -> tuple[list[list], set[tuple[int, int]]]:
    """
    Run data imputation on fol_calcdata.

    Args:
        fol_calcdata: list of lists (33 columns per row, may contain formula strings)
        msg_cleandata: optional list of message rows to sync into fol_calcdata

    Returns:
        (updated_data, changed_cells) where:
        - updated_data is a deep copy with imputed values (formulas resolved to numbers)
        - changed_cells is a set of (row, col) tuples indicating which cells were modified
    """
    data = copy.deepcopy(fol_calcdata)
    changed: set[tuple[int, int]] = set()

    # Pre-processing: temporarily propagate DG (col 0) and Typ (col 1) down to follow-up rows
    current_dg = None
    current_typ = None
    for row in data:
        if row[0] is not None and str(row[0]).strip() != "":
            current_dg = row[0]
        else:
            row[0] = current_dg

        if row[1] is not None and str(row[1]).strip() != "":
            current_typ = row[1]
        else:
            row[1] = current_typ

    # Pre-processing: resolve formula strings to numeric values
    _resolve_formulas(data)

    # Loop 1: Jump-level single checks
    _loop1_single_checks(data, changed)

    # Loop 2: DG-grouped checks
    _loop2_dg_grouped(data, changed)

    # Loop 3: Multi-variable calculation & dependency resolution
    _loop3_multi_variable(data, changed)

    # Restore explicit DG/Typ block starts before Jump_ID matching. Loop 4 may
    # then safely insert rows while preserving those markers.
    for row_index, row in enumerate(data):
        if _is_empty(fol_calcdata[row_index][0]):
            row[0] = None
        if _is_empty(fol_calcdata[row_index][1]):
            row[1] = None

    # Loop 4: sync workspace_msg with workspace_fol
    if msg_cleandata:
        changed = _loop4_sync_msg_with_fol(data, msg_cleandata, changed)

    return data, changed
