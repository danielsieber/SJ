"""Shared Jump_ID block construction and message matching."""

import re
from datetime import date, datetime, time, timedelta


def normalize_codex(value) -> str:
    """Normalize numeric and text Codex values for matching."""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)) and float(value).is_integer():
        return str(int(value))

    text = str(value).strip()
    try:
        numeric = float(text)
    except ValueError:
        return text.casefold()
    return str(int(numeric)) if numeric.is_integer() else text.casefold()


def parse_date_time(date_value, time_value) -> datetime | None:
    """Combine the date/time representations used by cleaned data."""
    if isinstance(date_value, datetime):
        if time_value is None:
            return date_value
        parsed_date = date_value.date()
    elif isinstance(date_value, date):
        parsed_date = date_value
    else:
        text = str(date_value).strip() if date_value is not None else ""
        parsed_date = None
        if text:
            try:
                parsed_date = datetime.fromisoformat(text).date()
            except ValueError:
                for date_format in ("%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
                    try:
                        parsed_date = datetime.strptime(text, date_format).date()
                        break
                    except ValueError:
                        pass

    if parsed_date is None:
        return None

    if isinstance(time_value, datetime):
        parsed_time = time_value.time()
    elif isinstance(time_value, time):
        parsed_time = time_value
    else:
        text = str(time_value).strip() if time_value is not None else ""
        parsed_time = None
        if text:
            try:
                parsed_time = time.fromisoformat(text)
            except ValueError:
                pass

    if parsed_time is None:
        return None
    return datetime.combine(parsed_date, parsed_time)


def normalized_dg(value) -> int | None:
    """Return a numeric DG from plain or visually decorated content."""
    if value is None or isinstance(value, bool):
        return None
    match = re.search(r"(-?\d+(?:\.0+)?)\s*$", str(value).strip())
    if match is None:
        return None
    numeric = float(match.group(1))
    return int(numeric) if numeric.is_integer() else None


def build_jump_blocks(
    fol_calcdata: list[list], filename_stem: str = ""
) -> list[dict]:
    """Build Jump_ID blocks, timestamps, Codex values, and row indexes."""
    blocks = []
    current = None

    for row_index, row in enumerate(fol_calcdata or []):
        raw_dg = row[0] if len(row) > 0 else None
        is_explicit_start = (
            raw_dg is not None
            and str(raw_dg).strip() != ""
            and not str(raw_dg).strip().startswith("=")
        )
        if is_explicit_start:
            dg = normalized_dg(raw_dg)
            typ = (
                str(row[1]).strip().upper()
                if len(row) > 1 and row[1] is not None
                else ""
            )
            current = None
            if dg is not None and typ:
                current = {
                    "jump_id": f"{filename_stem} {dg:02d}{typ}".strip(),
                    "dg": dg,
                    "typ": typ,
                    "start": None,
                    "end": None,
                    "codexes": set(),
                    "row_indexes": [],
                }
                blocks.append(current)

        if current is None:
            continue

        current["row_indexes"].append(row_index)
        codex = normalize_codex(row[27] if len(row) > 27 else None)
        if codex:
            current["codexes"].add(codex)

        row_datetime = parse_date_time(
            row[28] if len(row) > 28 else None,
            row[29] if len(row) > 29 else None,
        )
        if row_datetime is not None:
            if current["start"] is None or row_datetime < current["start"]:
                current["start"] = row_datetime
            if current["end"] is None or row_datetime > current["end"]:
                current["end"] = row_datetime

    return blocks


def assign_msg_block_indexes(
    msg_cleandata: list[list], jump_blocks: list[dict]
) -> list[int | None]:
    """Assign messages to nearest blocks, with Codex as the outer boundary."""
    assignments = [None] * len(msg_cleandata or [])
    messages_by_codex = {}

    for row_index, msg_row in enumerate(msg_cleandata or []):
        codex = normalize_codex(msg_row[7] if len(msg_row) > 7 else None)
        row_datetime = parse_date_time(
            msg_row[0] if len(msg_row) > 0 else None,
            msg_row[1] if len(msg_row) > 1 else None,
        )
        if codex and row_datetime is not None:
            messages_by_codex.setdefault(codex, []).append(
                (row_index, row_datetime)
            )

    for codex, messages in messages_by_codex.items():
        codex_blocks = [
            (position, block)
            for position, block in enumerate(jump_blocks or [])
            if codex in block.get("codexes", set())
            and block.get("start") is not None
            and block.get("end") is not None
        ]
        for row_index, row_datetime in messages:
            candidates = []
            for position, block in codex_blocks:
                start = block["start"]
                end = block["end"]
                if row_datetime < start:
                    distance = start - row_datetime
                elif row_datetime > end:
                    distance = row_datetime - end
                else:
                    distance = timedelta(0)
                candidates.append((distance, position))

            if candidates:
                assignments[row_index] = min(candidates)[1]

    return assignments


def assign_msg_jump_ids(
    msg_cleandata: list[list], jump_blocks: list[dict]
) -> list[str | None]:
    """Return the matched Jump_ID for every msg_clean row."""
    return [
        jump_blocks[index]["jump_id"] if index is not None else None
        for index in assign_msg_block_indexes(msg_cleandata, jump_blocks)
    ]
