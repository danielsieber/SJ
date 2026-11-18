"""Check jumpers and hills against SJ-Listen.xlsm.

- Liste_Springer_ID (table from the row with the header 'JumperID'): every
  jumper of the sheet Ranglisten must be listed. A jumper is found by name
  and NSA; the JumperID comes from the list.
- Liste_Schanzen_ID (header row with 'SchanzenID'): the hill of each PDF
  ('Einsiedeln HS117', built from Place and Hill Size) gives the value
  'Pts / Gate für DGW'.

The workbook is only read (openpyxl, cached values of formulas), so it may be
open in Excel.
"""

from __future__ import annotations

import re
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sj_paths import DATA_DIRECTORY  # noqa: E402


SJ_LISTEN_FILE = DATA_DIRECTORY / "SJ-Listen.xlsm"
JUMPER_SHEET = "Liste_Springer_ID"
HILL_SHEET = "Liste_Schanzen_ID"
DGW_HEADER = "Pts / Gate für DGW"
COUNTRY_SUFFIX = re.compile(r"\s*\([A-Z]{3}\)\s*$")


class ListenError(ValueError):
    """Raised when SJ-Listen.xlsm or one of its lists cannot be read."""


def _key(text: Any) -> str:
    """Compare names without case and repeated spaces."""
    return " ".join(str(text or "").split()).casefold()


@dataclass(frozen=True)
class Jumper:
    jumper_id: str
    name: str
    nsa: str
    fis_code: Any


@dataclass
class SjListen:
    jumper_ids: dict[tuple[str, str], str] = field(default_factory=dict)
    hill_dgw: dict[str, Any] = field(default_factory=dict)
    jumpers: list[Jumper] = field(default_factory=list)
    hills: dict[str, str] = field(default_factory=dict)  # key -> SchanzenID as written

    def all_by_fis_code(self, fis_code: Any) -> list[Jumper]:
        """All entries with this FIS code (a name change may be listed twice)."""
        code = str(fis_code).strip()
        return [jumper for jumper in self.jumpers if str(jumper.fis_code).strip() == code]

    def by_fis_code(self, fis_code: Any) -> Jumper | None:
        found = self.all_by_fis_code(fis_code)
        return found[0] if found else None

    def jumper_id(self, name: Any, nsa: Any) -> str | None:
        return self.jumper_ids.get((_key(name), _key(nsa)))

    def by_jumper_id(self, jumper_id: str) -> list[Jumper]:
        return [jumper for jumper in self.jumpers if _key(jumper.jumper_id) == _key(jumper_id)]

    def resolve(self, name: Any, nsa: Any) -> tuple[str | None, str | None]:
        """JumperID of an athlete of a PDF: (JumperID, problem).

        The ID is built by the rule from NSA and name and looked up in the
        list. It is only accepted when a listed jumper with this ID has the
        same surname (first word); 'AUT STEINA SIMON' for STEINBERGER Simon is
        not STEINACHER Simon. When the rule ID is not listed, name and NSA are
        looked up (special IDs such as 'ITA DEL LUCREZ').
        """
        rule_id = proposed_jumper_id(name, nsa)
        listed = self.by_jumper_id(rule_id)
        if any(_first_word(jumper.name) == _first_word(name) for jumper in listed):
            return listed[0].jumper_id, None
        # 'CHERVET Jules-Lucien' is listed as 'CHERVET Jules' (FRA CHERVE JULES).
        short_name = " ".join(word.split("-")[0] for word in str(name or "").split())
        short = [j for j in self.by_jumper_id(proposed_jumper_id(short_name, nsa))
                 if _first_word(j.name) == _first_word(name)]
        if short and not listed:
            return short[0].jumper_id, None
        initials = self._by_initials(name, nsa)
        if initials is not None and not listed:
            return initials, None
        if listed:
            others = ", ".join(f"{jumper.name} ({jumper.nsa})" for jumper in listed)
            return None, f"JumperID '{rule_id}' gehört in Liste_Springer_ID zu {others}"
        special = self.jumper_id(name, nsa)
        if special is not None:
            return special, None
        return None, f"JumperID '{rule_id}' fehlt in Liste_Springer_ID"

    def _by_initials(self, name: Any, nsa: Any) -> str | None:
        """Olympic lists print first names as initials: 'SUNDAL KE' is
        SUNDAL Kristoffer Eriksen, 'SCHARFENBERG AF' SCHARFENBERG Anna-Fay.
        Only a single listed jumper of the nation must fit."""
        words = str(name or "").split()
        if len(words) != 2 or not re.fullmatch(r"[A-Z]{1,3}", words[1]):
            return None
        found = [
            jumper for jumper in self.jumpers
            if _key(jumper.nsa) == _key(nsa) and _first_word(jumper.name) == _first_word(name)
            and "".join(part[0] for part in re.split(r"[\s-]+", " ".join(jumper.name.split()[1:])) if part).upper()
            == words[1]
        ]
        ids = {jumper.jumper_id for jumper in found}
        return ids.pop() if len(ids) == 1 else None

    def has_hill(self, hill: str) -> bool:
        return _key(hill) in self.hill_dgw

    def dgw(self, hill: str) -> Any:
        return self.hill_dgw.get(_key(hill))


ID_LETTERS = 6  # letters taken from the first and the second word of the name


def proposed_jumper_id(name: Any, nsa: Any) -> str:
    """JumperID by the rule of Liste_Springer_ID: NSA, then the first six
    letters of the first and of the second word of the name.

    'ACHLEITNER Alexander' (AUT) -> 'AUT ACHLEI ALEXAN'.
    """
    words = str(name or "").split()
    parts = [str(nsa or "").strip().upper()] + [word[:ID_LETTERS].upper() for word in words[:2]]
    return " ".join(part for part in parts if part)


def _first_word(name: Any) -> str:
    words = str(name or "").split()
    return words[0].casefold() if words else ""


def hill_name(place: Any, hill_size: Any) -> str | None:
    """'Einsiedeln (SUI)', 117 -> 'Einsiedeln HS117' (the SchanzenID)."""
    if not place or hill_size is None:
        return None
    if isinstance(hill_size, float) and hill_size.is_integer():
        hill_size = int(hill_size)
    return f"{COUNTRY_SUFFIX.sub('', str(place)).strip()} HS{hill_size}"


def _table(workbook, sheet_name: str, key_header: str) -> tuple[dict[str, int], Iterable]:
    """Return ({header: column index}, data rows) of the list in a sheet."""
    if sheet_name not in workbook.sheetnames:
        raise ListenError(f"Sheet {sheet_name} fehlt in {SJ_LISTEN_FILE.name}.")
    rows = workbook[sheet_name].iter_rows(values_only=True)
    for row in rows:
        if row and _key(row[0]) == _key(key_header):
            # Headers may contain line breaks ('Pts / Gate \nfür DGW').
            columns = {_key(value): index for index, value in enumerate(row) if value}
            return columns, rows
    raise ListenError(f"Spalte {key_header} fehlt im Sheet {sheet_name}.")


def read_sj_listen(path: Path) -> SjListen:
    if not path.is_file():
        raise ListenError(f"{path} nicht gefunden.")
    try:
        with warnings.catch_warnings():
            # openpyxl warns about print areas and data validation of the xlsm.
            warnings.simplefilter("ignore", UserWarning)
            workbook = load_workbook(path, read_only=True, data_only=True)
    except OSError as error:
        raise ListenError(f"{path.name} kann nicht gelesen werden: {error}") from error
    try:
        listen = SjListen()
        columns, rows = _table(workbook, JUMPER_SHEET, "JumperID")
        for needed in ("name", "nsa"):
            if needed not in columns:
                raise ListenError(f"Spalte {needed.capitalize()} fehlt im Sheet {JUMPER_SHEET}.")
        fis_column = columns.get("fis_code")
        for row in rows:
            jumper_id = row[0]
            if jumper_id:
                key = (_key(row[columns["name"]]), _key(row[columns["nsa"]]))
                listen.jumper_ids[key] = str(jumper_id).strip()
                listen.jumpers.append(Jumper(
                    str(jumper_id).strip(),
                    str(row[columns["name"]] or ""),
                    str(row[columns["nsa"]] or ""),
                    row[fis_column] if fis_column is not None else None,
                ))

        columns, rows = _table(workbook, HILL_SHEET, "SchanzenID")
        dgw_column = columns.get(_key(DGW_HEADER))
        if dgw_column is None:
            raise ListenError(f"Spalte '{DGW_HEADER}' fehlt im Sheet {HILL_SHEET}.")
        for row in rows:
            if row and row[0]:
                listen.hill_dgw[_key(row[0])] = row[dgw_column]
                listen.hills[_key(row[0])] = str(row[0])
        return listen
    finally:
        workbook.close()
