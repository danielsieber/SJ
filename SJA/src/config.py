"""
Path detection and data loading for the SJA application.
"""
from pathlib import Path
import os
import re
import pandas as pd
import streamlit as st


def _reference_header_key(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _read_reference_sheet(listen_file, sheet_name, id_headers):
    """
    Reads a reference sheet and detects its real table header row.

    Listen.xlsm has descriptive rows above its tables, so relying on one fixed
    Excel row is brittle. Search the first 50 rows for the sheet's ID header.
    """
    raw = pd.read_excel(listen_file, sheet_name=sheet_name, header=None)
    wanted_headers = {_reference_header_key(header) for header in id_headers}
    header_index = None

    for index, row in raw.head(50).iterrows():
        row_headers = {_reference_header_key(value) for value in row.tolist()}
        if wanted_headers.intersection(row_headers):
            header_index = index
            break

    if header_index is None:
        visible_values = []
        for value in raw.head(15).to_numpy().flatten():
            key = _reference_header_key(value)
            if key and key not in visible_values:
                visible_values.append(key)
        raise ValueError(
            f"Keine passende ID-Überschrift im Sheet '{sheet_name}' gefunden. "
            f"Erkannte Inhalte: {', '.join(visible_values[:12]) or 'keine'}"
        )

    headers = []
    for column_index, value in enumerate(raw.iloc[header_index].tolist(), start=1):
        if value is None or pd.isna(value) or not str(value).strip():
            headers.append(f"Unnamed_{column_index}")
        else:
            headers.append(str(value).strip())

    data = raw.iloc[header_index + 1:].copy()
    data.columns = headers
    return data.dropna(how="all").reset_index(drop=True)


def load_reference_lists(listen_file):
    """
    Reads the current Schanzen and Springer reference data without caching.

    This is intentionally uncached so an analysis rerun sees changes that the
    user has just saved manually in Listen.xlsm.
    """
    list_schanzen_id = pd.DataFrame(columns=["SchanzenID"])
    list_springer_id = pd.DataFrame(
        columns=["SpringerID", "Name", "NSA", "Gender", "DoB", "JG"]
    )

    listen_file = Path(listen_file)
    if not listen_file.exists() or not listen_file.is_file():
        raise FileNotFoundError(f"Listen.xlsm nicht gefunden: {listen_file}")

    try:
        list_schanzen_id = _read_reference_sheet(
            listen_file,
            "Liste_Schanzen_ID",
            ("SchanzenID", "Schanzen_ID"),
        )
        list_springer_id = _read_reference_sheet(
            listen_file,
            "Liste_Springer_ID",
            ("SpringerID", "Springer_ID", "JumperID", "Jumper_ID"),
        )
    except Exception as exc:
        raise RuntimeError(f"Listen.xlsm konnte nicht gelesen werden: {exc}") from exc

    return list_schanzen_id, list_springer_id


@st.cache_data
def init_paths():
    """
    Uses SJA_WEB_ROOT when configured and returns the derived data folders.
    """
    configured_root = os.environ.get("SJA_WEB_ROOT")
    candidates = {"SJA_WEB_ROOT": configured_root} if configured_root else {}
    
    path_basic = None
    found_key = None
    
    for key, path_str in candidates.items():
        p = Path(path_str)
        if p.exists() and p.is_dir():
            path_basic = p
            found_key = key
            break
            
    # Fallback-Logic, falls lokal getestet wird und Pfade nicht existieren
    if not path_basic:
        # Versuche relativen Pfad
        fallback_path = Path(__file__).parent.parent / "Family" / "Sport" / "FIS" / "Web"
        if fallback_path.exists():
            path_basic = fallback_path
            found_key = "Fallback-Relative"
        else:
            # Zeige Fehler in der App und erstelle Scheinstruktur für den lokalen Start
            path_basic = Path("C:/") # Platzhalter
            found_key = "None"
            
    # Abgeleitete Pfade (..\\Data_FIS und ..\\Data_Weekend)
    if found_key != "None":
        # Prüfe zuerst, ob die Ordner direkt im Web-Ordner liegen (z.B. C:\\...\\web\\Data_Weekend)
        p_fis = path_basic / "Data_FIS"
        if p_fis.exists() and p_fis.is_dir():
            path_Data_FIS = p_fis
        else:
            path_Data_FIS = path_basic.parent / "Data_FIS"
            
        p_wk = path_basic / "Data_Weekend"
        if p_wk.exists() and p_wk.is_dir():
            path_Weekend = p_wk
        else:
            path_Weekend = path_basic.parent / "Data_Weekend"
            
        path_Data_FIS = path_Data_FIS.resolve()
        path_Weekend = path_Weekend.resolve()
        # Keep the configurable data destinations together. Change this line
        # if PDF files should later be stored somewhere else.
        path_Data_PDF = (path_basic / "Data_pdf").resolve()
    else:
        # Fallback im aktuellen Workspace-Verzeichnis, um Abstürze zu verhindern
        path_Data_FIS = Path("./data/Data_FIS")
        path_Weekend = Path("./data/Data_Weekend")
        path_Data_PDF = Path("./data/Data_pdf")
        
    return path_basic, path_Data_FIS, path_Weekend, path_Data_PDF, found_key


@st.cache_data
def load_data_lists(path_basic, path_Data_FIS, path_Weekend):
    """
    Creates file lists and reads the Schanzen- and Springer-IDs from Listen.xlsm.
    """
    # 1. list_Data_FIS (Dateinamen aus path_Data_FIS)
    list_Data_FIS = []
    if path_Data_FIS.exists() and path_Data_FIS.is_dir():
        list_Data_FIS = [f.name for f in path_Data_FIS.iterdir() if f.is_file() and f.suffix.lower() == '.xlsx']
    
    # 2. list_Weekend_Folder (Dateinamen aus path_Weekend ohne .xlsx)
    list_Weekend_Folder = []
    if path_Weekend.exists() and path_Weekend.is_dir():
        list_Weekend_Folder = [f.stem for f in path_Weekend.iterdir() if f.is_file() and f.suffix.lower() == '.xlsx']
        # Sortiere absteigend nach Datum (die ersten 10 Zeichen)
        list_Weekend_Folder.sort(reverse=True)
        
    # 3. Listen aus Listen.xlsm (Liste_Schanzen_ID)
    listen_file = path_basic / "Listen.xlsm"
    list_Schanzen_ID = pd.DataFrame(columns=["SchanzenID"])
    
    if listen_file.exists() and listen_file.is_file():
        try:
            list_Schanzen_ID, _ = load_reference_lists(listen_file)
        except Exception:
            pass # Fehler wird bei Bedarf im Interface ausgegeben
            
    return list_Data_FIS, list_Weekend_Folder, list_Schanzen_ID
