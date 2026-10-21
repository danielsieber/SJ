"""
New Weekend view: codex file selection, validation table, and analysis trigger.
"""
from pathlib import Path
import re

import pandas as pd
import streamlit as st

from src.config import load_reference_lists
from src.pipeline.phase_10_codex_list import parse_filename, group_and_sort_by_weekend
from src.utils.excel_automation import add_schanze_to_excel, add_springer_to_excel


@st.cache_data
def _parse_and_group_files(list_Data_FIS):
    """Cached parsing and grouping of FIS filenames."""
    parsed_files = [parse_filename(filename, None) for filename in list_Data_FIS]
    return group_and_sort_by_weekend(parsed_files)


def render_new_weekend(
    list_Data_FIS, path_basic, path_Data_FIS, path_Weekend, path_Data_PDF
):
    """Renders the New Weekend tab with file selection, validation, and analysis."""
    # 1. Parsing von list_Data_FIS (cached)
    sorted_files = _parse_and_group_files(tuple(list_Data_FIS))
    
    # Erstelle ein Mapping von formatiertem Label zu Original-Dateinamen
    options_map = {}
    for f in sorted_files:
        weekend_nr = f.get("Weekend")
        filename = f.get("Filename", "")
        gender = f.get("Gender", "")
        cup = f.get("Cup", "")
        
        if weekend_nr is not None:
            label = f"WK {weekend_nr} – {filename} – {gender} – {cup}"
        else:
            label = f"Invalid – {filename}"
        options_map[label] = filename

    # 2. Erscheint im Workspace, oben in der Mitte ein Multiselect Dropdown Menu mit den neuen list_Data_FIS
    col_l, col_m, col_r = st.columns([1, 4, 1])
    with col_m:
        # Bestimme Standard-Auswahl anhand der bereits gespeicherten Dateinamen in st.session_state.list_weekend_codex
        default_selection = [
            lbl for lbl, fn in options_map.items() 
            if fn in st.session_state.list_weekend_codex
        ]
        
        selected_labels = st.multiselect(
            "Wählen Sie die Codex-Dateien für das neue Weekend aus:",
            options=list(options_map.keys()),
            default=default_selection,
            key="new_weekend_multiselect"
        )
        
        # 3. Speicher die ausgewählten Dateinamen in der Liste list_weekend_codex
        selected_filenames = [options_map[lbl] for lbl in selected_labels]
        st.session_state.list_weekend_codex = selected_filenames
        
        # Falls Dateien ausgewählt wurden, zeige diese in einem Dataframe an
        if selected_filenames:
            st.markdown('<div style="margin-top: 280px; font-size: 14px; font-weight: normal;">Ausgewählte Dateien:</div>', unsafe_allow_html=True)
            selected_info = [f for f in sorted_files if f["Filename"] in selected_filenames]
            df_selected = pd.DataFrame(selected_info)
            
            df_selected = df_selected.rename(columns={"Weekend": "Nr."})

            cols_order = ["Nr.", "Filename", "DateFile", "Codex", "Cup", "Gender", "HS", "Place"]
            cols_present = [c for c in cols_order if c in df_selected.columns]

            selection_column_config = {
                "Nr.": st.column_config.NumberColumn(width=60, format="%d"),
                "Filename": st.column_config.Column(width="large"),
                "DateFile": st.column_config.Column(width="small"),
                "Codex": st.column_config.Column(width="small"),
                "Cup": st.column_config.Column(width=60),
                "Gender": st.column_config.Column(width="small"),
                "HS": st.column_config.Column(width=60),
                "Place": st.column_config.Column(width="small"),
            }

            st.data_editor(
                df_selected[cols_present],
                num_rows="fixed",
                use_container_width=True,
                height=200,
                row_height=24,
                column_config=selection_column_config,
                column_order=cols_present,
                disabled=True,
                key="selected_files_editor",
            )

            _, col_analyze_btn = st.columns([5, 2.5])
            with col_analyze_btn:
                if st.button("analyze", key="check_data_btn", type="secondary", use_container_width=True):
                    if not st.session_state.list_weekend_codex:
                        st.warning("Bitte wählen Sie zuerst mindestens eine Datei aus!")
                    else:
                        _run_analysis(
                            path_basic,
                            path_Data_FIS,
                            path_Weekend,
                            path_Data_PDF,
                        )


def _run_analysis(path_basic, path_Data_FIS, path_Weekend, path_Data_PDF):
    """Runs the full analysis pipeline and transitions to the editing stage."""
    with st.spinner("Analysiere Raw-Daten und speichere..."):
        try:
            from src.pipeline.phase_20_preprocessing.import_raw import main_analysis_for_files
            from src.pipeline.phase_20_preprocessing.repair_data import repair_data
            from src.pipeline.phase_20_preprocessing.clean_data import clean_msg_data, clean_fol_data
            from src.pipeline.phase_20_preprocessing.calc_data import create_fol_calcdata
            
            # 1. Import
            result = main_analysis_for_files(
                st.session_state.list_weekend_codex,
                str(path_Data_FIS),
                str(path_Weekend)
            )
            
            if result:
                fol_rawdata, msg_rawdata, output_filename, output_path = result
                
                # 2. Repair
                fol_repairdata = repair_data(fol_rawdata)
                
                # 3. Clean
                msg_cleandata = clean_msg_data(msg_rawdata)
                fol_cleandata = clean_fol_data(fol_repairdata)
                
                # 4. Calculate
                fol_calcdata = create_fol_calcdata(fol_cleandata, msg_cleandata)

                # 5. Validate the current SchanzenID and every SpringerID against
                # a fresh, uncached read of Listen.xlsm.
                reference_data = _validate_reference_ids(
                    Path(path_basic) / "Listen.xlsm",
                    st.session_state.list_weekend_codex,
                    fol_cleandata,
                )
                if reference_data is None:
                    return

                current_schanzen_data, list_springer_id_actual = reference_data
                
                # Clean previous editing session state keys to force regeneration with new columns
                for key in ["fol_cleandata33column", "editor_version"]:
                    if key in st.session_state:
                        del st.session_state[key]
                # Clean versioned data_editor widget keys
                for key in list(st.session_state.keys()):
                    if key.startswith("fol_editor_") or key.startswith("msg_editor_"):
                        del st.session_state[key]
                
                # Set values to session state for manual editing stage
                st.session_state.fol_rawdata = fol_rawdata
                st.session_state.msg_rawdata = msg_rawdata
                st.session_state.fol_repairdata = fol_repairdata
                st.session_state.fol_cleandata = fol_cleandata
                st.session_state.fol_calcdata = fol_calcdata
                st.session_state.msg_cleandata = msg_cleandata
                st.session_state.output_filename = output_filename
                st.session_state.output_path = output_path
                st.session_state.path_Data_PDF = path_Data_PDF
                st.session_state.filename_weekend = Path(output_filename).stem
                st.session_state.current_schanzen_data = current_schanzen_data
                st.session_state.list_Springer_ID_actual = list_springer_id_actual
                st.session_state.show_editing_stage = True
                st.rerun()
            else:
                st.error("Fehler: Keine Daten konnten importiert werden.")
        except Exception as e:
            st.error(f"Fehler bei der Analyse: {e}")


def _normalized(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip().upper()


def _find_column(dataframe, *names):
    """Finds a column case-insensitively and ignores non-alphanumeric signs."""
    wanted = {
        re.sub(r"[^a-z0-9]", "", str(name).lower())
        for name in names
    }
    for column in dataframe.columns:
        normalized = re.sub(r"[^a-z0-9]", "", str(column).lower())
        if normalized in wanted:
            return column
    return None


def _current_schanze(selected_filenames):
    """Returns (SchanzenID, HS) using the same top-file rule as the importer."""
    parsed_files = [parse_filename(filename, None) for filename in selected_filenames]

    def sort_key(item):
        try:
            codex = int(item.get("Codex", 0) or 0)
        except (TypeError, ValueError):
            codex = 0
        return item.get("DateFile") or "", codex

    parsed_files.sort(key=sort_key, reverse=True)
    top_item = parsed_files[0]
    place = str(top_item.get("Place", "")).strip()
    hs = str(top_item.get("HS", "")).strip()
    return f"{place} HS{hs}", hs


def _build_springer_id_actual(fol_cleandata):
    """
    Builds the unique actual Springer list from the 28-column fol_clean data.

    Invalid source rows include the original FIS row number in the error.
    """
    springer_rows = []
    seen_ids = set()

    for list_index, row in enumerate(fol_cleandata):
        codex = row[0] if len(row) > 0 else None
        sheet = row[1] if len(row) > 1 else None
        source_row = row[2] if len(row) > 2 else list_index + 2
        name = row[7] if len(row) > 7 else None
        vorname = row[8] if len(row) > 8 else None
        nsa = row[9] if len(row) > 9 else None

        missing = []
        if not _normalized(nsa):
            missing.append("NSA")
        if not _normalized(name):
            missing.append("Name")
        if not _normalized(vorname):
            missing.append("Vorname")
        if missing:
            raise ValueError(
                "Ungültige Quelldaten in "
                f"Codex '{codex}', Sheet '{sheet}', row nr {source_row}: "
                f"{', '.join(missing)} fehlt."
            )

        springer_id = (
            f"{_normalized(nsa)} "
            f"{_normalized(name)[:5]} "
            f"{_normalized(vorname)[:5]}"
        )
        if springer_id in seen_ids:
            continue

        seen_ids.add(springer_id)
        springer_rows.append(
            {
                "SpringerID": springer_id,
                "Name": str(name).strip(),
                "Vorname": str(vorname).strip(),
                "NSA": str(nsa).strip().upper(),
            }
        )

    return pd.DataFrame(
        springer_rows, columns=["SpringerID", "Name", "Vorname", "NSA"]
    )


def _validate_reference_ids(listen_file, selected_filenames, fol_cleandata):
    """
    Runs steps 9 and 10.

    Returns the current Schanzen row and enriched actual Springer list when
    everything is valid. Returns None after opening Listen.xlsm for a missing
    entry, because the user must save it manually and rerun Analyze.
    """
    try:
        list_schanzen_id, list_springer_id = load_reference_lists(listen_file)
    except Exception as exc:
        st.error(str(exc))
        return None

    schanzen_column = _find_column(list_schanzen_id, "SchanzenID", "Schanzen_ID")
    if schanzen_column is None:
        st.error("Spalte 'SchanzenID' fehlt im Sheet 'Liste_Schanzen_ID'.")
        return None

    schanzen_id, hs = _current_schanze(selected_filenames)
    schanzen_mask = (
        list_schanzen_id[schanzen_column].map(_normalized)
        == _normalized(schanzen_id)
    )
    if not schanzen_mask.any():
        try:
            add_schanze_to_excel(listen_file, schanzen_id, hs)
        except Exception as exc:
            st.error(f"Fehler beim Ergänzen der SchanzenID '{schanzen_id}': {exc}")
            return None
        st.warning(
            f"SchanzenID '{schanzen_id}' wurde in Listen.xlsm ergänzt. "
            "Bitte Listen.xlsm manuell speichern und anschließend erneut "
            "auf 'analyze' klicken."
        )
        return None

    current_schanzen_data = list_schanzen_id.loc[schanzen_mask].iloc[[0]].copy()

    try:
        actual = _build_springer_id_actual(fol_cleandata)
    except ValueError as exc:
        st.error(str(exc))
        return None

    springer_id_column = _find_column(
        list_springer_id,
        "SpringerID",
        "Springer_ID",
        "JumperID",
        "Jumper_ID",
    )
    demographic_columns = {
        output_name: _find_column(list_springer_id, output_name)
        for output_name in ["Gender", "DoB", "JG"]
    }
    if springer_id_column is None:
        st.error(
            "Spalte 'SpringerID' oder 'JumperID' fehlt im Sheet "
            "'Liste_Springer_ID'."
        )
        return None
    missing_columns = [
        name for name, column in demographic_columns.items() if column is None
    ]
    if missing_columns:
        st.error(
            "Folgende Spalten fehlen im Sheet 'Liste_Springer_ID': "
            + ", ".join(missing_columns)
        )
        return None

    reference_by_id = {}
    for _, reference_row in list_springer_id.iterrows():
        key = _normalized(reference_row.get(springer_id_column))
        if key and key not in reference_by_id:
            reference_by_id[key] = reference_row

    for column in ["Gender", "DoB", "JG"]:
        actual[column] = None

    for index, actual_row in actual.iterrows():
        key = _normalized(actual_row["SpringerID"])
        reference_row = reference_by_id.get(key)
        if reference_row is None:
            full_name = f"{actual_row['Name']} {actual_row['Vorname']}".strip()
            try:
                add_springer_to_excel(
                    listen_file,
                    actual_row["SpringerID"],
                    full_name,
                    actual_row["NSA"],
                )
            except Exception as exc:
                st.error(
                    f"Fehler beim Ergänzen der SpringerID "
                    f"'{actual_row['SpringerID']}': {exc}"
                )
                return None
            st.warning(
                f"SpringerID '{actual_row['SpringerID']}' wurde alphabetisch "
                "in Listen.xlsm eingefügt. Bitte Listen.xlsm manuell speichern "
                "und anschließend erneut auf 'analyze' klicken."
            )
            return None

        for output_name, source_column in demographic_columns.items():
            actual.at[index, output_name] = reference_row.get(source_column)

    st.toast("Alle Schanzen- und SpringerIDs sind gültig.")
    return current_schanzen_data, actual
