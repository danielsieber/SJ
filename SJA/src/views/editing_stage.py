"""
Editing stage view: manual data editing workspace with Streamlit data_editor tables.
"""
import datetime

import pandas as pd
from html import escape

import streamlit as st

from src.constants import calc_headers, clean_msg_headers, dg_summary_headers
from src.pipeline.phase_20_preprocessing.calc_data import create_fol_calcdata


def _expand_cleandata_to_33col(fol_cleandata, msg_cleandata=None):
    """
    Expands 28-column fol_cleandata to 33-column layout matching calc_headers order.
    Uses create_fol_calcdata for DG/Typ detection, then resolves all formulas to plain values.
    """
    calcdata = create_fol_calcdata(fol_cleandata, msg_cleandata)

    # Resolve formula references to plain values (top-to-bottom so refs chain correctly)
    for i in range(len(calcdata)):
        row = calcdata[i]
        for j in range(len(row)):
            val = row[j]
            if isinstance(val, str) and val.strip().startswith("="):
                if j == 14 and i < len(fol_cleandata):
                    # Jpts formula: use raw value from fol_cleandata
                    row[j] = fol_cleandata[i][26]
                else:
                    row[j] = None

    return calcdata


def _build_dg_summary(data):
    """Builds DG summary rows from fol_cleandata33column data.

    Iterates top-to-bottom. Each time DG changes (and is not None),
    a new summary row is created. Jumps counts ALL rows from first BIB
    to last BIB in that block. Jpts is the max of all Jpts values.
    Codex contains each distinct non-empty value in first-seen order.
    """
    import re

    def combine_date_time(date_value, time_value):
        parsed_date = None
        if isinstance(date_value, datetime.datetime):
            parsed_date = date_value.date()
        elif isinstance(date_value, datetime.date):
            parsed_date = date_value
        elif isinstance(date_value, str) and date_value.strip():
            try:
                parsed_date = datetime.date.fromisoformat(date_value.strip())
            except ValueError:
                pass

        parsed_time = None
        if isinstance(time_value, datetime.datetime):
            parsed_time = time_value.time()
        elif isinstance(time_value, datetime.time):
            parsed_time = time_value
        elif isinstance(time_value, str) and time_value.strip():
            try:
                parsed_time = datetime.time.fromisoformat(time_value.strip()[:8])
            except ValueError:
                pass

        if parsed_date is not None and parsed_time is not None:
            return datetime.datetime.combine(parsed_date, parsed_time)
        return None

    def clean_dg(val):
        if val is None:
            return None
        if isinstance(val, str):
            cleaned = re.sub(r'^[🔵🔹🔷\s]+', '', val).strip()
            return cleaned if cleaned else None
        return val

    # First pass: find DG start indices
    dg_starts = []
    prev_dg = None
    for i, row in enumerate(data):
        dg_val = clean_dg(row[0])
        if dg_val is not None and str(dg_val).strip() != "":
            if prev_dg is None or str(dg_val) != str(prev_dg):
                dg_starts.append(i)
                prev_dg = dg_val

    # Second pass: build summary for each DG block
    # Each block spans from its start index to (next DG start - 1) or end of data
    summary = []
    for block_idx, start_i in enumerate(dg_starts):
        end_i = dg_starts[block_idx + 1] - 1 if block_idx + 1 < len(dg_starts) else len(data) - 1
        first_row = data[start_i]
        last_row = data[end_i]
        dg_val = clean_dg(first_row[0])
        start_datetime = combine_date_time(first_row[28], first_row[29])
        end_datetime = combine_date_time(last_row[28], last_row[29])
        duration_min = None
        if start_datetime is not None and end_datetime is not None:
            duration_min = round((end_datetime - start_datetime).total_seconds() / 60, 2)

        gap_min = None
        if block_idx + 1 < len(dg_starts) and end_datetime is not None:
            next_first_row = data[dg_starts[block_idx + 1]]
            next_start_datetime = combine_date_time(next_first_row[28], next_first_row[29])
            if next_start_datetime is not None:
                gap_min = round((next_start_datetime - end_datetime).total_seconds() / 60, 2)

        block = {
            "dg": dg_val,
            "typ": first_row[1] if first_row[1] else None,
            "jpts_max": None,
            "codex": None,
            "jumps": end_i - start_i + 1,
            "bib_first": first_row[2] if first_row[2] else None,
            "bib_last": last_row[2] if last_row[2] else None,
            "date": first_row[28] if first_row[28] else None,
            "begin": str(first_row[29]) if first_row[29] else None,
            "end": str(last_row[29]) if last_row[29] else None,
            "duration_min": duration_min,
            "gap_min": gap_min,
        }

        # Find max Jpts and collect unique Codex values across the block.
        codex_values = []
        seen_codex = set()
        for j in range(start_i, end_i + 1):
            if data[j][27] is not None:
                codex = str(data[j][27]).strip()
                if codex and codex not in seen_codex:
                    seen_codex.add(codex)
                    codex_values.append(codex)

            if data[j][14] is not None:
                try:
                    jpts_val = float(data[j][14])
                    if block["jpts_max"] is None or jpts_val > block["jpts_max"]:
                        block["jpts_max"] = jpts_val
                except (ValueError, TypeError):
                    pass

        block["codex"] = ", ".join(codex_values) if codex_values else None

        summary.append(block)

    # Convert to list of lists matching dg_summary_headers order
    result = []
    for block in summary:
        result.append([
            block["dg"], block["typ"], block["jpts_max"], block["codex"],
            block["jumps"], block["bib_first"], block["bib_last"],
            block["date"], block["begin"], block["end"], block["duration_min"],
            block["gap_min"],
        ])
    return result


def _validate_dg_summary(dg_summary_data):
    """Validates sequential integer DG values and the allowed jump types."""
    errors = []
    allowed_types = {"T", "P", "Q", "W"}

    for position, row in enumerate(dg_summary_data, start=1):
        raw_dg = row[0] if len(row) > 0 else None
        raw_typ = row[1] if len(row) > 1 else None

        dg_value = None
        if raw_dg is not None and not isinstance(raw_dg, bool):
            cleaned_dg = str(raw_dg).strip()
            for marker in ("🔵", "🔹", "🔷"):
                cleaned_dg = cleaned_dg.lstrip(marker).strip()
            try:
                numeric_dg = float(cleaned_dg)
                if numeric_dg.is_integer():
                    dg_value = int(numeric_dg)
            except (TypeError, ValueError):
                pass

        if dg_value is None:
            display_value = "leer" if raw_dg is None or str(raw_dg).strip() == "" else repr(raw_dg)
            errors.append(
                f"DG Position {position}: Wert {display_value} ist keine Ganzzahl."
            )
        elif dg_value != position:
            errors.append(
                f"DG Position {position}: erwartet {position}, gefunden {dg_value}."
            )

        normalized_typ = str(raw_typ).strip().upper() if raw_typ is not None else ""
        if normalized_typ not in allowed_types:
            dg_label = dg_value if dg_value is not None else raw_dg
            display_typ = "leer" if not normalized_typ else repr(raw_typ)
            errors.append(
                f"Typ bei DG {dg_label}: ungültiger Wert {display_typ}; "
                "erlaubt sind T, P, Q oder W."
            )

    return errors


def _normalize_types_for_save(fol_calcdata):
    """Uppercases valid DG type values on each DG start row before export."""
    allowed_types = {"T", "P", "Q", "W"}
    for row in fol_calcdata:
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
        normalized_typ = str(row[1]).strip().upper() if row[1] is not None else ""
        if normalized_typ in allowed_types:
            row[1] = normalized_typ


def render_editing_stage():
    """Renders the manual editing workspace with Streamlit data_editor tables."""

    # ── 1. INITIALIZE DATA SOURCES ──
    if "fol_cleandata33column" not in st.session_state:
        st.session_state.fol_cleandata33column = _expand_cleandata_to_33col(
            st.session_state.fol_cleandata, st.session_state.msg_cleandata
        )

    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0

    if "fol_selected_row_key" not in st.session_state:
        st.session_state.fol_selected_row_key = ""

    # ── 2. WORKSPACE DG ── (placeholder filled after fol sync for fresh data)
    dg_placeholder = st.container()

    # ── 3. WORKSPACE FOL ──
    st.markdown(
        '<div style="margin-top: 10px; font-size: 0.78rem; font-family: Century Gothic, CenturyGothic, sans-serif; color: #CCCCCC;">workspace_fol (Daten: fol_cleandata33column)</div>',
        unsafe_allow_html=True,
    )

    df_fol = pd.DataFrame(
        st.session_state.fol_cleandata33column,
        columns=calc_headers,
    )

    # Configure custom column widths to make workspace_fol grid compact
    fol_column_config = {}
    for col in calc_headers:
        if col in ["Name", "Vorname", "MSG"]:
            fol_column_config[col] = st.column_config.Column(width="medium")
        elif col in ["Gate_Captain", "NPS_DSQ", "Dist_oben", "To_beat", "Codex", "Date", "Time", "Gap", "Sheet", "Row"]:
            fol_column_config[col] = st.column_config.Column(width="small")
        else:
            fol_column_config[col] = st.column_config.Column(width=60)
    # Explicitly configure BIB as NumberColumn to prevent Streamlit from
    # auto-inferring it as boolean (which shows a green dot when value is 1).
    fol_column_config["BIB"] = st.column_config.NumberColumn(width=60, format="%g")

    # Custom column order for workspace_fol: Codex, Gap, Time, Date placed right after BIB
    fol_column_order = [
        "DG", "Typ", "BIB", "Codex", "Gap", "Time", "Date",
        "Name", "Vorname", "NSA", "km/h", "Dist_m", "Dpts",
        "A", "B", "C", "D", "E", "Jpts", "Gate", "Gpts", "Wind", "Wpts", "Rpts",
        "Total", "To_beat", "Dist_oben", "Rk_LDG", "NPS_DSQ", "Gate_Captain", "MSG",
        "Sheet", "Row"
    ]

    edited_fol = st.data_editor(
        df_fol,
        num_rows="fixed",
        use_container_width=False,
        height=400,
        row_height=24,
        column_config=fol_column_config,
        column_order=fol_column_order,
        key=f"fol_editor_{st.session_state.editor_version}",
    )

    # Sync edited fol data to fol_calcdata in calc_headers order (replace NaN with None for downstream compatibility)
    st.session_state.fol_calcdata = (
        edited_fol[calc_headers].astype(object).where(edited_fol[calc_headers].notna(), None).values.tolist()
    )

    # ── 4. WORKSPACE MSG ──
    st.markdown(
        '<div style="margin-top: 20px; font-size: 0.78rem; font-family: Century Gothic, CenturyGothic, sans-serif; color: #CCCCCC;">workspace_msg (Daten: msg_cleandata)</div>',
        unsafe_allow_html=True,
    )

    df_msg = pd.DataFrame(
        st.session_state.msg_cleandata,
        columns=clean_msg_headers,
    )

    # Configure custom column widths to make workspace_msg grid compact
    msg_column_config = {}
    for col in clean_msg_headers:
        if col in ["Message"]:
            msg_column_config[col] = st.column_config.Column(width="large")
        elif col in ["Status"]:
            msg_column_config[col] = st.column_config.Column(width="medium")
        elif col in ["Date", "Time", "Codex"]:
            msg_column_config[col] = st.column_config.Column(width="small")
        else:
            msg_column_config[col] = st.column_config.Column(width=60)

    edited_msg = st.data_editor(
        df_msg,
        num_rows="fixed",
        use_container_width=True,
        height=200,
        row_height=24,
        column_config=msg_column_config,
        key=f"msg_editor_{st.session_state.editor_version}",
    )

    # Sync edited msg data back
    st.session_state.msg_cleandata = (
        edited_msg.astype(object).where(edited_msg.notna(), None).values.tolist()
    )

    # ── 5. SAVE VALIDATION DIALOG ──
    if st.session_state.get("save_validation_errors"):
        _show_save_validation_dialog()

    # ── 6. FILL WORKSPACE DG (after fol sync so data is always fresh) ──
    with dg_placeholder:
        st.markdown(
            '<div style="margin-top: 10px; font-size: 0.78rem; font-family: Century Gothic, CenturyGothic, sans-serif; color: #CCCCCC;">workspace_DG</div>',
            unsafe_allow_html=True,
        )

        dg_summary_data = _build_dg_summary(st.session_state.fol_calcdata)
        df_dg = pd.DataFrame(dg_summary_data, columns=dg_summary_headers)

        dg_column_config = {}
        for col in dg_summary_headers:
            if col in ["Codex", "Date", "Begin", "End"]:
                dg_column_config[col] = st.column_config.Column(width="small")
            elif col in ["Duration_min", "Gap_min"]:
                dg_column_config[col] = st.column_config.NumberColumn(width="small", format="%.2f")
            else:
                dg_column_config[col] = st.column_config.Column(width=60)

        dg_column_order = [
            "DG", "Typ", "Codex", "Jumps", "BIB_first", "BIB_last", "Jpts",
            "Date", "Begin", "End", "Duration_min", "Gap_min",
        ]

        st.data_editor(
            df_dg,
            num_rows="fixed",
            use_container_width=True,
            height=200,
            row_height=24,
            column_config=dg_column_config,
            column_order=dg_column_order,
            disabled=True,
            key=f"dg_editor_{st.session_state.editor_version}",
        )

        # Reference IDs were already validated before entering this stage.
        _, col_dg_save = st.columns([5, 2.5])
        with col_dg_save:
            if st.button(
                "save to xlsx",
                key="dg_save_xlsx_btn",
                type="secondary",
                use_container_width=True,
            ):
                validation_errors = _validate_dg_summary(dg_summary_data)
                if validation_errors:
                    st.session_state.save_validation_errors = validation_errors
                    st.rerun()
                else:
                    _normalize_types_for_save(st.session_state.fol_calcdata)
                    _save_weekend_xlsx()


def _save_weekend_xlsx():
    """Writes the validated editing data to the Weekend workbook."""
    with st.spinner("Speichere in Excel-Datei..."):
        try:
            from src.pipeline.phase_20_preprocessing.format_save import format_and_save

            format_and_save(
                st.session_state.fol_rawdata,
                st.session_state.msg_rawdata,
                st.session_state.msg_cleandata,
                st.session_state.fol_repairdata,
                st.session_state.fol_cleandata,
                st.session_state.output_filename,
                st.session_state.output_path,
                st.session_state.fol_calcdata,
                st.session_state.current_schanzen_data,
                st.session_state.list_Springer_ID_actual,
                pdf_output_path=st.session_state.path_Data_PDF,
            )
            st.success(
                f"Erfolgreich gespeichert: {st.session_state.output_filename}"
            )

            _cleanup_editing_state()
            st.rerun()
        except Exception as e:
            st.error(
                _format_save_error(e, st.session_state.output_path)
            )


def _format_save_error(error, output_path):
    """Returns a helpful save error while preserving unexpected details."""
    if isinstance(error, PermissionError):
        return (
            "Die Excel-Datei konnte nicht gespeichert werden. "
            "Möglicherweise ist sie noch in Excel geöffnet. "
            "Bitte schließen Sie die Datei und versuchen Sie es erneut.\n\n"
            f"Datei: {output_path}"
        )

    return f"Fehler beim Speichern: {error}"


def _cleanup_editing_state():
    """Cleans up all editing stage session state keys."""
    for key in [
        "show_editing_stage", "fol_rawdata", "msg_rawdata", "fol_repairdata",
        "fol_cleandata", "fol_calcdata", "msg_cleandata", "output_filename", "output_path",
        "filename_weekend", "fol_cleandata33column", "editor_version",
        "save_validation_errors",
        "current_schanzen_data", "list_Springer_ID_actual", "path_Data_PDF",
    ]:
        if key in st.session_state:
            del st.session_state[key]
    # Clean versioned data_editor widget keys
    for key in list(st.session_state.keys()):
        if key.startswith("fol_editor_") or key.startswith("msg_editor_") or key.startswith("dg_editor_"):
            del st.session_state[key]


@st.dialog("Validation before saving", width="medium")
def _show_save_validation_dialog():
    """Shows all DG/Typ validation errors and returns to manual editing."""
    errors = st.session_state.get("save_validation_errors", [])
    error_items = "".join(f"<li>{escape(error)}</li>" for error in errors)
    st.markdown(
        """
        <style>
        .save-validation-text,
        .save-validation-text * {
            color: #111111 !important;
        }
        .save-validation-text {
            font-family: 'Century Gothic', 'CenturyGothic', sans-serif;
            font-size: 1rem;
            line-height: 1.5;
        }
        .save-validation-text ul {
            margin-top: 0.75rem;
            padding-left: 1.4rem;
        }
        div[data-testid="stDialog"] h2 {
            color: #111111 !important;
        }
        </style>
        """
        f'<div class="save-validation-text">'
        "<strong>Die Daten können noch nicht gespeichert werden.</strong>"
        f"<ul>{error_items}</ul>"
        "</div>",
        unsafe_allow_html=True,
    )

    if st.button(
        "I will correct",
        key="save_validation_correct_btn",
        type="primary",
        use_container_width=True,
    ):
        st.session_state.pop("save_validation_errors", None)
        st.rerun()


