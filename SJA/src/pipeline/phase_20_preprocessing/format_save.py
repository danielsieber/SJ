import os
import re
from datetime import timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from src.pipeline.phase_20_preprocessing.fol_calc_excel import (
    collect_codex_download_keys,
    process_fol_calc_worksheet,
)
from src.pipeline.phase_20_preprocessing.jump_id_matching import (
    assign_msg_jump_ids,
    build_jump_blocks,
    normalize_codex,
    normalized_dg,
    parse_date_time,
)
from src.utils.fis_pdf import start_fis_pdf_downloads

FOL_CALC_DEFAULT_COLUMN_WIDTH = 7
FOL_CALC_COLUMN_WIDTHS = {
    "Name": 14,
    "Vorname": 14,
    "NPS_DSQ": 9,
    "Gate_Captain": 12,
    "MSG": 30,
    "Remark": 30,
    "Date": 11,
    "Time": 11,
}


def _excel_safe(value):
    """Convert pandas-style missing values to an empty Excel cell."""
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def _normalize_identity(value):
    """Normalize an athlete identity field for reliable ID generation."""
    value = _excel_safe(value)
    if value is None:
        return ""
    return " ".join(str(value).strip().casefold().split())


def _normalize_codex(value) -> str:
    """Normalize numeric and text Codex values for matching."""
    return normalize_codex(_excel_safe(value))


def _build_jumper_id(name, vorname, nsa):
    """Build the canonical Jumper_ID used by ``fol_calc`` and ``fol_wind``."""
    name, vorname, nsa = (
        _normalize_identity(value) for value in (name, vorname, nsa)
    )
    if not all((name, vorname, nsa)):
        return None
    return f"{nsa.upper()} {name.upper()[:5]} {vorname.upper()[:5]}"


def _parse_date_time(date_value, time_value):
    """Combine the date/time representations used by the FOL exports."""
    return parse_date_time(date_value, time_value)


def _normalized_dg(value):
    """Return a numeric DG from plain or visually decorated cell content."""
    return normalized_dg(value)


def _build_jump_blocks(fol_calcdata, filename_stem):
    """Build each DG's Jump_ID and first/last timestamp from fol_calcdata."""
    return build_jump_blocks(fol_calcdata, filename_stem)


def _assign_msg_jump_ids(
    msg_cleandata: list[list], jump_blocks: list[dict]
) -> list[str | None]:
    """Match msg_clean rows to the nearest Jump_ID within the same Codex."""
    return assign_msg_jump_ids(msg_cleandata, jump_blocks)


def _assign_wind_jump_id(repair_row, jump_blocks):
    """Assign a wind row using each DG's time span with ±1 minute tolerance."""
    row_datetime = _parse_date_time(
        repair_row[3] if len(repair_row) > 3 else None,
        repair_row[4] if len(repair_row) > 4 else None,
    )
    if row_datetime is None:
        return None

    candidates = []
    for position, block in enumerate(jump_blocks):
        start = block.get("start")
        end = block.get("end")
        if start is None or end is None:
            continue
        if start - timedelta(minutes=1) <= row_datetime <= end + timedelta(minutes=1):
            if row_datetime < start:
                distance = start - row_datetime
            elif row_datetime > end:
                distance = row_datetime - end
            else:
                distance = timedelta(0)
            candidates.append((distance, position, block["jump_id"]))

    if not candidates:
        return None
    return min(candidates, key=lambda candidate: (candidate[0], candidate[1]))[2]


def _open_saved_workbook(output_path):
    """Open the saved workbook with the Windows-associated Excel application."""
    if not hasattr(os, "startfile"):
        raise OSError("Automatic workbook opening is only supported on Windows.")
    os.startfile(str(Path(output_path).resolve()))


def format_and_save(
    fol_rawdata,
    msg_rawdata,
    msg_cleandata,
    fol_repairdata,
    fol_cleandata,
    output_filename,
    output_path,
    fol_calcdata=None,
    schanzen_data=None,
    springer_data=None,
    pdf_output_path=None,
):
    print(f"\n_28-01 Speichere in: {output_filename}")
    try:
        new_wb = Workbook()
        
        # Header für msg_raw generieren
        max_msg_cols = max(len(r) for r in msg_rawdata) if msg_rawdata else 5
        msg_headers = ["Codex", "Sheet", "row", "DateTime", "message"]
        for idx in range(len(msg_headers), max_msg_cols):
            col_letter = get_column_letter(idx - 3)
            msg_headers.append(f"col{col_letter}")
            
        # Header für fol_raw generieren
        max_fol_cols = max(len(r) for r in fol_rawdata) if fol_rawdata else 4
        fol_headers = ["Codex", "Sheet", "row", "DateTime"]
        for idx in range(len(fol_headers), max_fol_cols):
            col_letter = get_column_letter(idx - 3)
            fol_headers.append(f"col{col_letter}")
            
        ws_msg = new_wb.active
        ws_msg.title = "msg_raw"
        ws_msg.append(msg_headers)
        for row_data in msg_rawdata:
            ws_msg.append(row_data)
            
        # Bereinigtes Sheet msg_clean erstellen (nach msg_raw)
        clean_msg_headers = [
            "Date", "Time", "Message", "Jump", "Status", "BIB", "Gate",
            "Codex", "Sheet", "Row", "Jump_ID",
        ]
        ws_msg_clean = new_wb.create_sheet(title="msg_clean")
        ws_msg_clean.append(clean_msg_headers)
            
        ws_fol = new_wb.create_sheet(title="fol_raw")
        ws_fol.append(fol_headers)
        for row_data in fol_rawdata:
            ws_fol.append(row_data)
            
        # Repariertes Sheet fol_repair erstellen
        repair_headers = [
            "Codex", "Sheet", "Row", "Date", "Time", "Gap", "BIB", "Name", "Vorname", "NSA",
            "Dist_oben", "Total_pts", "Rk_LDG", "To_beat", "km/h", "Wind", "Wind_pts", "Gate", 
            "Gate_pts", "Dist_m", "Dist_pts", "Judge_1", "Judge_2", "Judge_3", "Judge_4", 
            "Judge_5", "Judge_pts", "Rnd_pts"
        ]
        ws_repair = new_wb.create_sheet(title="fol_repair")
        ws_repair.append(repair_headers)
        for row_data in fol_repairdata:
            ws_repair.append(row_data)
            
        # Bereinigtes Sheet fol_clean erstellen
        ws_fol_clean = new_wb.create_sheet(title="fol_clean")
        ws_fol_clean.append(repair_headers)
        for row_data in fol_cleandata:
            ws_fol_clean.append(row_data)
            
        # Berechnetes Sheet fol_calc erstellen
        if fol_calcdata is None:
            from src.pipeline.phase_20_preprocessing.calc_data import create_fol_calcdata
            fol_calcdata = create_fol_calcdata(fol_cleandata, msg_cleandata)

        filename_stem = Path(output_filename).stem
        original_jump_blocks = _build_jump_blocks(
            fol_calcdata, filename_stem
        )
        msg_jump_ids = _assign_msg_jump_ids(
            msg_cleandata, original_jump_blocks
        )

        # Saving must always include msg_clean assignments, even when the UI
        # calculation dialog was not opened (and for CLI exports).
        from src.pipeline.phase_20_preprocessing.list_calculate import (
            find_trialjumper_remark_rows,
            sync_messages_with_fol,
        )
        fol_calcdata, _ = sync_messages_with_fol(
            fol_calcdata, msg_cleandata
        )
        trialjumper_remark_rows = find_trialjumper_remark_rows(
            fol_calcdata, msg_cleandata
        )
            
        calc_headers = [
            "DG", "Typ", "BIB", "Name", "Vorname", "NSA", "km/h", "Dist_m", "Dpts",
            "A", "B", "C", "D", "E", "Jpts", "Gate", "Gpts", "Wind", "Wpts", "Rpts", 
            "Total", "To_beat", "Dist_oben", "Rk_LDG", "NPS_DSQ", "Gate_Captain", "MSG", "Remark",
            "Codex", "Date", "Time", "Gap", "Sheet", "Row", "Jumper_ID", "Jump_ID"
        ]

        excel_filename_stem = filename_stem.replace('"', '""')
        jump_blocks = _build_jump_blocks(fol_calcdata, filename_stem)
        for row_data, jump_id in zip(msg_cleandata, msg_jump_ids):
            ws_msg_clean.append(list(row_data) + [jump_id])

        # Prepare the Springer reference records once for both worksheets.
        springer_headers = []
        springer_records = []
        if springer_data is not None and len(springer_data) > 0:
            if hasattr(springer_data, "columns"):
                springer_headers = list(springer_data.columns)
                springer_records = springer_data.to_dict(orient="records")
            else:
                springer_records = list(springer_data)
                springer_headers = list(springer_records[0].keys())

        def convert_html_formula_to_excel(val):
            if isinstance(val, str):
                if val.startswith("="):
                    def shift_col(match):
                        col = match.group(1).upper()
                        row = match.group(2)
                        idx = 0
                        for char in col:
                            idx = idx * 26 + (ord(char) - ord('A') + 1)
                        idx = idx - 1
                        idx -= 1
                        new_col = ""
                        while idx >= 0:
                            remainder = idx % 26
                            new_col = chr(65 + remainder) + new_col
                            idx = (idx // 26) - 1
                        return new_col + row
                    return re.sub(r'\b([A-Z]+)([0-9]+)\b', shift_col, val)

                # Strip visual markers for Excel output so DG becomes a pure number
                cleaned = re.sub(r'^[🔵🔹🔷\s]+', '', val).strip()
                if cleaned.isdigit():
                    return int(cleaned)
                try:
                    return float(cleaned)
                except ValueError:
                    pass
            return val

        ws_fol_calc = new_wb.create_sheet(title="fol_calc")
        ws_fol_calc.append(calc_headers)
        calc_indexes = {
            header: index for index, header in enumerate(calc_headers)
        }
        for row_index, row_data in enumerate(fol_calcdata):
            excel_row = row_index + 2
            converted_row = [convert_html_formula_to_excel(cell) for cell in row_data]
            springer_id = _build_jumper_id(
                *(
                    converted_row[calc_indexes[header]]
                    if calc_indexes[header] < len(converted_row)
                    else None
                    for header in ("Name", "Vorname", "NSA")
                )
            )
            jump_id = (
                f'="{excel_filename_stem} "&TEXT(A{excel_row},"00")'
                f'&UPPER(TRIM(B{excel_row}))'
            )
            remark = (
                "trialjumper"
                if row_index in trialjumper_remark_rows
                else None
            )
            exported_row = converted_row[:27] + [remark] + converted_row[27:]
            ws_fol_calc.append(exported_row + [springer_id, jump_id])

        # Wind rows from fol_repair, enriched with the IDs used by fol_calc.
        ws_fol_wind = new_wb.create_sheet(title="fol_wind")
        ws_fol_wind.append(["Datum", "Zeit", "Wind", "Jumper_ID", "Jump_ID"])
        for repair_row in fol_repairdata or []:
            wind = _excel_safe(repair_row[15] if len(repair_row) > 15 else None)
            if wind is None or (isinstance(wind, str) and not wind.strip()):
                continue
            ws_fol_wind.append(
                [
                    _excel_safe(repair_row[3] if len(repair_row) > 3 else None),
                    _excel_safe(repair_row[4] if len(repair_row) > 4 else None),
                    wind,
                    _build_jumper_id(
                        repair_row[7] if len(repair_row) > 7 else None,
                        repair_row[8] if len(repair_row) > 8 else None,
                        repair_row[9] if len(repair_row) > 9 else None,
                    ),
                    _assign_wind_jump_id(repair_row, jump_blocks),
                ]
            )
            ws_fol_wind.cell(ws_fol_wind.max_row, 4).data_type = "s"
            ws_fol_wind.cell(ws_fol_wind.max_row, 5).data_type = "s"

        # Current Schanzen reference data, transposed as Header | Value.
        ws_schanzen = new_wb.create_sheet(title="SchanzenID")
        if schanzen_data is not None and len(schanzen_data) > 0:
            if hasattr(schanzen_data, "iloc"):
                schanzen_row = schanzen_data.iloc[0]
                schanzen_items = schanzen_row.items()
            elif isinstance(schanzen_data, list):
                first_row = schanzen_data[0] if schanzen_data else {}
                schanzen_items = first_row.items()
            else:
                schanzen_items = schanzen_data.items()
            for header, value in schanzen_items:
                ws_schanzen.append([header, _excel_safe(value)])

        # Unique actual Springer list enriched with Gender, DoB and JG.
        ws_springer = new_wb.create_sheet(title="SpringerID")
        if springer_records:
            ws_springer.append(springer_headers)
            for springer_row in springer_records:
                ws_springer.append(
                    [_excel_safe(springer_row.get(header)) for header in springer_headers]
                )

        # Complete calculated values and Excel-specific presentation immediately
        # before the workbook is frozen, sized, and saved.
        process_fol_calc_worksheet(ws_fol_calc, ws_schanzen)
        pdf_download_keys = collect_codex_download_keys(ws_fol_calc)

        all_sheets = [
            ws_msg,
            ws_msg_clean,
            ws_fol,
            ws_repair,
            ws_fol_clean,
            ws_fol_calc,
            ws_fol_wind,
            ws_schanzen,
            ws_springer,
        ]
            
        # Freeze panes und Auto-Breite für alle Sheets
        for ws in all_sheets:
            if ws.title != "SchanzenID":
                ws.freeze_panes = "E2" if ws is ws_fol_calc else "A2"
            
            for col_num, col_cells in enumerate(ws.columns, 1):
                max_length = 0
                column = get_column_letter(col_num)
                # Um Performance zu sparen, nur die ersten 1000 Zeilen prüfen
                for cell in col_cells[:1000]:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                
                adjusted_width = max_length + 2
                header = ws.cell(1, col_num).value
                normalized_header = str(header).strip()
                if ws is ws_fol_calc and normalized_header == "Jump_ID":
                    # Autofit to the calculated value rather than the much
                    # longer formula stored in each Jump_ID cell.
                    adjusted_width = len(Path(output_filename).stem) + 6
                if ws is ws_fol_calc and normalized_header in {
                    "Jumper_ID",
                    "Jump_ID",
                }:
                    ws.column_dimensions[column].width = adjusted_width
                elif ws is ws_fol_calc and normalized_header in FOL_CALC_COLUMN_WIDTHS:
                    ws.column_dimensions[column].width = FOL_CALC_COLUMN_WIDTHS[
                        normalized_header
                    ]
                elif ws is ws_fol_calc:
                    ws.column_dimensions[column].width = FOL_CALC_DEFAULT_COLUMN_WIDTH
                else:
                    ws.column_dimensions[column].width = adjusted_width

        # Store the workbook view so Excel opens on fol_calc with A2 selected.
        new_wb.active = new_wb.sheetnames.index("fol_calc")
        for ws in new_wb.worksheets:
            ws.sheet_view.tabSelected = ws is ws_fol_calc
        selection = ws_fol_calc.sheet_view.selection[0]
        selection.activeCell = "A2"
        selection.sqref = "A2"

        # Force Excel to evaluate formulas and formula-based conditional
        # formatting immediately when the newly generated file is opened.
        new_wb.calculation.calcMode = "auto"
        new_wb.calculation.fullCalcOnLoad = True
        new_wb.calculation.forceFullCalc = True
        new_wb.calculation.calcOnSave = True
        new_wb.calculation.calcCompleted = False
        new_wb.calculation.calcId = 0

        new_wb.save(output_path)
        _open_saved_workbook(output_path)
        # Excel is available immediately; PDF requests continue independently
        # and display a short native notification when all attempts finish.
        try:
            destination = (
                Path(pdf_output_path)
                if pdf_output_path is not None
                else Path(output_path).resolve().parent / "Data_pdf"
            )
            start_fis_pdf_downloads(pdf_download_keys, destination)
        except Exception as pdf_error:
            # PDF handling is optional and must not turn a successful workbook
            # export into an error.
            print(f"PDF-Download konnte nicht gestartet werden: {pdf_error}")
        print(f"Erfolgreich gespeichert unter {output_path}")
    except Exception as e:
        print(f"Fehler beim Speichern der Datei: {e}")
        raise
