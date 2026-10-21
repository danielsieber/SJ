import os
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

from src.pipeline.phase_10_codex_list.parse_filename import parse_filename
from src.pipeline.phase_10_codex_list.group_weekends import group_and_sort_by_weekend


def get_scraper_files():
    """Reads files from a path and returns them as a list of dictionaries."""
    current_dir = Path(__file__).resolve().parent
    # Navigate from src/pipeline/phase_10_codex_list/ up to repo root, then into data/Data_FIS
    repo_root = current_dir.parent.parent.parent
    FOLDER_Data_FIS = str(repo_root / "data" / "Data_FIS")
        
    print()
    print(f"_10-01 Lese Dateien aus: {FOLDER_Data_FIS}")
    
    if not os.path.exists(FOLDER_Data_FIS):
        print(f"Hinweis: Der Ordner '{FOLDER_Data_FIS}' wurde nicht gefunden.")
        return []
        
    raw_files = [f for f in os.listdir(FOLDER_Data_FIS) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    # Liste mit den aufgeschlüsselten Datei-Infos erstellen
    file_list = [parse_filename(f, FOLDER_Data_FIS) for f in raw_files]
    
    # Gruppieren und Sortieren
    file_list = group_and_sort_by_weekend(file_list)
    
    return file_list


def export_file_list_to_excel(file_list, output_path):
    """Saves the list of dictionaries in a formatted Excel file."""
    if not file_list:
        print("Keine Dateien zum Speichern vorhanden.")
        return
        
    wb = Workbook()
    ws = wb.active
    ws.title = "Liste Codex"
    
    # Header extrahieren
    headers = list(file_list[0].keys())
    ws.append(headers)
    
    # Header formatieren (Fett, Gelber Hintergrund)
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        
    # Datenzeilen hinzufügen
    for row_data in file_list:
        ws.append([row_data.get(h) for h in headers])
        
    # Freeze Panes (Zeile 1 fixieren)
    ws.freeze_panes = "A2"
    
    # Autofilter setzen
    max_col = get_column_letter(len(headers))
    max_row = len(file_list) + 1
    ws.auto_filter.ref = f"A1:{max_col}{max_row}"
    
    center_alignment = Alignment(horizontal="center")
    
    # 1. Alle Spalten außer 'Filename' zentrieren
    filename_col_idx = headers.index("Filename") + 1
    for row in range(2, max_row + 1):
        for col in range(1, len(headers) + 1):
            if col != filename_col_idx:
                ws.cell(row=row, column=col).alignment = center_alignment
        
    # 2. Bedingte Formatierung (ungerade: hellgrau, gerade: grau)
    light_gray_fill = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
    gray_fill = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
    
    ws.conditional_formatting.add(
        f"A2:{max_col}{max_row}",
        FormulaRule(formula=['MOD($A2,2)=1'], fill=light_gray_fill)
    )
    ws.conditional_formatting.add(
        f"A2:{max_col}{max_row}",
        FormulaRule(formula=['MOD($A2,2)=0'], fill=gray_fill)
    )
    
    # Spaltenbreiten automatisch anpassen (Autofit)
    for col_num, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        column = get_column_letter(col_num)
        for cell in col_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        # Ein bisschen Platz hinzugeben
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width
        
    wb.save(output_path)
    print(f"\n_10-02 Excel-Datei erfolgreich gespeichert unter:\n{output_path}")
