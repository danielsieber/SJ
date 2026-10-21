import os
import re
from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime

from src.pipeline.phase_10_codex_list.parse_filename import parse_filename


def main_analysis(weekend_id, root_dir):
    liste_codex_path = os.path.join(root_dir, "Liste_Codex.xlsx")
    
    if not os.path.exists(liste_codex_path):
        print(f"Fehler: {liste_codex_path} nicht gefunden.")
        return
        
    print(f"\n_20-02 Lese Liste: {liste_codex_path}")
    wb_list = load_workbook(filename=liste_codex_path, data_only=True, read_only=True)
    ws_list = wb_list.active
    
    # Header extrahieren
    headers = [cell.value for cell in next(ws_list.iter_rows(min_row=1, max_row=1))]
    
    listCodex = []
    for row in ws_list.iter_rows(min_row=2, values_only=True):
        if row[0] == weekend_id:  # Spalte A (Index 0) ist "Weekend"
            row_dict = {headers[i]: row[i] for i in range(len(headers))}
            listCodex.append(row_dict)
            
    wb_list.close()
    
    if not listCodex:
        print(f"Keine Dateien für Weekend {weekend_id} in der Liste gefunden.")
        return
        
    print(f"\nFolgende Dateien für Weekend {weekend_id} gefunden:")
    for item in listCodex:
        print(f"- {item.get('Filename')}")
        
    def sort_key_for_top(x):
        try:
            codex_val = int(x.get("Codex", 0) or 0)
        except (ValueError, TypeError):
            codex_val = 0
        return (x.get("DateFile") or "", codex_val)
        
    listCodex.sort(key=sort_key_for_top, reverse=True)
    
    # Metadaten aus dem obersten Eintrag lesen für den Dateinamen
    top_item = listCodex[0]
    date_file = top_item.get("DateFile", "")
    place = top_item.get("Place", "")
    cup = top_item.get("Cup", "")
    gender = top_item.get("Gender", "")
    hs = top_item.get("HS", "")
    
    # Gender Logik (Mixed -> A)
    if not gender:
        gender_letter = ""
    elif gender.lower() == "mixed":
        gender_letter = "A"
    else:
        gender_letter = gender[0].upper()
        
    output_filename = f"{date_file} {cup}{gender_letter} {place} HS{hs}.xlsx"
    output_path = os.path.join(root_dir, output_filename)
    
    data_fis_dir = os.path.join(root_dir, "Data_FIS")
    
    fol_rawdata = []
    msg_rawdata = []
    
    print("\n_20-03 Auslesen der einzelnen Excel-Dateien...")
    for item in listCodex:
        filename = item.get("Filename")
        if not filename:
            continue
            
        file_path = os.path.join(data_fis_dir, filename)
        if not os.path.exists(file_path):
            print(f"  Achtung: {filename} nicht im Ordner Data_FIS gefunden.")
            continue
            
        codex = item.get("Codex", "")
        print(f"  Lese: {filename} (Codex: {codex})")
        
        try:
            wb = load_workbook(filename=file_path, data_only=True, read_only=True)
            for sheet_name in wb.sheetnames:
                if sheet_name.startswith('fol_') or sheet_name.startswith('follow_') or sheet_name.startswith('msg_'):
                    sheet = wb[sheet_name]
                    is_fol = sheet_name.startswith('fol_') or sheet_name.startswith('follow_')
                    
                    for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                        if any(cell is not None for cell in row):
                            # Skip rows where timestamp says timestamp
                            if isinstance(row[0], str) and 'timestamp' in row[0].lower():
                                continue
                                
                            new_row = [codex, sheet_name, row_idx + 1] + list(row)
                            
                            if is_fol:
                                fol_rawdata.append(new_row)
                            else:
                                msg_rawdata.append(new_row)
            wb.close()
        except Exception as e:
            print(f"Fehler beim Lesen von {filename}: {e}")
            
    # Sortier-Hilfsfunktionen
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

    def sort_key(row):
        val = row[3]
        codex_val = str(row[0]) if row[0] is not None else ""
        sheetname_val = str(row[1]) if row[1] is not None else ""
        sheetname_key = natural_sort_key(sheetname_val)
        
        if val is None:
            return (codex_val, sheetname_key, datetime.min)
        if isinstance(val, datetime):
            return (codex_val, sheetname_key, val)
        if isinstance(val, str):
            try:
                return (codex_val, sheetname_key, datetime.fromisoformat(val))
            except ValueError:
                return (codex_val, sheetname_key, str(val))
        return (codex_val, sheetname_key, str(val))

    print("\n_20-04 Sortiere zusammengeführte Daten...")
    try:
        fol_rawdata.sort(key=sort_key)
        msg_rawdata.sort(key=sort_key)
    except TypeError:
        print("  Hinweis: Gemischte Datentypen in Timestamp gefunden. Sortiere als String.")
        fol_rawdata.sort(key=lambda row: (str(row[0]), natural_sort_key(str(row[1])), str(row[3]) if row[3] is not None else ""))
        msg_rawdata.sort(key=lambda row: (str(row[0]), natural_sort_key(str(row[1])), str(row[3]) if row[3] is not None else ""))

    return fol_rawdata, msg_rawdata, output_filename, output_path


def main_analysis_for_files(selected_filenames, data_fis_dir, output_dir):
    listCodex = []
    for fn in selected_filenames:
        item = parse_filename(fn, None)
        listCodex.append(item)
        
    if not listCodex:
        return None
        
    def sort_key_for_top(x):
        try:
            codex_val = int(x.get("Codex", 0) or 0)
        except (ValueError, TypeError):
            codex_val = 0
        return (x.get("DateFile") or "", codex_val)
        
    listCodex.sort(key=sort_key_for_top, reverse=True)
        
    top_item = listCodex[0]
    date_file = top_item.get("DateFile", "")
    place = top_item.get("Place", "")
    cup = top_item.get("Cup", "")
    gender = top_item.get("Gender", "")
    hs = top_item.get("HS", "")
    
    if not gender:
        gender_letter = ""
    elif gender.lower() == "mixed":
        gender_letter = "A"
    else:
        gender_letter = gender[0].upper()
        
    output_filename = f"{date_file} {cup}{gender_letter} {place} HS{hs}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    fol_rawdata = []
    msg_rawdata = []
    
    for item in listCodex:
        filename = item.get("Filename")
        if not filename:
            continue
            
        file_path = os.path.join(data_fis_dir, filename)
        if not os.path.exists(file_path):
            continue
            
        codex = item.get("Codex", "")
        
        try:
            wb = load_workbook(filename=file_path, data_only=True, read_only=True)
            for sheet_name in wb.sheetnames:
                if sheet_name.startswith('fol_') or sheet_name.startswith('follow_') or sheet_name.startswith('msg_'):
                    sheet = wb[sheet_name]
                    is_fol = sheet_name.startswith('fol_') or sheet_name.startswith('follow_')
                    
                    for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                        if any(cell is not None for cell in row):
                            if isinstance(row[0], str) and 'timestamp' in row[0].lower():
                                continue
                                
                            new_row = [codex, sheet_name, row_idx + 1] + list(row)
                            
                            if is_fol:
                                fol_rawdata.append(new_row)
                            else:
                                msg_rawdata.append(new_row)
            wb.close()
        except Exception as e:
            print(f"Fehler beim Lesen von {filename}: {e}")
            
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

    def sort_key(row):
        val = row[3]
        codex_val = str(row[0]) if row[0] is not None else ""
        sheetname_val = str(row[1]) if row[1] is not None else ""
        sheetname_key = natural_sort_key(sheetname_val)
        
        if val is None:
            return (codex_val, sheetname_key, datetime.min)
        if isinstance(val, datetime):
            return (codex_val, sheetname_key, val)
        if isinstance(val, str):
            try:
                return (codex_val, sheetname_key, datetime.fromisoformat(val))
            except ValueError:
                return (codex_val, sheetname_key, str(val))
        return (codex_val, sheetname_key, str(val))

    try:
        fol_rawdata.sort(key=sort_key)
        msg_rawdata.sort(key=sort_key)
    except TypeError:
        fol_rawdata.sort(key=lambda row: (str(row[0]), natural_sort_key(str(row[1])), str(row[3]) if row[3] is not None else ""))
        msg_rawdata.sort(key=lambda row: (str(row[0]), natural_sort_key(str(row[1])), str(row[3]) if row[3] is not None else ""))

    return fol_rawdata, msg_rawdata, output_filename, output_path
