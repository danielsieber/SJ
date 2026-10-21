import re


def parse_filename(filename, folder_path):
    """Parses the filename and extracts specific information as a dictionary."""
    result = {
        "Weekend": None,
        "Filename": filename,
        "DateFile": None,
        "Place": None,
        "Land": None,
        "Cupraw": None,
        "Cup": None,
        "Gender": None,
        "HS": None,
        "Codex": None,
        "Saison": None
    }
    
    # 1. DateFile: Die ersten 6 digits
    if len(filename) >= 6 and filename[:6].isdigit():
        raw_date = filename[:6]
        # Format als YYYY-MM-DD (z.B. 250919 -> 2025-09-19)
        result["DateFile"] = f"20{raw_date[0:2]}-{raw_date[2:4]}-{raw_date[4:6]}"
        
        # Saison berechnen: Wettkämpfe ab Mai zählen zum nächsten Jahr
        year = int(f"20{raw_date[0:2]}")
        month = int(raw_date[2:4])
        if month >= 5:
            result["Saison"] = year + 1
        else:
            result["Saison"] = year
    else:
        return result
        
    # Rest des Dateinamens nach dem ersten Bindestrich z.B. "Val Di Fiemme (ITA) GP-Men's HS109-3024.xlsx"
    rest = filename[7:] 
    
    # Regex Pattern um die einzelnen Bestandteile sicher zu trennen
    # (.*) ist greedy für Cupraw, damit bei z.B. "WC,Training-QUA-Women's" das Cupraw = "WC,Training-QUA" wird
    pattern = r"^(.*?)\s*\((.*?)\)\s*(.*)-(.*?)\s*HS(\d+)-(\d+)\.xlsx$"
    match = re.match(pattern, rest)
    
    if match:
        result["Place"] = match.group(1).strip()
        result["Land"] = match.group(2).strip()
        result["Cupraw"] = match.group(3).strip()
        
        # Gender bereinigen (ein evtl. vorhandenes "'s" weglassen)
        gender = match.group(4).strip()
        if gender.endswith("'s"):
            gender = gender[:-2]
        result["Gender"] = gender
        
        result["HS"] = match.group(5).strip()
        result["Codex"] = match.group(6).strip()
        
        # Logik für 'Cup': Entspricht Cupraw, außer wenn Cupraw == 'QUA' oder (',' und 'QUA' enthält)
        cup = result["Cupraw"]
        if cup == "QUA" or ("," in cup and "QUA" in cup):
            try:
                # DateFile ist jetzt z.B. 2026-03-15, Monat steht an Index 5-7
                month = int(result["DateFile"][5:7])
                if 5 <= month <= 10:
                    cup = "GP"
                else:
                    cup = "WC"
            except ValueError:
                pass
        result["Cup"] = cup
        
    # Kein Lesen der Excel-Datei mehr, was das Skript extrem beschleunigt
        
    return result
