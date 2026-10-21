from datetime import datetime


def group_and_sort_by_weekend(file_list):
    """Groups files by Place, Cup, Gender, HS and Date (+/- 5 days) into Weekends."""
    valid = []
    invalid = []
    
    for f in file_list:
        if f["DateFile"] and f["Place"] and f["Cup"] and f["Gender"] and f["HS"]:
            try:
                dt = datetime.strptime(f["DateFile"], "%Y-%m-%d")
                f["_dt"] = dt
                valid.append(f)
            except ValueError:
                invalid.append(f)
        else:
            invalid.append(f)
            
    # Aufsteigend nach Datum sortieren, damit das älteste Weekend die Nummer 1 bekommt
    valid.sort(key=lambda x: x["_dt"])
    
    weekends = []
    weekend_counter = 1
    
    for f in valid:
        matched = False
        # Rückwärts durch die Wochenenden suchen, um das neueste passende zu finden
        for w in reversed(weekends):
            if (w["place"] == f["Place"] and 
                w["cup"] == f["Cup"] and 
                w["gender"] == f["Gender"] and 
                w["hs"] == f["HS"]):
                
                # Liegt das Datum innerhalb von 5 Tagen?
                if abs((f["_dt"] - w["last_date"]).days) <= 5:
                    w["last_date"] = f["_dt"]
                    f["Weekend"] = w["id"]
                    matched = True
                    break
                    
        # Neues Weekend anlegen, wenn nichts gepasst hat
        if not matched:
            weekends.append({
                "place": f["Place"],
                "cup": f["Cup"],
                "gender": f["Gender"],
                "hs": f["HS"],
                "last_date": f["_dt"],
                "id": weekend_counter
            })
            f["Weekend"] = weekend_counter
            weekend_counter += 1
            
    # Cleanup der temporären Datums-Spalte
    for f in valid:
        del f["_dt"]
        
    # Hilfsfunktion für die Sortierung
    def sort_key(x):
        # Codex als Zahl parsen
        try:
            codex_val = int(x.get("Codex", 0) or 0)
        except (ValueError, TypeError):
            codex_val = 0
            
        return (x["Weekend"], x["DateFile"], codex_val)

    # Sortieren nach Weekend (absteigend), DateFile (absteigend) und Codex (höchster zuerst)
    valid.sort(key=sort_key, reverse=True)
    
    return valid + invalid
