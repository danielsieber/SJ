"""
Parses weekend folder filenames into structured components.
"""


def parse_weekend_filename(name):
    """
    Parses a filename (without extension) according to specification:
    Example: '2026-01-24 SFWCM Oberstdorf HS235'
    - weekend_date: the first 10 characters (yyyy-mm-dd)
    - weekend_Cup: text up to the next space, minus the last character
    - weekend_gender: the last character of the first abbreviation
    - weekend_place: text between Cup and HS
    - weekend_HS: the number after HS
    """
    try:
        if len(name) < 11:
            raise ValueError("Name zu kurz für Standard-Parsing")
            
        weekend_date = name[0:10]
        rest = name[11:] # z.B. "SFWCM Oberstdorf HS235"
        
        parts = rest.split(' ', 1)
        cup_gender_token = parts[0] # "SFWCM"
        rest_after_cup = parts[1] # "Oberstdorf HS235"
        
        # Trenne Cup und Gender
        weekend_Cup = cup_gender_token[:-1] # "SFWC"
        weekend_gender = cup_gender_token[-1] # "M"
        
        if "HS" in rest_after_cup:
            place_part, hs_part = rest_after_cup.split("HS", 1)
            weekend_place = place_part.strip()
            weekend_HS = hs_part.strip()
        else:
            weekend_place = rest_after_cup.strip()
            weekend_HS = ""
            
        return {
            "weekend_date": weekend_date,
            "weekend_Cup": weekend_Cup,
            "weekend_gender": weekend_gender,
            "weekend_place": weekend_place,
            "weekend_HS": weekend_HS
        }
    except Exception:
        # Fallback bei unerwartetem Format
        return {
            "weekend_date": name[0:10] if len(name) >= 10 else "",
            "weekend_Cup": "",
            "weekend_gender": "A", # Standardmäßig 'A' (All)
            "weekend_place": name[11:] if len(name) > 11 else name,
            "weekend_HS": ""
        }
