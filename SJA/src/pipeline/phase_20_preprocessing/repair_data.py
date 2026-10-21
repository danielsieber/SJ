import datetime
import re


def is_numeric(val):
    if val is None or val == "":
        return False
    if isinstance(val, (int, float)):
        return True
    try:
        float(str(val))
        return True
    except ValueError:
        return False

def to_numeric(val):
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return val
    try:
        f = float(str(val))
        if f.is_integer():
            return int(f)
        return f
    except ValueError:
        return val

def get_safe(lst, index, default=None):
    if 0 <= index < len(lst):
        return lst[index]
    return default

def is_three_letters(val):
    if not val:
        return False
    val_str = str(val).strip()
    return len(val_str) == 3 and val_str.isalpha() and val_str.isupper()

def repair_data(fol_rawdata):
    fol_repairdata = []
    
    for row in fol_rawdata:
        # Check column 4 (index 3 in Python) for "Time" to skip header/empty rows
        checkTime = str(get_safe(row, 3, ""))
        if "time" in checkTime.lower():
            continue
            
        repaired_row = [None] * 28
        
        # Codex, Sheet, Row
        repaired_row[0] = get_safe(row, 0)
        repaired_row[1] = get_safe(row, 1)
        repaired_row[2] = get_safe(row, 2)
        
        # Date & Time parsing
        rawTS = str(get_safe(row, 3, ""))
        if len(rawTS) >= 19:
            sDate = rawTS[:10]
            try:
                repaired_row[3] = datetime.date.fromisoformat(sDate)
            except ValueError:
                repaired_row[3] = sDate
            
            pT = rawTS.find("T")
            if pT != -1 and len(rawTS) >= pT + 9:
                sTime = rawTS[pT+1:pT+9]
                try:
                    repaired_row[4] = datetime.time.fromisoformat(sTime)
                except ValueError:
                    repaired_row[4] = sTime
        else:
            repaired_row[3] = rawTS
            
        # Check team competition Rk in column 10 (index 9)
        col10_val = str(get_safe(row, 9, ""))
        if col10_val == "Rk":
            # Normal, NSA in column 15 or 19
            val19 = str(get_safe(row, 18, ""))
            if is_three_letters(val19):
                ankerNSA = 19
            else:
                ankerNSA = 15
        else:
            # Without "Rk" (Team competition), NSA in column 14 or 18
            val18 = str(get_safe(row, 17, ""))
            if is_three_letters(val18):
                ankerNSA = 18
            else:
                ankerNSA = 14
                
        # Fill BIB, Name, Vorname, NSA
        repaired_row[6] = to_numeric(get_safe(row, ankerNSA - 4))  # BIB (index 6)
        repaired_row[7] = get_safe(row, ankerNSA - 3)              # Name (index 7)
        repaired_row[8] = get_safe(row, ankerNSA - 2)              # Vorname (index 8)
        repaired_row[9] = get_safe(row, ankerNSA - 1)              # NSA (index 9)
        
        # Find Speed
        aSpeed = 0
        for c in range(4, 57):
            val = get_safe(row, c - 1)
            if val is not None and str(val) == "Speed":
                aSpeed = c
                break
                
        # Find To beat
        aToBeat = 0
        if aSpeed > 0:
            for c in range(aSpeed - 1, aSpeed - 3, -1):
                if c > 3:
                    val = get_safe(row, c - 1)
                    if val is not None and str(val) == "To beat":
                        aToBeat = c
                        break
                        
        # Collect numeric values between ankerNSA and aSpeed
        collectedNums = []
        if aSpeed > ankerNSA:
            for c in range(ankerNSA + 1, aSpeed):
                val = get_safe(row, c - 1)
                if val is not None and val != "" and is_numeric(val):
                    collectedNums.append(to_numeric(val))
                    
        numCount = len(collectedNums)
        countZahlen = numCount
        
        # Check and extract To beat value
        if aToBeat > 0:
            for c in range(aToBeat + 1, aSpeed):
                val = get_safe(row, c - 1)
                if val is not None and val != "" and is_numeric(val):
                    repaired_row[13] = to_numeric(val)  # To_beat (index 13)
                    if numCount > 0:
                        numCount -= 1
                    countZahlen = numCount
                    break
                    
        # Assign Dist_oben, Total_pts, Rk_LDG
        if countZahlen == 1:
            repaired_row[10] = collectedNums[0]
        elif countZahlen == 2:
            repaired_row[10] = collectedNums[0]
            repaired_row[11] = collectedNums[1]
        elif countZahlen >= 3:
            repaired_row[10] = collectedNums[0]
            repaired_row[11] = collectedNums[1]
            repaired_row[12] = collectedNums[2]
            
        # Find Slash "/"
        aSlash = 0
        startIdx = 43
        if startIdx > len(row):
            startIdx = len(row)
        for c in range(startIdx, 3, -1):
            val = get_safe(row, c - 1)
            if val is not None and "/" in str(val):
                aSlash = c
                break
                
        if aSlash > 0:
            for k in range(3, 0, -1):
                target_c = aSlash - k
                if target_c > 3:
                    checkVal = get_safe(row, target_c - 1)
                    if checkVal is not None and checkVal != "" and is_numeric(checkVal):
                        dVal = float(checkVal)
                        if k == 3 and dVal > 60:
                            repaired_row[14] = to_numeric(dVal)      # km/h  (Speed)
                        elif k == 2 and -10 <= dVal <= 10:
                            repaired_row[15] = to_numeric(dVal)      # Wind
                        elif k == 1 and dVal.is_integer() and 0 <= dVal <= 50:
                            repaired_row[17] = int(dVal)             # Gate

            # Parse slash contents (Wind_pts / Gate_pts)
            slash_val = str(get_safe(row, aSlash - 1, ""))
            parts = slash_val.split("/")
            if len(parts) >= 1:
                p0 = parts[0].strip()
                if is_numeric(p0):
                    repaired_row[16] = to_numeric(p0)  # Wind_pts (index 16)
            if len(parts) >= 2:
                p1 = parts[1].strip()
                if is_numeric(p1):
                    repaired_row[18] = to_numeric(p1)  # Gate_pts (index 18)
                    
            # Cells after the slash (Dist_m, Dist_pts)
            val_plus1 = get_safe(row, aSlash)
            if val_plus1 is not None and is_numeric(val_plus1):
                repaired_row[19] = to_numeric(val_plus1)  # Dist_m (index 19)
            val_plus2 = get_safe(row, aSlash + 1)
            if val_plus2 is not None and is_numeric(val_plus2):
                repaired_row[20] = to_numeric(val_plus2)  # Dist_pts (index 20)
                
            # Locate "Rd" starting from column 53 down to aSlash
            aRd = 0
            startIdx_rd = 53
            if startIdx_rd > len(row):
                startIdx_rd = len(row)
            for c in range(startIdx_rd, aSlash - 1, -1):
                val = get_safe(row, c - 1)
                if val is not None and "Rd" in str(val):
                    aRd = c
                    break
                    
            if aRd > 0:
                judgesNums = []
                for c in range(aRd + 1, len(row) + 1):
                    val = get_safe(row, c - 1)
                    if val is not None and val != "" and is_numeric(val):
                        judgesNums.append(to_numeric(val))
                jCount = len(judgesNums)
                
                if jCount == 1:
                    repaired_row[27] = judgesNums[0]  # Rnd_pts
                elif jCount > 1:
                    targetCol = 22
                    for val in judgesNums:
                        if targetCol <= 28:
                            repaired_row[targetCol - 1] = val
                            targetCol += 1
                    # Shift Judge_1 to Rnd_pts if Judge_2 is empty and Total_pts matches Judge_1
                    if repaired_row[22] is None or repaired_row[22] == "":
                        val11 = repaired_row[11]
                        val21 = repaired_row[21]
                        if val11 is not None and val21 is not None and float(val11) == float(val21):
                            repaired_row[27] = repaired_row[21]
                            repaired_row[21] = None
                            
        fol_repairdata.append(repaired_row)
        
    return fol_repairdata
