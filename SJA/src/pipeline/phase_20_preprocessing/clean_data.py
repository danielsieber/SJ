import re
import datetime


def clean_msg_data(msg_rawdata):
    """
    Cleans message raw data according to rules:
    - Skip rows ending with 'seconds' in message.
    - Skip rows with identical message to the previous row (resetting per sheet/codex).
    - Map Jump, Status, BIB, Gate columns based on message patterns.
    - Split timestamp into separate Date and Time columns.
    """
    if not msg_rawdata:
        return []
        
    msg_cleandata = []
    
    last_codex = None
    last_sheet = None
    last_message = None
    
    # Pre-compile regexes for performance
    # Jump patterns
    p_training_xy = re.compile(r'Training\s+(\d+)\s*(?:-|–)?\s*(?:Jump|Round)\s+(\d+)', re.IGNORECASE)
    p_training_y = re.compile(r'Training\s+(?:Jump|Round)\s+(\d+)', re.IGNORECASE)
    p_trail = re.compile(r'tr(?:ai|ia)l\s+round', re.IGNORECASE)
    p_qualification = re.compile(r'(?<!dis)qualification', re.IGNORECASE) # Excludes disqualification
    p_round1 = re.compile(r'(1st|first)\s+round', re.IGNORECASE)
    p_round2 = re.compile(r'(2nd|second|final)\s+round', re.IGNORECASE) # Treats Final Round as second round
    p_round3 = re.compile(r'(3rd|third)\s+round', re.IGNORECASE) # Treats Third/3rd Round as W3
    
    # Status patterns
    p_nps = re.compile(r'(NPS|not\s+permitt?ed?)', re.IGNORECASE)
    p_dsq = re.compile(r'(DSQ|disqu?alifi?cation|disqualified)', re.IGNORECASE)
    p_trialjumper = re.compile(r'trial\s*jumper', re.IGNORECASE)
    
    # BIB pattern
    p_bib = re.compile(r'\bBIB\b(?:\s+No\.?)?\s*(\d+(?:-\d+)?)', re.IGNORECASE)
    
    # Gate patterns
    p_jury = re.compile(r'Jury', re.IGNORECASE)
    p_coach = re.compile(r'Coach', re.IGNORECASE)
    
    unique_rounds = {}
    next_t = 1
    
    for row in msg_rawdata:
        codex = row[0]
        sheet = row[1]
        row_num = row[2]
        dt = row[3]
        message = row[4]
        
        message_str = str(message).strip() if message is not None else ""
        
        # Reset last_message if sheet or codex changes
        if codex != last_codex or sheet != last_sheet:
            last_message = None
            last_codex = codex
            last_sheet = sheet
            
        # 1. Skip rows ending with 'seconds' in message
        if message_str.lower().endswith('seconds'):
            continue
            
        # 2. Skip rows with identical message to the previous row
        if message_str == last_message:
            continue
            
        last_message = message_str
        
        # Split timestamp into Date and Time
        date_val = None
        time_val = None
        if dt is not None:
            if isinstance(dt, datetime.datetime):
                date_val = dt.date()
                time_val = dt.time()
            elif isinstance(dt, datetime.date):
                date_val = dt
            elif isinstance(dt, datetime.time):
                time_val = dt
            else:
                rawTS = str(dt).strip()
                if len(rawTS) >= 19:
                    sDate = rawTS[:10]
                    try:
                        date_val = datetime.date.fromisoformat(sDate)
                    except ValueError:
                        date_val = sDate
                    
                    pT = rawTS.find("T")
                    if pT == -1:
                        pT = rawTS.find(" ")
                    if pT != -1 and len(rawTS) >= pT + 9:
                        sTime = rawTS[pT+1:pT+9]
                        try:
                            time_val = datetime.time.fromisoformat(sTime)
                        except ValueError:
                            time_val = sTime
                else:
                    date_val = rawTS
        
        # 3. Map Jump
        jump_val = None
        
        # Check direct short codes (exact or word match)
        m_code_old = re.search(r'\b(\d+)T-(\d+)\b', message_str, re.IGNORECASE)
        m_code_new = re.search(r'\bT(\d+)-(\d+)T\b', message_str, re.IGNORECASE)
        
        x_val, y_val = None, None
        
        if m_code_old:
            x_val = int(m_code_old.group(1))
            y_val = int(m_code_old.group(2))
        elif m_code_new:
            x_val = int(m_code_new.group(1))
            y_val = int(m_code_new.group(2))
        elif re.search(r'\bW1\b', message_str, re.IGNORECASE):
            jump_val = "W1"
        elif re.search(r'\bW2\b', message_str, re.IGNORECASE):
            jump_val = "W2"
        elif re.search(r'\bW3\b', message_str, re.IGNORECASE):
            jump_val = "W3"
        elif re.search(r'\bQ\b', message_str, re.IGNORECASE):
            jump_val = "Q"
        elif re.search(r'\b(trail|trial)\b', message_str, re.IGNORECASE) and not re.search(r'round', message_str, re.IGNORECASE) and not re.search(r'jumper', message_str, re.IGNORECASE):
            jump_val = "P"
            
        # Fall back to descriptive text checks
        if not jump_val and x_val is None:
            m_t_xy = p_training_xy.search(message_str)
            if m_t_xy:
                x_val = int(m_t_xy.group(1))
                y_val = int(m_t_xy.group(2))
            else:
                m_t_y = p_training_y.search(message_str)
                if m_t_y:
                    x_val = 1  # default training session
                    y_val = int(m_t_y.group(1))
                elif p_trail.search(message_str):
                    jump_val = "P"
                elif p_qualification.search(message_str):
                    jump_val = "Q"
                elif p_round1.search(message_str):
                    jump_val = "W1"
                elif p_round2.search(message_str):
                    jump_val = "W2"
                elif p_round3.search(message_str):
                    jump_val = "W3"
                    
        # If we identified a training round (x_val, y_val), assign its sequential T number
        if x_val is not None and y_val is not None:
            round_key = (x_val, y_val)
            if round_key not in unique_rounds:
                unique_rounds[round_key] = f"T{next_t}"
                next_t += 1
            jump_val = unique_rounds[round_key]
                
        # 4. Map Status
        status_val = None
        if p_nps.search(message_str):
            status_val = "NPS"
        elif p_dsq.search(message_str):
            status_val = "DSQ"
        elif p_trialjumper.search(message_str):
            status_val = "trialjumper"
            
        # 5. Map BIB
        bib_val = None
        m_bib = p_bib.search(message_str)
        if m_bib:
            bib_val_str = m_bib.group(1)
            if bib_val_str.isdigit():
                bib_val = int(bib_val_str)
            else:
                bib_val = bib_val_str
                
        # 6. Map Gate
        gate_val = None
        if p_jury.search(message_str):
            gate_val = "Jury"
        elif p_coach.search(message_str):
            gate_val = "Coach"
            
        # Build cleaned row with separate Date and Time
        # Schema: Date, Time, Message, Jump, Status, BIB, Gate, Codex, Sheet, Row
        cleaned_row = [date_val, time_val, message, jump_val, status_val, bib_val, gate_val, codex, sheet, row_num]
        msg_cleandata.append(cleaned_row)
        
    def msg_sort_key(row):
        d = row[0] # Date
        t = row[1] # Time
        
        if isinstance(d, datetime.date):
            d_str = d.isoformat()
        elif d is not None:
            d_str = str(d).strip()
        else:
            d_str = ""
            
        if isinstance(t, datetime.time):
            t_str = t.isoformat()
        elif t is not None:
            t_str = str(t).strip()
        else:
            t_str = ""
            
        return (d_str, t_str)
        
    msg_cleandata.sort(key=msg_sort_key)
    return msg_cleandata

def clean_fol_data(fol_repairdata):
    """
    Cleans follow repair data by:
    1. Sorting all entries by Date and Time.
    2. Calculating the gap (in minutes) from the previous entry.
    3. Grouping rows into jumps. A new jump starts if:
       - Codex changes
       - BIB changes
       - The gap to the previous entry is > 15 minutes.
    4. Merging the rows of each jump, starting from the last (newest) entry backwards,
       taking the first non-empty value for each column.
    5. Recalculating the gap between merged jumps.
    """
    if not fol_repairdata:
        return []

    # Helper to parse Date and Time into datetime.datetime for gap calculations
    def get_datetime_object(d, t):
        dt_date = None
        if isinstance(d, datetime.date):
            dt_date = d
        elif isinstance(d, str) and d.strip():
            try:
                dt_date = datetime.date.fromisoformat(d.strip())
            except ValueError:
                pass
                
        dt_time = None
        if isinstance(t, datetime.time):
            dt_time = t
        elif isinstance(t, str) and t.strip():
            try:
                dt_time = datetime.time.fromisoformat(t.strip()[:8])
            except ValueError:
                pass
                
        if dt_date and dt_time:
            return datetime.datetime.combine(dt_date, dt_time)
        elif dt_date:
            return datetime.datetime.combine(dt_date, datetime.time(0, 0))
        return None

    # Helper for sorting: converts date and time to ISO format strings for stable sorting
    def fol_sort_key(row):
        d = row[3]
        t = row[4]
        
        if isinstance(d, datetime.date):
            d_str = d.isoformat()
        elif d is not None:
            d_str = str(d).strip()
        else:
            d_str = ""
            
        if isinstance(t, datetime.time):
            t_str = t.isoformat()
        elif t is not None:
            t_str = str(t).strip()
        else:
            t_str = ""
            
        return (d_str, t_str)

    # 1. Sort fol_repairdata by date and time in-place
    fol_repairdata.sort(key=fol_sort_key)
    
    # 2. Calculate the gap (in seconds) to the previous row for all rows in fol_repairdata
    for i in range(len(fol_repairdata)):
        if i == 0:
            fol_repairdata[i][5] = None
        else:
            prev_dt = get_datetime_object(fol_repairdata[i-1][3], fol_repairdata[i-1][4])
            curr_dt = get_datetime_object(fol_repairdata[i][3], fol_repairdata[i][4])
            if prev_dt and curr_dt:
                diff = curr_dt - prev_dt
                gap_seconds = diff.total_seconds()
                fol_repairdata[i][5] = round(gap_seconds, 1)
            else:
                fol_repairdata[i][5] = None

    # 3. Group rows into jumps based on Codex, BIB, and Time Gap (> 15 minutes)
    jumps = []
    current_jump = []
    
    for row in fol_repairdata:
        if not current_jump:
            current_jump.append(row)
        else:
            prev_row = current_jump[-1]
            
            # Condition 1: Codex changes
            codex_changed = (prev_row[0] != row[0])
            
            # Condition 2: BIB changes
            bib_changed = (prev_row[6] != row[6])
            
            # Condition 3: Time gap to the last entry is > 15 minutes (900 seconds)
            gap = row[5]
            time_gap_exceeded = (gap is not None and gap > 900.0)
            
            if codex_changed or bib_changed or time_gap_exceeded:
                jumps.append(current_jump)
                current_jump = [row]
            else:
                current_jump.append(row)
                
    if current_jump:
        jumps.append(current_jump)
        
    # 4. Merge rows within each jump using 'loop_BIB' logic (searching backwards from newest entry)
    fol_cleandata = []
    num_cols = len(fol_repairdata[0]) if fol_repairdata else 28
    for jump_rows in jumps:
        # Determine the Time where km/h first appears
        first_kmh_time = None
        for r in jump_rows:
            kmh_val = r[14]
            time_val = r[4]
            if kmh_val is not None and str(kmh_val).strip() != "":
                if time_val is not None and str(time_val).strip() != "":
                    first_kmh_time = time_val
                    break

        best_row = [None] * num_cols
        for c in range(num_cols):
            if c == 4 and first_kmh_time is not None:
                best_row[c] = first_kmh_time
                continue
                
            # Loop backwards from the last (newest) entry
            for r in reversed(jump_rows):
                val = r[c]
                if val is not None and str(val).strip() != "":
                    best_row[c] = val
                    break
                    
        # Check if we already have a row with the same Codex, BIB, and Wind in fol_cleandata
        is_duplicate = False
        if best_row[6] is not None:
            for existing in fol_cleandata:
                if (existing[0] == best_row[0] and 
                    existing[6] == best_row[6] and 
                    existing[15] == best_row[15]):
                    is_duplicate = True
                    break
                    
        if not is_duplicate:
            fol_cleandata.append(best_row)
        
    # 5. Recalculate gaps in fol_cleandata for the cleaned sheet
    for i in range(len(fol_cleandata)):
        if i == 0:
            fol_cleandata[i][5] = None
        else:
            prev_dt = get_datetime_object(fol_cleandata[i-1][3], fol_cleandata[i-1][4])
            curr_dt = get_datetime_object(fol_cleandata[i][3], fol_cleandata[i][4])
            if prev_dt and curr_dt:
                diff = curr_dt - prev_dt
                gap_seconds = diff.total_seconds()
                fol_cleandata[i][5] = round(gap_seconds, 1)
            else:
                fol_cleandata[i][5] = None
                
    return fol_cleandata
