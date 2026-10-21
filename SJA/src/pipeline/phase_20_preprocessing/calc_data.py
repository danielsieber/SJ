import datetime


def create_fol_calcdata(fol_cleandata, msg_cleandata=None):
    """
    Converts 28-column fol_cleandata into 33-column fol_calcdata.
    Adds DG and Typ columns at the start, and NPS_DSQ, Gate_Captain, MSG before Codex.
    Calculates DG (Durchgang) and Typ (Type) based on rules.
    """
    if not fol_cleandata:
        return []

    # Helper to parse BIB into prefix and suffix
    def parse_bib_parts(val):
        if val is None:
            return None, None
        val_str = str(val).strip()
        if '-' in val_str:
            parts = val_str.split('-')
            if len(parts) == 2:
                try:
                    return float(parts[0]), float(parts[1])
                except ValueError:
                    return parts[0], parts[1]
        try:
            return float(val_str), None
        except ValueError:
            return val_str, None

    def parse_compact_team_bib(val):
        """Parse a compact team BIB such as 124 as prefix 12, suffix 4."""
        if val is None:
            return None, None

        val_str = str(val).strip()
        if val_str.endswith(".0"):
            val_str = val_str[:-2]
        if not val_str.isdigit() or len(val_str) < 2:
            return parse_bib_parts(val)

        return float(val_str[:-1]), float(val_str[-1])

    def is_plain_integer_bib(val):
        """Return an integer for a plain numeric BIB, otherwise None."""
        if val is None:
            return None

        val_str = str(val).strip()
        if val_str.endswith(".0"):
            val_str = val_str[:-2]
        if not val_str.isdigit():
            return None

        return int(val_str)

    def starts_compact_team_sequence(start_idx):
        """Detect e.g. 11, 21, 31 as compact forms of 1-1, 2-1, 3-1."""
        if start_idx + 2 >= len(fol_cleandata):
            return False

        rows = fol_cleandata[start_idx:start_idx + 3]
        values = [is_plain_integer_bib(row[6]) for row in rows]
        if any(value is None or value < 10 or value % 10 != 1 for value in values):
            return False

        prefixes = [value // 10 for value in values]
        if not (prefixes[1] == prefixes[0] + 1 and prefixes[2] == prefixes[1] + 1):
            return False

        # Do not infer one numbering format across a definite DG boundary.
        first_codex = str(rows[0][0]).strip() if rows[0][0] is not None else ""
        first_date = str(rows[0][3]).strip() if rows[0][3] is not None else ""
        for row in rows[1:]:
            codex = str(row[0]).strip() if row[0] is not None else ""
            date = str(row[3]).strip() if row[3] is not None else ""
            if codex != first_codex or date != first_date:
                return False
            gap_val = row[5]
            if gap_val is not None and float(gap_val) > 900.0:
                return False

        return True

    # Helper to convert Date and Time to datetime object
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

    # Helper to map Jump to Type category (first letter of the jump value)
    def map_jump_to_type(jump):
        if not jump:
            return None
        jump_str = str(jump).strip()
        if jump_str:
            return jump_str[0].upper()
        return None

    # Build mapping of codex to list of jumps from msg_cleandata
    msg_jumps_by_codex = {}
    if msg_cleandata:
        for m_row in msg_cleandata:
            if len(m_row) >= 10:
                codex = str(m_row[7]).strip() if m_row[7] is not None else ""
                jump = m_row[3]
                m_date = m_row[0]
                m_time = m_row[1]
                if jump:
                    mapped_t = map_jump_to_type(jump)
                    if mapped_t:
                        m_dt = get_datetime_object(m_date, m_time)
                        msg_jumps_by_codex.setdefault(codex, []).append((m_dt, mapped_t))

    # Identify DG boundaries
    dg_boundaries = [0] # List of start indices for each DG
    consecutive_increases = 0
    compact_team_mode = starts_compact_team_sequence(0)
    
    for i in range(1, len(fol_cleandata)):
        row_i = fol_cleandata[i]
        row_prev = fol_cleandata[i-1]
        
        codex_changed = (str(row_i[0]).strip() if row_i[0] is not None else "") != (str(row_prev[0]).strip() if row_prev[0] is not None else "")
        date_changed = (str(row_i[3]).strip() if row_i[3] is not None else "") != (str(row_prev[3]).strip() if row_prev[3] is not None else "")
        
        gap_val = row_i[5]
        gap_exceeded = (gap_val is not None and float(gap_val) > 900.0)
        
        # Detect compact team numbering after a hard boundary as well. In this
        # format 11, 21, 31, ..., 124 mean 1-1, 2-1, 3-1, ..., 12-4.
        if codex_changed or date_changed or gap_exceeded:
            compact_team_mode = starts_compact_team_sequence(i)
        elif not compact_team_mode and starts_compact_team_sequence(i - 1):
            compact_team_mode = True

        # BIB logic including explicit (14-1) and compact (141) team suffixes.
        parser = parse_compact_team_bib if compact_team_mode else parse_bib_parts
        bib_curr_prefix, bib_curr_suffix = parser(row_i[6])
        bib_prev_prefix, bib_prev_suffix = parser(row_prev[6])
        
        is_bib_new_dg = False
        if bib_curr_prefix is not None and bib_prev_prefix is not None:
            has_team_suffixes = bib_curr_suffix is not None and bib_prev_suffix is not None

            # Team BIBs use the suffix to delimit DGs. A prefix reset such as
            # 17-1 -> 1-2 (or compact 171 -> 12) is the next team round inside
            # the same DG, while 17-4 -> 1-1 starts the next DG.
            if has_team_suffixes:
                if isinstance(bib_curr_suffix, (int, float)) and isinstance(bib_prev_suffix, (int, float)):
                    if bib_curr_suffix < bib_prev_suffix:
                        is_bib_new_dg = True
             
            # Apply the ordinary numeric-prefix rule only to non-team BIBs.
            # Otherwise a sufficiently long team round would be split when
            # its prefix restarts at 1 for the next suffix.
            if not has_team_suffixes and isinstance(bib_curr_prefix, (int, float)) and isinstance(bib_prev_prefix, (int, float)):
                if bib_curr_prefix > bib_prev_prefix:
                    consecutive_increases += 1
                elif bib_curr_prefix < bib_prev_prefix:
                    if consecutive_increases > 10:
                        is_bib_new_dg = True
                    consecutive_increases = 0
                else:
                    consecutive_increases = 0
            else:
                consecutive_increases = 0
        else:
            consecutive_increases = 0
            
        if codex_changed or date_changed or gap_exceeded or is_bib_new_dg:
            dg_boundaries.append(i)
            consecutive_increases = 0

    # Build DGs with start/end index
    dgs = []
    for idx, start_idx in enumerate(dg_boundaries):
        end_idx = dg_boundaries[idx+1] - 1 if idx + 1 < len(dg_boundaries) else len(fol_cleandata) - 1
        dgs.append((idx + 1, start_idx, end_idx))

    # Map each DG to its most likely type
    dg_types = {}
    for dg_num, start_idx, end_idx in dgs:
        dg_codex = str(fol_cleandata[start_idx][0]).strip() if fol_cleandata[start_idx][0] is not None else ""
        dg_start_dt = get_datetime_object(fol_cleandata[start_idx][3], fol_cleandata[start_idx][4])
        dg_end_dt = get_datetime_object(fol_cleandata[end_idx][3], fol_cleandata[end_idx][4])
        
        candidates = msg_jumps_by_codex.get(dg_codex, [])
        
        # Method 1: Count messages within time range of this DG (with tight 1-minute tolerance)
        tally = {}
        if dg_start_dt and dg_end_dt:
            margin = datetime.timedelta(minutes=1)
            for m_dt, m_type in candidates:
                if m_dt and (dg_start_dt - margin <= m_dt <= dg_end_dt + margin):
                    tally[m_type] = tally.get(m_type, 0) + 1
                    
        # Method 2 (Fallback): Count messages within time range with 5-minute tolerance
        if not tally and dg_start_dt and dg_end_dt:
            margin = datetime.timedelta(minutes=5)
            for m_dt, m_type in candidates:
                if m_dt and (dg_start_dt - margin <= m_dt <= dg_end_dt + margin):
                    tally[m_type] = tally.get(m_type, 0) + 1
                    
        # Method 3 (Fallback): Find closest message in time
        if not tally and dg_start_dt and candidates:
            closest_candidate = None
            min_diff = None
            for m_dt, m_type in candidates:
                if m_dt:
                    diff = abs(m_dt - dg_start_dt)
                    if min_diff is None or diff < min_diff:
                        min_diff = diff
                        closest_candidate = m_type
            if min_diff and min_diff <= datetime.timedelta(hours=2):
                tally[closest_candidate] = 1
                
        # Method 4 (Fallback): Most common type for the codex
        if not tally and candidates:
            for _, m_type in candidates:
                tally[m_type] = tally.get(m_type, 0) + 1
                
        # Check if Jpts has a value in this DG range
        has_jpts = False
        for i in range(start_idx, end_idx + 1):
            jpts_val = fol_cleandata[i][26]
            if jpts_val is not None and str(jpts_val).strip() != "":
                try:
                    if float(jpts_val) > 0:
                        has_jpts = True
                        break
                except ValueError:
                    pass

        # Determine best type
        best_type = None
        if tally:
            best_type = max(tally, key=tally.get)
            
        # Refine type based on Jpts presence
        if has_jpts:
            # Must be Q or W. If in doubt, use W
            if best_type not in ('Q', 'W'):
                best_type = 'W'
        else:
            # Must be T or P. If in doubt, use T
            if best_type not in ('T', 'P'):
                best_type = 'T'
            
        dg_types[dg_num] = best_type

    # Assemble fol_calcdata
    fol_calcdata = []
    current_dg_num = 1
    boundary_set = set(dg_boundaries)
    
    for i, row in enumerate(fol_cleandata):
        if len(row) < 28:
            row = list(row) + [None] * (28 - len(row))
            
        # Determine DG and Typ cell values/formulas (only note once at the top of every new DG block)
        if i in boundary_set:
            # First row of a DG
            current_dg_num = dg_boundaries.index(i) + 1
            dg_val = f"🔵 {current_dg_num}"
            typ_val = dg_types.get(current_dg_num, 'T')
        else:
            # Follow-up row: leave blank
            dg_val = None
            typ_val = None

        jpts_formula = row[26]


        new_row = [
            dg_val,    # DG (index 0)
            typ_val,   # Typ (index 1)
            row[6],    # BIB (index 2)
            row[7],    # Name (index 3)
            row[8],    # Vorname (index 4)
            row[9],    # NSA (index 5)
            row[14],   # km/h (index 6)
            row[19],   # Dist_m (index 7)
            row[20],   # Dpts (index 8)
            row[21],   # A (index 9)
            row[22],   # B (index 10)
            row[23],   # C (index 11)
            row[24],   # D (index 12)
            row[25],   # E (index 13)
            jpts_formula, # Jpts (index 14)
            row[17],   # Gate (index 15)
            row[18],   # Gpts (index 16)
            row[15],   # Wind (index 17)
            row[16],   # Wpts (index 18)
            row[27],   # Rpts (index 19)
            row[11],   # Total (index 20)
            row[13],   # To_beat (index 21)
            row[10],   # Dist_oben (index 22)
            row[12],   # Rk_LDG (index 23)
            None,      # NPS_DSQ (index 24)
            None,      # Gate_Captain (index 25)
            None,      # MSG (index 26)
            row[0],    # Codex (index 27)
            row[3],    # Date (index 28)
            row[4],    # Time (index 29)
            row[5],    # Gap (index 30)
            row[1],    # Sheet (index 31)
            row[2]     # Row (index 32)
        ]
        fol_calcdata.append(new_row)
        
    return fol_calcdata
