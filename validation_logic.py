import re
import pandas as pd
from collections import defaultdict

EXCEL_FIELDS = [
    "DATE TIME PRE KO (UTC)", "KO (UTC)", "REGION", "SPORT", "PROPERTY",
    "FIXTURE", "BROADCAST TIER", "SUPPORT TIER", "TX TYPE", "OVERRIDE ID",
    "MFL ID", "HEVC", "CLOSED CAPTIONS", "MULTI-TRACK AUDIO", "AUDIO LANG",
    "OTHER INFO (MULTIVIEW)"
]

HDR_TX_TYPES = ['DAI59 MR 1080p HDR', 'DAI59 1080p HDR']
SDR_TX_TYPES = ['DAI59 1080p', 'DAI59 MR 1080p', 'TX59 1080p']

def extract_numeric(value):
    match = re.search(r'(\d+)', str(value))
    return match.group(1) if match else None

def in_any_range(value, ranges):
    try:
        val = int(str(value).strip())
        return any(start <= val <= end for start, end in ranges)
    except Exception:
        return False

def match_sdr_override_id(value):
    try:
        val = int(value)
    except (TypeError, ValueError):
        return False

    return (
        1601 <= val <= 1660 or
        1681 <= val <= 1690 or
        2641 <= val <= 2660 or
        4601 <= val <= 4654 or
        (1000 <= val <= 9999 and (val // 100) % 10 == 5)  # x5xx pattern
    )

def match_hdr_override_id(value):
    try:
        val = int(str(value).strip())
        return (
            1601 <= val <= 1660 or
            1681 <= val <= 1690 or
            2641 <= val <= 2660 or
            4601 <= val <= 4654
        )
    except Exception:
        return False

def safe_str(val):
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    return str(val)

def is_sdr_override_id(value):
    return match_sdr_override_id(value)

def is_hdr_override_id(value):
    return match_hdr_override_id(value)

def get_dynamic_csv_bases_and_suffixes(df):
    """Dynamically get all base names and suffixes from DataFrame columns."""
    base_to_cols = defaultdict(list)
    seen_suffixes = set()
    for col in df.columns:
        if '_' in col:
            base, suffix = col.rsplit('_', 1)
            base_to_cols[base].append(col)
            seen_suffixes.add(suffix)
    return base_to_cols, seen_suffixes

def find_duplicate_rows(df):
    duplicated = df.duplicated(keep='first')
    return set(df[duplicated].index)

def build_csv_inconsistent_cells(df, base_to_cols):
    """Returns a dict of (row_idx, col): 'csvunmatch' or 'csvmatch' for inconsistent/matching CSV cells."""
    csv_inconsistent_cells = {}
    for row_idx, row in df.iterrows():
        for base, cols in base_to_cols.items():
            values = [safe_str(row[col]).strip() for col in cols]
            unique_vals = set(values)
            if len(unique_vals) > 1:
                for col in cols:
                    for v in unique_vals:
                        if values[cols.index(col)] != v:
                            csv_inconsistent_cells[(row_idx, col)] = 'csvunmatch'
            elif len(unique_vals) == 1 and list(unique_vals)[0] != "":
                for col in cols:
                    csv_inconsistent_cells[(row_idx, col)] = 'csvmatch'
    return csv_inconsistent_cells

def validate_cell(
    col,
    val,
    row,
    row_idx,
    csv_inconsistent_cells=None,
    duplicate_rows=None,
    dynamic_suffixes=None
):
    # Provide defaults for old calls
    if csv_inconsistent_cells is None:
        csv_inconsistent_cells = {}
    if duplicate_rows is None:
        duplicate_rows = set()
    if dynamic_suffixes is None:
        dynamic_suffixes = set()

    tx_type = safe_str(row.get('TX TYPE', '')).strip()
    override_id = safe_str(row.get('OVERRIDE ID', '')).strip()
    hevc = safe_str(row.get('HEVC', '')).strip()
    cc = safe_str(row.get('CLOSED CAPTIONS', '')).strip()
    mta = safe_str(row.get('MULTI-TRACK AUDIO', '')).strip()
    alang = safe_str(row.get('AUDIO LANG', '')).strip()
    broadcast_tier = safe_str(row.get('BROADCAST TIER', '')).strip()
    mfl_id = safe_str(row.get('MFL ID', '')).strip()
    pre_ko = safe_str(row.get('DATE TIME PRE KO (UTC)', '')).strip()

    is_hdr = tx_type in HDR_TX_TYPES
    is_sdr = tx_type in SDR_TX_TYPES

    val = safe_str(val)

    # CSV consistency validation
    if (row_idx, col) in csv_inconsistent_cells:
        if csv_inconsistent_cells[(row_idx, col)] == 'csvunmatch':
            return 'csvred'
        elif csv_inconsistent_cells[(row_idx, col)] == 'csvmatch':
            return 'csvgreen'

    # Duplicate row (top priority)
    if row_idx in duplicate_rows:
        return "duplicate"

    # SDR OVERRIDE ID with HDR TX TYPE: fail both fields
    sdr_override = is_sdr_override_id(override_id)
    if (col == 'TX TYPE' or col == 'OVERRIDE ID') and sdr_override and is_hdr:
        return False

    # HDR OVERRIDE ID with SDR TX TYPE: fail both fields
    hdr_override = is_hdr_override_id(override_id)
    if (col == 'TX TYPE' or col == 'OVERRIDE ID') and hdr_override and is_sdr:
        return False

    if col == 'TX TYPE':
        return bool(is_hdr or is_sdr)
    if col == 'OVERRIDE ID':
        if is_hdr:
            valid_ranges = [(1601,1660), (1681,1690), (2641,2660), (4601,4654)]
            return in_any_range(override_id, valid_ranges)
        elif is_sdr:
            return match_sdr_override_id(override_id)
        return False
    if col == 'HEVC':
        if is_hdr:
            return "hevc" in hevc.lower()
        elif is_sdr:
            return hevc.strip() == ""
        return False
    if col == 'CLOSED CAPTIONS':
        return cc in ['US English', 'US Spanish']
    if col == 'MULTI-TRACK AUDIO':
        return mta == 'No'
    if col == 'AUDIO LANG':
        if is_hdr:
            return "5.1" in alang
        elif is_sdr:
            return "5.1" not in alang
        return False

    # --- CSV V8 Fields and originalTier/tier logic ---
    # Loop through dynamically detected suffixes
    for suffix in dynamic_suffixes:
        suffix_str = f"_{suffix}"
        if col.endswith(suffix_str):
            base = col[:-(len(suffix)+1)]
            tx_type = safe_str(row.get('TX TYPE', '')).strip()
            broadcast_tier = safe_str(row.get('BROADCAST TIER', '')).strip()
            override_id = safe_str(row.get('OVERRIDE ID', '')).strip()
            mfl_id = safe_str(row.get('MFL ID', '')).strip()
            pre_ko = safe_str(row.get('DATE TIME PRE KO (UTC)', '')).strip()
            is_hdr = tx_type in HDR_TX_TYPES
            is_sdr = tx_type in SDR_TX_TYPES

            # MFL ID vs clientContentId
            if base == "clientContentId":
                client_content_id = safe_str(row.get(col, '')).strip()
                if mfl_id == "":
                    return "unvalidated"
                return (client_content_id == mfl_id)

            # DATE TIME PRE KO (UTC) vs day
            if base == "day":
                day_val = safe_str(row.get(col, '')).strip()
                if pre_ko == "" or day_val == "":
                    return "unvalidated"
                pre_ko_date = pre_ko[:10]
                return (day_val == pre_ko_date)

            # originalTier and tier must match numeric part from Broadcast Tier
            if base in ("originalTier", "tier"):
                expected_tier = extract_numeric(broadcast_tier)
                if expected_tier:
                    return (val == expected_tier)
                return False

            # heEventTypeName
            if base == "heEventTypeName":
                if is_hdr:
                    return "hevc_hdr10_5994" in val.lower()
                elif is_sdr:
                    return "avc_5994_freemium" in val.lower()
                return False
            # policies
            if base == "policies":
                val_lower = val.lower()
                if is_hdr:
                    return "captions708" in val_lower and ("dolby" in val_lower or "dolby5994" in val_lower)
                elif is_sdr:
                    return "captions708" in val_lower
                return False
            # drmRequired
            if base == "drmRequired":
                return val.lower() == "false"
            # performChannel
            if base == "performChannel":
                return val == override_id
            # variants
            if base == "variants":
                return "english single" in val.lower()
            # watermarking
            if base == "watermarking":
                return val == "NO_WATERMARKING"
            # heResilience
            if base == "heResilience":
                return val == "MAC"
            # Not a validated base
            return "unvalidated"

    # Unvalidated columns
    return "unvalidated"

def style_dataframe(df):
    duplicate_rows = find_duplicate_rows(df)
    base_to_cols, dynamic_suffixes = get_dynamic_csv_bases_and_suffixes(df)
    csv_inconsistent_cells = build_csv_inconsistent_cells(df, base_to_cols)

    def style_func(row):
        columns = list(row.index)
        style = []
        row_idx = row.name
        for i, col in enumerate(columns):
            val = row.get(col, "")
            valid = validate_cell(col, val, row, row_idx, csv_inconsistent_cells, duplicate_rows, dynamic_suffixes)
            if valid == "duplicate":
                style.append("background-color: #cce6ff")  # blue for duplicate row
            elif valid == "csvred":
                style.append("background-color: #b32400; color: #fff")  # dark red for csv mismatch
            elif valid == "csvgreen":
                style.append("background-color: #a5f5a6")  # different green for csv match
            elif valid is True:
                style.append("background-color: #d9f9d9")  # green
            elif valid is False:
                style.append("background-color: #b32400; color: #fff")  # dark red
            elif valid == "unvalidated":
                style.append("background-color: #fffbe6")  # light yellow
            else:
                style.append("")  # default
        return style
    return df.style.apply(style_func, axis=1)
