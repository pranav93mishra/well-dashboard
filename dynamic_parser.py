"""
Dynamic parser for ONGC Well Card Excel files - V2 (Fixed).
Reads data from correct cell positions based on deep inspection.
"""
import os
import re
import json
import traceback
import pandas as pd
import numpy as np
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding='utf-8')

WELL_CARDS_DIR = os.path.join(os.path.dirname(__file__), "well_cards", "Well Cards")
OUTPUT_JSON = r"C:\Users\ongca\Downloads\well_dashboard\all_wells_data.json"

EXCLUDE_KEYWORDS = [
    'calculation', 'invoice', 'survey', 'cost estimate', 'costestimate',
    'bha', 'chemical consumption', 'sdpr', 'pressure points',
    'technical summary', 'fpt_rt', 'wellpath', 'rotary', 'recap',
    'call out', 'barite', 'chemical cost'
]


def safe_float(val, default=0.0):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return default
    try:
        s = str(val).strip().replace(',', '').replace('₹', '').replace('INR', '')
        # Extract numeric part
        m = re.search(r'[-+]?\d*\.?\d+', s)
        return float(m.group()) if m else default
    except:
        return default


def safe_str(val, default=""):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return default
    s = str(val).strip()
    return s if s and s != 'nan' else default


def cell(rows, r, c, default=""):
    """Safe cell access."""
    if r < len(rows) and c < len(rows[r]):
        v = rows[r][c]
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        return v
    return default


def find_well_card_files(base_dir):
    all_xlsx = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if f.endswith('.xlsx') and not f.startswith('~$'):
                all_xlsx.append(os.path.join(root, f))
    well_cards = []
    for f in all_xlsx:
        fname = os.path.basename(f).lower()
        if any(kw in fname for kw in EXCLUDE_KEYWORDS):
            continue
        well_cards.append(f)
    return well_cards


def get_asset_from_path(filepath):
    rel = os.path.relpath(filepath, WELL_CARDS_DIR)
    parts = rel.split(os.sep)
    folder_map = {
        'B & S (ST)': 'B&S Asset', 'BS (DEV)': 'B&S Asset',
        'DW': 'Deepwater', 'Exploratory': 'Exploratory',
        'MH': 'Mumbai High', 'NH': 'Neelam-Heera',
    }
    return folder_map.get(parts[0], parts[0]) if parts else "Unknown"


def get_sub_category_from_path(filepath):
    rel = os.path.relpath(filepath, WELL_CARDS_DIR)
    parts = rel.split(os.sep)
    if len(parts) >= 2:
        sub = parts[1].upper()
        if sub in ('ST',):
            return 'Side Track'
        elif sub in ('DEV',):
            return 'Development'
        elif sub in ('WO',):
            return 'Workover'
    return None


def parse_lat_lon(raw_str):
    if not raw_str:
        return 0.0
    raw_str = str(raw_str)
    # Normalize all degree/minute/second Unicode variants to ASCII
    raw_str = raw_str.replace('º', '°').replace('\u00b0', '°').replace('\u02da', '°')
    # Normalize single-quote variants: curly quotes, primes, backtick
    raw_str = raw_str.replace('\u2018', "'").replace('\u2019', "'")  # LEFT/RIGHT SINGLE QUOTATION
    raw_str = raw_str.replace('\u02b9', "'").replace('\u2032', "'")  # MODIFIER LETTER PRIME, PRIME
    raw_str = raw_str.replace('\u02bc', "'").replace('\u02ca', "'")  # MODIFIER LETTER APOSTROPHE
    raw_str = raw_str.replace("''", '"')
    # Normalize double-quote variants: curly double quotes, double prime
    raw_str = raw_str.replace('\u201c', '"').replace('\u201d', '"')  # LEFT/RIGHT DOUBLE QUOTATION
    raw_str = raw_str.replace('\u2033', '"')  # DOUBLE PRIME
    raw_str = raw_str.replace('\u02ba', '"')  # MODIFIER LETTER DOUBLE PRIME
    try:
        val = float(raw_str)
        if -180 <= val <= 180:
            return val
    except:
        pass
    pattern = r'(\d+)\s*[°]\s*(\d+)\s*[\'′ʹ]\s*([\d.]+)\s*[\"″]?\s*([NSEW])?'
    m = re.search(pattern, raw_str)
    if m:
        deg, mins, secs = float(m.group(1)), float(m.group(2)), float(m.group(3))
        direction = m.group(4) or ''
        decimal = deg + mins / 60 + secs / 3600
        if direction in ('S', 'W'):
            decimal = -decimal
        return round(decimal, 6)
    return 0.0


def parse_depth_str(s):
    """Extract depth number from strings like '2603m', '2478m MD RKB'."""
    if not s:
        return 0
    m = re.search(r'([\d.]+)\s*m', str(s).lower())
    if m:
        return float(m.group(1))
    return safe_float(s)


def parse_single_well(filepath):
    """Parse a single well card Excel file."""
    result = {'filepath': filepath, 'asset': get_asset_from_path(filepath)}

    try:
        xl = pd.ExcelFile(filepath, engine='openpyxl')
        sheets = xl.sheet_names
    except Exception as e:
        return {'error': f"Cannot open: {e}", 'filepath': filepath}

    if 'Well Details' not in sheets:
        return {'error': "No 'Well Details' sheet", 'filepath': filepath}

    # ═══════════════════════════════════════════════
    # PARSE WELL DETAILS SHEET
    # ═══════════════════════════════════════════════
    df = pd.read_excel(filepath, sheet_name='Well Details', header=None, engine='openpyxl')
    rows = df.fillna("").values

    # Fixed positions from inspection
    well_no = safe_str(cell(rows, 6, 3))
    well_name = safe_str(cell(rows, 7, 3))
    wbs = safe_str(cell(rows, 8, 3))
    category = safe_str(cell(rows, 10, 3))
    basin = safe_str(cell(rows, 11, 3))
    field = safe_str(cell(rows, 12, 3))
    location = safe_str(cell(rows, 13, 3))
    lat_raw = safe_str(cell(rows, 14, 3))
    lon_raw = safe_str(cell(rows, 15, 3))
    rig_type = safe_str(cell(rows, 16, 3))
    water_depth_raw = safe_str(cell(rows, 17, 3))
    rig_deployed = safe_str(cell(rows, 18, 3))

    # Cost details from Well Details (rows 19-22, col 11)
    cumulative_cost = safe_float(cell(rows, 20, 11))
    cost_per_barrel = safe_float(cell(rows, 21, 11))
    cost_per_metre = safe_float(cell(rows, 22, 11))
    msl_rkb = safe_str(cell(rows, 22, 3))

    # Target & Drilled depth (row 24 & 26)
    target_depth = parse_depth_str(safe_str(cell(rows, 24, 3)))
    drilled_depth = parse_depth_str(safe_str(cell(rows, 26, 3)))
    well_status = safe_str(cell(rows, 24, 7))

    # Spudding & TD dates (rows 30-34)
    spud_date = safe_str(cell(rows, 30, 3))
    td_date = safe_str(cell(rows, 31, 3))
    rig_release_date = safe_str(cell(rows, 34, 3))

    # Total planned & actual days (row 40)
    planned_days = safe_float(cell(rows, 40, 3))
    actual_days = safe_float(cell(rows, 40, 4))

    # Phases from Well Details (rows 11-17, cols 6-11)
    # Col 6=Phase, Col 7=Mud System, Col 9=Depth From, Col 11=Depth To
    # Stop when we reach non-phase rows (Cost Details section starts around row 19)
    PHASE_STOP_KEYWORDS = {
        'cost details', 'cumulative cost', 'cost per barrel', 'cost per metre',
        'cost per meter', 'well status', 'msl to rkb', 'target depth',
        'drilled depth', 'timeline', 'spudding', 'rig deployed', 'rig release',
        'well alerts', 'base officials'
    }
    phases_detail = []
    for i in range(11, min(25, len(rows))):
        phase_name = safe_str(cell(rows, i, 6))
        mud_system = safe_str(cell(rows, i, 7))
        if not phase_name or phase_name in ('0', 'Phase', 'Mud System Used', 'Cost Details'):
            continue
        # Stop if this row is part of Cost Details or other non-phase sections
        if phase_name.strip().lower() in PHASE_STOP_KEYWORDS:
            continue
        # Also check col 2 label to detect if we've entered the cost/metadata section
        col2_label = safe_str(cell(rows, i, 2)).strip().lower()
        if col2_label in PHASE_STOP_KEYWORDS or col2_label in ('msl to rkb', 'target depth', 'drilled depth'):
            continue
        depth_from = safe_float(cell(rows, i, 9))
        depth_to = safe_float(cell(rows, i, 11))

        # Determine hole size from phase name
        hole_size = "N/A"
        hs_match = re.search(r'(\d+\.?\d*(?:\s*[-/]\s*\d+)?)\s*["\u201c\u201d]', phase_name)
        if hs_match:
            hole_size = hs_match.group(0).strip()
        elif re.match(r'^\d+', phase_name):
            hs_match2 = re.search(r'^(\d+\.?\d*(?:\s*[-/]\s*\d+)?["\u201c\u201d\']?)', phase_name)
            if hs_match2:
                hs = hs_match2.group(1).strip()
                if '"' in hs or "'" in hs:
                    hole_size = hs

        phases_detail.append({
            'phase': phase_name, 'mud_type': mud_system, 'hole_size': hole_size,
            'depth_from': depth_from, 'depth_to': depth_to,
        })

    # ═══════════════════════════════════════════════
    # PARSE PERFORMANCE SHEET
    # ═══════════════════════════════════════════════
    perf_phases = []
    npt_data = {'mud_loss_hrs': 0, 'activity_hrs': 0, 'stuck_up_hrs': 0, 'waiting_hrs': 0}
    phase_costs = {}  # phase -> {cost_per_m, cost_per_bbl, corrected_cost_per_bbl}
    key_indicators = {}
    perf_total_planned = 0
    perf_total_actual = 0

    if 'Performance Sheet' in sheets:
        df = pd.read_excel(filepath, sheet_name='Performance Sheet', header=None, engine='openpyxl')
        rows_p = df.fillna("").values

        # Phase performance (rows 3-9)
        for i in range(3, min(12, len(rows_p))):
            phase = safe_str(cell(rows_p, i, 1))
            if not phase or phase == '0':
                continue
            if 'total' in phase.lower():
                perf_total_planned = safe_float(cell(rows_p, i, 2))
                perf_total_actual = safe_float(cell(rows_p, i, 7))
                continue
            p_days = safe_float(cell(rows_p, i, 2))
            a_days = safe_float(cell(rows_p, i, 7))
            # Fix serial number dates stored as actual days
            if a_days > 1000:
                a_days = 0
            start = safe_str(cell(rows_p, i, 3))
            end = safe_str(cell(rows_p, i, 5))

            # Cost from performance sheet (cols 12, 15, 18)
            est_cost = safe_float(cell(rows_p, i, 12))
            act_cost = safe_float(cell(rows_p, i, 15))

            perf_phases.append({
                'phase': phase, 'planned_days': p_days, 'actual_days': a_days,
                'start_date': start, 'end_date': end,
                'cost_estimate_inr': est_cost, 'actual_cost_inr': act_cost,
            })

        # Row 10: Totals
        if perf_total_planned == 0:
            perf_total_planned = safe_float(cell(rows_p, 10, 2))
            perf_total_actual = safe_float(cell(rows_p, 10, 7))

        # NPT data (rows 15-22, cols 1-5)
        # Row 15 has headers: Phase, Mud Loss (Hrs), Activity (Hrs), Stuck-up (Hrs), Chemical Waiting(Hrs)
        # Rows 16+ have data per phase
        # First try to find the NPT section header row
        npt_header_row = -1
        for i in range(13, min(20, len(rows_p))):
            label = safe_str(cell(rows_p, i, 1)).lower()
            if 'npt' in label:
                npt_header_row = i
                break

        if npt_header_row >= 0:
            # Check if header row+1 has the column labels (Phase, Mud Loss, Activity, etc.)
            # Data starts from header_row+2 (or +1 if header row itself has labels)
            label_row = npt_header_row + 1
            data_start = npt_header_row + 2

            # Check if label_row has "Phase" or "Mud Loss" text
            label_check = safe_str(cell(rows_p, label_row, 2)).lower()
            if 'mud' in label_check or 'loss' in label_check:
                data_start = label_row + 1
            else:
                data_start = label_row  # data starts right after NPT header

            found_total = False
            for i in range(data_start, min(data_start + 12, len(rows_p))):
                phase = safe_str(cell(rows_p, i, 1))
                if not phase or phase == '0':
                    continue
                if 'total' in phase.lower() or 'bridging' in phase.lower():
                    # Use total row
                    if npt_data['mud_loss_hrs'] == 0 and npt_data['activity_hrs'] == 0:
                        npt_data['mud_loss_hrs'] = safe_float(cell(rows_p, i, 2))
                        npt_data['activity_hrs'] = safe_float(cell(rows_p, i, 3))
                        npt_data['stuck_up_hrs'] = safe_float(cell(rows_p, i, 4))
                        npt_data['waiting_hrs'] = safe_float(cell(rows_p, i, 5))
                    found_total = True
                    break
                ml = safe_float(cell(rows_p, i, 2))
                act = safe_float(cell(rows_p, i, 3))
                su = safe_float(cell(rows_p, i, 4))
                wt = safe_float(cell(rows_p, i, 5))
                npt_data['mud_loss_hrs'] += ml
                npt_data['activity_hrs'] += act
                npt_data['stuck_up_hrs'] += su
                npt_data['waiting_hrs'] += wt

        # Cost analysis per phase (rows 15+, cols 13-18)
        # Col 13=Phase, Col 14=Cost/m, Col 16=Cost/bbl, Col 18=Corrected Cost/bbl
        for i in range(15, min(26, len(rows_p))):
            phase = safe_str(cell(rows_p, i, 13))
            if not phase or phase in ('0', 'Phase', 'Comp/Testing'):
                continue
            cpm = safe_float(cell(rows_p, i, 14))
            cpb = safe_float(cell(rows_p, i, 16))
            corrected_cpb = safe_float(cell(rows_p, i, 18))
            if phase == 'Total' or phase == 'Drilling ':
                # These are summary rows
                phase_costs['_total'] = {'cost_per_m': cpm, 'cost_per_bbl': cpb, 'corrected_cost_per_bbl': corrected_cpb}
            else:
                phase_costs[phase.strip()] = {'cost_per_m': cpm, 'cost_per_bbl': cpb, 'corrected_cost_per_bbl': corrected_cpb}

        # Key Indicators (rows 29-39, col 11=label, col 14=value)
        for i in range(29, min(40, len(rows_p))):
            label = safe_str(cell(rows_p, i, 11)).strip().lower()
            val = safe_float(cell(rows_p, i, 14))
            if 'cumulative cost' in label:
                key_indicators['cumulative_cost'] = val
            elif 'drilling cost' in label and 'completion' not in label:
                key_indicators['drilling_cost'] = val
            elif 'completion cost' in label:
                key_indicators['completion_cost'] = val
            elif 'meterage drilled' in label:
                key_indicators['meterage_drilled'] = val
            elif 'drilling volume' in label and 'corrected' not in label:
                key_indicators['drilling_volume'] = val
            elif 'drilling corrected volume' in label:
                key_indicators['drilling_corrected_volume'] = val
            elif 'total well volume' in label and 'corrected' not in label:
                key_indicators['total_well_volume'] = val
            elif 'total corrected well volume' in label:
                key_indicators['total_corrected_volume'] = val
            elif 'completion volume' in label and 'corrected' not in label:
                key_indicators['completion_volume'] = val

    # ═══════════════════════════════════════════════
    # PARSE COST SHEET (for chemicals)
    # ═══════════════════════════════════════════════
    chemicals = []
    cost_phase_columns = []  # track phase column positions

    if 'Cost' in sheets:
        df = pd.read_excel(filepath, sheet_name='Cost', header=None, engine='openpyxl')
        rows_c = df.fillna("").values
        num_cols = df.shape[1]

        # Row 2 has phase names at columns 9, 15, 21, 27, 33, 39, 45 (every 6 cols)
        # Row 3 has mud types at same columns
        # Row 5 has Volume Handled at cols 14, 20, 26, 32, 38, 44, 50 (col + 5)
        # Row 7 has depth from/to
        # Row 9 headers: Conc(+0), Consumption(+1), Cost(+2), Consumption(+3), Cost(+4), Conc(+5)
        # So Actual Consumption = phase_col + 3, Actual Cost = phase_col + 4

        phase_start_cols = [9, 15, 21, 27, 33, 39, 45]
        for pc in phase_start_cols:
            if pc >= num_cols:
                break
            pname = safe_str(cell(rows_c, 2, pc))
            if pname and pname not in ('0', '', 'CUMULATIVE'):
                mud = safe_str(cell(rows_c, 3, pc))
                vol = safe_float(cell(rows_c, 5, pc + 5))  # Volume Handled
                d_from = safe_float(cell(rows_c, 7, pc))
                d_to = safe_float(cell(rows_c, 7, pc + 3))
                interval = safe_float(cell(rows_c, 7, pc + 5))
                cost_phase_columns.append({
                    'phase': pname, 'mud_type': mud, 'col': pc,
                    'volume_handled': vol, 'depth_from': d_from, 'depth_to': d_to,
                    'interval': interval,
                    'actual_consumption_col': pc + 3,
                    'actual_cost_col': pc + 4,
                })

        # Parse chemical rows (row 10+)
        for i in range(10, min(100, len(rows_c))):
            item_name = safe_str(cell(rows_c, i, 3))
            if not item_name or item_name == '0':
                continue
            unit_size = safe_str(cell(rows_c, i, 4))
            unit = safe_str(cell(rows_c, i, 5))
            packaging = safe_str(cell(rows_c, i, 6))
            price_inr = safe_float(cell(rows_c, i, 8))

            # Cumulative consumption & cost (cols 51, 53)
            cum_consumption = safe_float(cell(rows_c, i, 51))
            cum_cost = safe_float(cell(rows_c, i, 53))

            # Per-phase data
            for cpc in cost_phase_columns:
                consumption = safe_float(cell(rows_c, i, cpc['actual_consumption_col']))
                cost_val = safe_float(cell(rows_c, i, cpc['actual_cost_col']))
                if consumption > 0 or cost_val > 0:
                    chemicals.append({
                        'chemical_name': item_name,
                        'unit_size': f"{unit_size} {unit}".strip(),
                        'unit': unit,
                        'packaging': packaging,
                        'phase': cpc['phase'],
                        'consumption': consumption,
                        'actual_cost_inr': cost_val,
                        'price_per_unit_inr': price_inr,
                    })

            # If no phase data but has cumulative, store as cumulative
            if cum_consumption > 0 and not any(
                c['chemical_name'] == item_name for c in chemicals
                if c.get('consumption', 0) > 0
            ):
                chemicals.append({
                    'chemical_name': item_name,
                    'unit_size': f"{unit_size} {unit}".strip(),
                    'unit': unit,
                    'packaging': packaging,
                    'phase': 'Cumulative',
                    'consumption': cum_consumption,
                    'actual_cost_inr': cum_cost,
                    'price_per_unit_inr': price_inr,
                })

    # ═══════════════════════════════════════════════
    # PARSE DAILY MUD PARA SHEET
    # ═══════════════════════════════════════════════
    mud_parameters = []  # list of per-phase last-recorded mud parameters

    if 'DAILY MUD PARA' in sheets:
        try:
            df_mud = pd.read_excel(filepath, sheet_name='DAILY MUD PARA', header=None, engine='openpyxl')
            rows_mud = df_mud.fillna("").values

            # Column mapping (Row 2): 11=MW, 12=FV, 13=PV, 14=YP, 15=GEL0, 16=GEL10,
            # 17=R6, 18=R3, 19=OWR(Oil%), 20=Water%, 21=Solid%, 22=Chlorides, 23=HTHP,
            # 24=Ex.Lime, 25=ES, 26=WPS, 27=pH, 28=FLT
            MUD_PARA_COLS = {
                'mud_weight_ppg': 11, 'fv_sec': 12, 'pv_cp': 13, 'yp_lb100ft2': 14,
                'gel0_lb100ft2': 15, 'gel10_lb100ft2': 16, 'r6': 17, 'r3': 18,
                'owr_oil_pct': 19, 'owr_water_pct': 20, 'solid_pct': 21,
                'chlorides_ppm': 22, 'hthp_fl_ml': 23, 'ex_lime_ppb': 24,
                'es_v': 25, 'wps_ppm': 26, 'ph': 27, 'flt_c': 28,
            }

            # Collect ALL mud param values per phase to compute min/max/last ranges
            from collections import defaultdict
            phase_all_values = defaultdict(lambda: {
                'rows': [], 'mud_system': '', 'formation': '', 'layer': '', 'lithology': '',
                'depth_max': 0, 'values': defaultdict(list)
            })

            for i in range(4, len(rows_mud)):
                day_val = safe_str(cell(rows_mud, i, 1))
                if not day_val or not day_val.replace('.', '').isdigit():
                    continue
                phase = safe_str(cell(rows_mud, i, 6))
                if not phase or phase in ('0', 'Hole Size (in)'):
                    continue
                mw = safe_float(cell(rows_mud, i, 11))
                if mw > 0:  # Only count rows with actual mud weight recorded
                    pdata = phase_all_values[phase]
                    pdata['rows'].append(i)
                    pdata['mud_system'] = safe_str(cell(rows_mud, i, 5)) or pdata['mud_system']
                    pdata['formation'] = safe_str(cell(rows_mud, i, 8)) or pdata['formation']
                    pdata['layer'] = safe_str(cell(rows_mud, i, 9)) or pdata['layer']
                    pdata['lithology'] = safe_str(cell(rows_mud, i, 10)) or pdata['lithology']
                    depth_val = safe_float(cell(rows_mud, i, 4))
                    if depth_val > pdata['depth_max']:
                        pdata['depth_max'] = depth_val
                    for param_name, col_idx in MUD_PARA_COLS.items():
                        v = safe_float(cell(rows_mud, i, col_idx))
                        if v > 0:
                            pdata['values'][param_name].append(v)

            for phase, pdata in phase_all_values.items():
                params = {
                    'phase': phase,
                    'depth': pdata['depth_max'],
                    'mud_system': pdata['mud_system'],
                    'formation': pdata['formation'],
                    'layer': pdata['layer'],
                    'lithology': pdata['lithology'],
                }
                for param_name in MUD_PARA_COLS:
                    vals = pdata['values'].get(param_name, [])
                    if vals:
                        params[param_name + '_min'] = round(min(vals), 2)
                        params[param_name + '_max'] = round(max(vals), 2)
                        params[param_name + '_last'] = round(vals[-1], 2)
                    else:
                        params[param_name + '_min'] = 0
                        params[param_name + '_max'] = 0
                        params[param_name + '_last'] = 0
                mud_parameters.append(params)

        except Exception as e:
            pass  # Skip if sheet is unreadable

    # ═══════════════════════════════════════════════
    # PARSE COMPLICATION SHEET
    # ═══════════════════════════════════════════════
    complications = {'mud_loss': [], 'well_activity': [], 'stuck_up': []}

    if 'COMPLICATION' in sheets:
        df = pd.read_excel(filepath, sheet_name='COMPLICATION', header=None, engine='openpyxl')
        rows_comp = df.fillna("").values

        current_section = None
        for i in range(len(rows_comp)):
            # Detect section headers
            for j in range(min(3, len(rows_comp[i]))):
                c_val = safe_str(rows_comp[i][j]).upper()
                if 'MUD LOSS' in c_val and 'SR' not in c_val:
                    current_section = 'mud_loss'
                elif 'WELL ACTIVITY' in c_val and 'SR' not in c_val:
                    current_section = 'well_activity'
                elif 'STUCK UP' in c_val and 'SR' not in c_val:
                    current_section = 'stuck_up'
                elif 'NPT DUE TO WOL' in c_val:
                    current_section = 'npt_wol'

            if current_section is None or current_section == 'npt_wol':
                continue

            sr = safe_str(cell(rows_comp, i, 1))
            if not sr or not sr.replace('.', '').replace('-', '').isdigit():
                continue

            if current_section == 'mud_loss':
                event = {
                    'date': safe_str(cell(rows_comp, i, 2)),
                    'phase': safe_str(cell(rows_comp, i, 3)),
                    'drill_depth_m': safe_float(cell(rows_comp, i, 4)),
                    'depth_occ_m': safe_float(cell(rows_comp, i, 5)),
                    'contract': safe_str(cell(rows_comp, i, 6)),
                    'mud_system': safe_str(cell(rows_comp, i, 7)),
                    'operation': safe_str(cell(rows_comp, i, 8)),
                    'type': safe_str(cell(rows_comp, i, 9)),
                    'formation': safe_str(cell(rows_comp, i, 10)),
                    'layer': safe_str(cell(rows_comp, i, 11)),
                    'pill_type': safe_str(cell(rows_comp, i, 20)),
                    'volume_lost_bbl': safe_float(cell(rows_comp, i, 28)),
                    'npt_hrs': safe_float(cell(rows_comp, i, 30)),
                }
                complications['mud_loss'].append(event)

            elif current_section == 'well_activity':
                event = {
                    'date': safe_str(cell(rows_comp, i, 2)),
                    'phase': safe_str(cell(rows_comp, i, 3)),
                    'drill_depth_m': safe_float(cell(rows_comp, i, 4)),
                    'depth_occ_m': safe_float(cell(rows_comp, i, 5)),
                    'mud_system': safe_str(cell(rows_comp, i, 6)),
                    'operation': safe_str(cell(rows_comp, i, 7)),
                    'type': safe_str(cell(rows_comp, i, 8)),
                    'formation': safe_str(cell(rows_comp, i, 9)),
                    'layer': safe_str(cell(rows_comp, i, 10)),
                    'pill_type': safe_str(cell(rows_comp, i, 20)),
                    'volume_lost_bbl': 0,
                    'npt_hrs': safe_float(cell(rows_comp, i, 30)),
                }
                complications['well_activity'].append(event)

            elif current_section == 'stuck_up':
                event = {
                    'date': safe_str(cell(rows_comp, i, 2)),
                    'phase': safe_str(cell(rows_comp, i, 3)),
                    'drill_depth_m': safe_float(cell(rows_comp, i, 4)),
                    'depth_occ_m': safe_float(cell(rows_comp, i, 5)),
                    'mud_system': safe_str(cell(rows_comp, i, 6)),
                    'operation': safe_str(cell(rows_comp, i, 7)),
                    'type': safe_str(cell(rows_comp, i, 8)),
                    'formation': safe_str(cell(rows_comp, i, 9)),
                    'layer': safe_str(cell(rows_comp, i, 10)),
                    'pill_type': safe_str(cell(rows_comp, i, 15)),
                    'action_taken': safe_str(cell(rows_comp, i, 20)),
                    'volume_lost_bbl': 0,
                    'npt_hrs': safe_float(cell(rows_comp, i, 30)),
                }
                complications['stuck_up'].append(event)

    # ═══════════════════════════════════════════════
    # BUILD FINAL WELL RECORD
    # ═══════════════════════════════════════════════

    # Derive well name
    final_name = well_name or well_no
    if not final_name:
        fname = os.path.basename(filepath).replace('.xlsx', '')
        fname = re.sub(r'(?i)(well[_ ]card[_ ]?|inclusive.*|latest.*|updated.*|new format.*|\(\d+\))', '', fname).strip()
        final_name = fname.replace('_', ' ').strip() or "Unknown"

    # ── METERAGE: Sum of Interval from Cost sheet, EXCLUDING "ST Prep" and "Comp" phases ──
    EXCLUDE_PHASES_FOR_METERAGE = {'ST PREP', 'COMP', 'COMPLETION', 'COMP/TESTING'}
    meterage = 0.0
    for cpc in cost_phase_columns:
        pname_upper = cpc['phase'].strip().upper()
        if pname_upper in EXCLUDE_PHASES_FOR_METERAGE:
            continue
        interval = cpc.get('interval', 0)
        if interval > 0:
            meterage += interval
    # Fallback 1: Key Indicators "Meterage Drilled"
    if meterage == 0:
        meterage = key_indicators.get('meterage_drilled', 0)
    # Fallback 2: sum of depth intervals from Well Details phases (excluding ST Prep/Comp)
    if meterage == 0:
        for p in phases_detail:
            pname_upper = p.get('phase', '').strip().upper()
            if pname_upper in EXCLUDE_PHASES_FOR_METERAGE:
                continue
            diff = max(0, p['depth_to'] - p['depth_from'])
            meterage += diff

    # ── MAX DEPTH: Maximum "Phase To" from Cost sheet columns (most reliable) ──
    max_depth = 0.0
    for cpc in cost_phase_columns:
        d_to = cpc.get('depth_to', 0)
        if 0 < d_to < 10000 and d_to > max_depth:
            max_depth = d_to
    # Fallback 1: maximum depth_to from Well Details phases
    # Only use phases that look like real drilling phases (have a hole size or known phase name)
    if max_depth == 0:
        for p in phases_detail:
            d_to = p.get('depth_to', 0)
            pname = p.get('phase', '').strip().upper()
            # Skip non-drilling phases
            if pname in ('COMP', 'COMPLETION', 'ST PREP'):
                continue
            if 0 < d_to < 10000 and d_to > max_depth:
                max_depth = d_to
    # Fallback 2: drilled_depth from Well Details sheet (with sanity check)
    if max_depth == 0:
        dd = drilled_depth if drilled_depth < 10000 else 0
        if dd > 0:
            max_depth = dd
    # Fallback 3: target_depth
    if max_depth == 0:
        max_depth = target_depth if target_depth < 10000 else 0

    total_mud = key_indicators.get('total_well_volume', 0)
    if total_mud == 0:
        total_mud = key_indicators.get('drilling_volume', 0)
    if total_mud == 0:
        total_mud = sum(cpc.get('volume_handled', 0) for cpc in cost_phase_columns)

    # Use Well Details cost values as primary (they are pre-calculated)
    final_cost = cumulative_cost
    if final_cost == 0:
        final_cost = key_indicators.get('cumulative_cost', 0)

    # Cost per meter/barrel - use Well Details values, fall back to Key Indicators, then calculate
    final_cpm = cost_per_metre
    if final_cpm <= 1:
        # Try from Key Indicators total
        tc = phase_costs.get('_total', {})
        final_cpm = tc.get('cost_per_m', 0)
    if final_cpm <= 1:
        tc = phase_costs.get('Drilling ', {})
        final_cpm = tc.get('cost_per_m', 0)
    if final_cpm <= 1:
        tc = phase_costs.get('Total', {})
        final_cpm = tc.get('cost_per_m', 0)
    if final_cpm <= 1 and meterage > 0 and final_cost > 0:
        final_cpm = final_cost / meterage

    final_cpb = cost_per_barrel
    if final_cpb <= 1:
        tc = phase_costs.get('_total', {})
        final_cpb = tc.get('corrected_cost_per_bbl', 0) or tc.get('cost_per_bbl', 0)
    if final_cpb <= 1:
        tc = phase_costs.get('Total', {})
        final_cpb = tc.get('corrected_cost_per_bbl', 0) or tc.get('cost_per_bbl', 0)
    if final_cpb <= 1 and total_mud > 0 and final_cost > 0:
        final_cpb = final_cost / total_mud

    # Planned/actual days
    final_planned = planned_days or perf_total_planned
    final_actual = actual_days or perf_total_actual
    if final_actual > 1000:
        final_actual = perf_total_actual if perf_total_actual < 1000 else 0

    # Merge phase data
    merged_phases = []
    for pd_phase in phases_detail:
        pname = pd_phase['phase']
        # Find performance match
        perf_match = None
        for pp in perf_phases:
            if pp['phase'].strip().replace('"', '"') == pname.strip().replace('"', '"'):
                perf_match = pp
                break
            # Partial match
            p1 = pname.split('"')[0].strip()
            p2 = pp['phase'].split('"')[0].strip()
            if p1 and p2 and (p1 in pp['phase'] or p2 in pname):
                perf_match = pp
                break

        # Find cost phase match
        cost_match = None
        for cpc in cost_phase_columns:
            if cpc['phase'].strip().replace('"', '"') == pname.strip().replace('"', '"'):
                cost_match = cpc
                break
            p1 = pname.split('"')[0].strip()
            p2 = cpc['phase'].split('"')[0].strip()
            if p1 and p2 and (p1 in cpc['phase'] or p2 in pname):
                cost_match = cpc
                break

        # Cost/m and cost/bbl from performance sheet phase_costs
        pc = phase_costs.get(pname.strip(), {})

        merged_phases.append({
            'phase': pname,
            'hole_size': pd_phase.get('hole_size', 'N/A'),
            'mud_type': pd_phase.get('mud_type', ''),
            'depth_from': pd_phase.get('depth_from', 0),
            'depth_to': pd_phase.get('depth_to', 0),
            'planned_days': perf_match['planned_days'] if perf_match else 0,
            'actual_days': perf_match['actual_days'] if perf_match else 0,
            'start_date': perf_match.get('start_date', '') if perf_match else '',
            'end_date': perf_match.get('end_date', '') if perf_match else '',
            'cost_estimate_inr': perf_match.get('cost_estimate_inr', 0) if perf_match else 0,
            'actual_cost_inr': perf_match.get('actual_cost_inr', 0) if perf_match else 0,
            'volume_handled': cost_match.get('volume_handled', 0) if cost_match else 0,
            'cost_per_meter': pc.get('cost_per_m', 0),
            'cost_per_barrel': pc.get('cost_per_bbl', 0),
            'corrected_cost_per_barrel': pc.get('corrected_cost_per_bbl', 0),
        })

    # If no phases from details, use perf phases
    if not merged_phases and perf_phases:
        for pp in perf_phases:
            merged_phases.append({
                'phase': pp['phase'], 'hole_size': 'N/A', 'mud_type': '',
                'depth_from': 0, 'depth_to': 0,
                'planned_days': pp['planned_days'], 'actual_days': pp['actual_days'],
                'start_date': pp.get('start_date', ''), 'end_date': pp.get('end_date', ''),
                'cost_estimate_inr': pp.get('cost_estimate_inr', 0),
                'actual_cost_inr': pp.get('actual_cost_inr', 0),
                'volume_handled': 0, 'cost_per_meter': 0, 'cost_per_barrel': 0,
                'corrected_cost_per_barrel': 0,
            })

    # Override category if folder hints at subcategory
    sub_cat = get_sub_category_from_path(filepath)
    if sub_cat and not category:
        category = sub_cat

    well_record = {
        'well_name': final_name,
        'well_no': well_no,
        'wbs': wbs,
        'asset': get_asset_from_path(filepath),
        'basin': basin,
        'field': field,
        'category': category,
        'location': location,
        'rig_type': rig_type,
        'rig_deployed': rig_deployed,
        'latitude': parse_lat_lon(lat_raw),
        'longitude': parse_lat_lon(lon_raw),
        'water_depth_m': parse_depth_str(water_depth_raw),
        'msl_rkb': msl_rkb,
        'well_status': well_status,
        'spud_date': spud_date,
        'td_date': td_date,
        'rig_release_date': rig_release_date,
        'planned_days': final_planned,
        'actual_days': final_actual,
        'target_depth_m': target_depth,
        'drilled_depth_m': drilled_depth,
        'max_depth_m': max_depth,
        'meterage_m': meterage,
        'total_mud_bbl': total_mud,
        'total_cost_inr': final_cost,
        'cost_per_meter_inr': round(final_cpm, 2),
        'cost_per_barrel_inr': round(final_cpb, 2),
        'key_indicators': key_indicators,
        'npt': npt_data,
        'phases': merged_phases,
        'chemicals': chemicals,
        'complications_mud_loss': complications['mud_loss'],
        'complications_well_activity': complications['well_activity'],
        'complications_stuck_up': complications['stuck_up'],
        'mud_parameters': mud_parameters,
    }

    return well_record


def build_unified_data(well_cards_dir=WELL_CARDS_DIR):
    files = find_well_card_files(well_cards_dir)
    print(f"Found {len(files)} potential well card files")

    all_wells = []
    errors = []
    seen_wells = set()

    for i, fp in enumerate(files):
        fname = os.path.basename(fp)
        print(f"  [{i+1}/{len(files)}] {fname}")
        try:
            result = parse_single_well(fp)
            if 'error' in result:
                errors.append({'file': fp, 'error': result['error']})
                continue

            well_name = result.get('well_name', 'Unknown')
            name_key = well_name.upper().replace(' ', '').replace('-', '').replace('_', '')
            if name_key in seen_wells:
                print(f"    -> Skip duplicate: {well_name}")
                continue
            seen_wells.add(name_key)

            all_wells.append(result)
            npt_total = sum(result['npt'].values())
            print(f"    -> {well_name} | depth={result['max_depth_m']}m | mud={result['total_mud_bbl']}bbl | "
                  f"cost={result['total_cost_inr']:,.0f} | cpm={result['cost_per_meter_inr']:,.0f} | "
                  f"cpb={result['cost_per_barrel_inr']:,.0f} | npt={npt_total}hrs")

        except Exception as e:
            errors.append({'file': fp, 'error': str(e)})
            print(f"    -> ERROR: {e}")

    # ── Post-processing: infer missing coordinates from same field/platform ──
    # Build mapping: field -> (lat, lon) from wells that have both
    field_coords = {}
    platform_coords = {}
    for w in all_wells:
        lat, lon = w.get('latitude', 0), w.get('longitude', 0)
        if lat != 0 and lon != 0:
            f = (w.get('field', '') or '').strip().upper()
            loc = (w.get('location', '') or '').strip().upper()
            if f and f not in field_coords:
                field_coords[f] = (lat, lon)
            if loc and loc not in platform_coords:
                platform_coords[loc] = (lat, lon)

    inferred_count = 0
    for w in all_wells:
        lat, lon = w.get('latitude', 0), w.get('longitude', 0)
        if lat != 0 and lon != 0:
            continue  # already has both coords

        # Try 1: match by field name
        f = (w.get('field', '') or '').strip().upper()
        if f and f in field_coords:
            ref_lat, ref_lon = field_coords[f]
            if lat == 0:
                w['latitude'] = ref_lat
            if lon == 0:
                w['longitude'] = ref_lon
            inferred_count += 1
            continue

        # Try 2: match by location/platform
        loc = (w.get('location', '') or '').strip().upper()
        if loc and loc in platform_coords:
            ref_lat, ref_lon = platform_coords[loc]
            if lat == 0:
                w['latitude'] = ref_lat
            if lon == 0:
                w['longitude'] = ref_lon
            inferred_count += 1
            continue

        # Try 3: match field with partial/fuzzy match
        for known_f, coords in field_coords.items():
            if (f and known_f and (f in known_f or known_f in f)):
                if lat == 0:
                    w['latitude'] = coords[0]
                if lon == 0:
                    w['longitude'] = coords[1]
                inferred_count += 1
                break

    print(f"  Coordinates inferred for {inferred_count} additional wells")

    final_with = sum(1 for w in all_wells if w.get('latitude', 0) != 0 and w.get('longitude', 0) != 0)
    print(f"  Total wells with coordinates: {final_with}/{len(all_wells)}")

    print(f"\nParsed: {len(all_wells)} wells, Errors: {len(errors)}")
    return all_wells, errors


def save_to_json(all_wells, output_path=OUTPUT_JSON):
    def clean(obj):
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [clean(item) for item in obj]
        elif isinstance(obj, (np.integer, np.int64)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64)):
            return float(obj) if not np.isnan(obj) else 0.0
        elif isinstance(obj, float) and (np.isnan(obj) or np.isinf(obj)):
            return 0.0
        return obj

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(clean(all_wells), f, indent=2, ensure_ascii=False, default=str)
    print(f"Saved {len(all_wells)} wells to {output_path}")


if __name__ == '__main__':
    wells, errors = build_unified_data()
    save_to_json(wells)

    print(f"\n{'='*60}")
    print(f"SUMMARY: {len(wells)} wells")
    # Verify key data
    for w in wells[:5]:
        print(f"  {w['well_name']}: depth={w['max_depth_m']}m, mud={w['total_mud_bbl']}bbl, "
              f"cost={w['total_cost_inr']:,.0f}, cpm={w['cost_per_meter_inr']:,.0f}, "
              f"cpb={w['cost_per_barrel_inr']:,.0f}, NPT={sum(w['npt'].values())}hrs")
