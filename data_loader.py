"""
Dynamic data loader for Well Dashboard.
Reads from all_wells_data.json (produced by dynamic_parser.py).
"""
import os
import json
import re
import glob
import threading
import time
from pathlib import Path

import pandas as pd
import numpy as np

# Paths
DATA_JSON = os.path.join(os.path.dirname(__file__), "all_wells_data.json")
WELL_CARDS_DIR = os.environ.get("WELL_CARDS_DIR", os.path.join(os.path.dirname(__file__), "well_cards", "Well Cards"))

# Category normalization
CATEGORY_MAP = {
    'development': 'Development',
    'side track': 'Side Track',
    'sidetrack': 'Side Track',
    'side-track': 'Side Track',
    'side track (development)': 'Side Track',
    'side-track (op)': 'Side Track',
    'side track (op)': 'Side Track',
    'side track(op)': 'Side Track',
    'side track job': 'Side Track',
    'side-track job': 'Side Track',
    'st': 'Side Track',
    'development sidetrack': 'Side Track',
    'exploratory': 'Exploratory',
    'expl b type': 'Exploratory',
    'exp': 'Exploratory',
    'exploratory b': 'Exploratory',
    'exploratory, b': 'Exploratory',
    "exploratory,'b'": 'Exploratory',
    "exploratory, 'b'": 'Exploratory',
    'exploratory b / appraisal': 'Exploratory',
    'workover': 'Workover',
    'work over': 'Workover',
    'work over job': 'Workover',
    'wo': 'Workover',
    'woj': 'Workover',
    'woj - 2': 'Workover',
    'development gas well': 'Development',
    'oil producer': 'Development',
}


def load_wells_json():
    """Load all wells from JSON."""
    with open(DATA_JSON, 'r', encoding='utf-8') as f:
        return json.load(f)


def normalize_category(cat):
    """Normalize category string."""
    if not cat:
        return 'Other'
    return CATEGORY_MAP.get(cat.strip().lower(), cat.strip().title() or 'Other')


def normalize_mud_type(mud):
    """Normalize mud type names."""
    if not mud:
        return 'Unknown'
    mud_up = mud.strip().upper()
    if 'SOBM' in mud_up or 'MI-SOBM' in mud_up or 'HLB-SOBM' in mud_up or 'MI-LTSOBM' in mud_up:
        return 'OBM (SOBM)'
    if 'NDDF' in mud_up or 'NA-FO' in mud_up:
        return 'NDDF'
    if 'GEL MUD' in mud_up or 'GEL POLY' in mud_up:
        return 'Gel Polymer'
    if 'KPP' in mud_up or 'KCL' in mud_up:
        return 'KCl Polymer'
    if 'WBM' in mud_up or 'WATER BASED' in mud_up:
        return 'WBM'
    if 'RAW' in mud_up or 'SW' in mud_up:
        return 'Raw/Sea Water'
    if mud_up in ('0', ''):
        return 'Unknown'
    return mud.strip()


def build_wells_dataframe():
    """Build main wells DataFrame."""
    wells = load_wells_json()
    records = []
    for w in wells:
        cat = normalize_category(w.get('category', ''))
        total_npt = sum(w.get('npt', {}).values())
        records.append({
            "Well Name": w['well_name'],
            "Well No": w.get('well_no', w['well_name']),
            "WBS": w.get('wbs', ''),
            "Asset": w.get('asset', ''),
            "Basin": w.get('basin', ''),
            "Field": w.get('field', ''),
            "Category": cat,
            "Location": w.get('location', ''),
            "Rig Type": w.get('rig_type', ''),
            "Rig": w.get('rig_deployed', ''),
            "Latitude": w.get('latitude', 0),
            "Longitude": w.get('longitude', 0),
            "Water Depth (m)": w.get('water_depth_m', 0),
            "Spud Date": w.get('spud_date', ''),
            "TD Date": w.get('td_date', ''),
            "Planned Days": w.get('planned_days', 0),
            "Actual Days": w.get('actual_days', 0),
            "Time Variance (days)": round(w.get('planned_days', 0) - w.get('actual_days', 0), 2),
            "Target Depth (m)": w.get('max_depth_m', 0),
            "Max Depth (m)": w.get('max_depth_m', 0),
            "Meterage (m)": w.get('meterage_m', 0),
            "Total Mud Handled (bbl)": w.get('total_mud_bbl', 0),
            "Cost per Meter (INR)": round(w.get('cost_per_meter_inr', 0), 2),
            "Cost per Barrel (INR)": round(w.get('cost_per_barrel_inr', 0), 2),
            "Total Cost (INR)": w.get('total_cost_inr', 0),
            "Mud Loss NPT (Hrs)": w.get('npt', {}).get('mud_loss_hrs', 0),
            "Activity NPT (Hrs)": w.get('npt', {}).get('activity_hrs', 0),
            "Unplanned Waiting NPT (Hrs)": w.get('npt', {}).get('waiting_hrs', 0),
            "Stuck Up NPT (Hrs)": w.get('npt', {}).get('stuck_up_hrs', 0),
            "Total NPT (Hrs)": total_npt,
            "Mud Loss Events": len(w.get('complications_mud_loss', [])),
            "Well Activity Events": len(w.get('complications_well_activity', [])),
            "Stuck Up Events": len(w.get('complications_stuck_up', [])),
            "Total Complications": (len(w.get('complications_mud_loss', []))
                                    + len(w.get('complications_well_activity', []))
                                    + len(w.get('complications_stuck_up', []))),
            "Well Status": w.get('well_status', ''),
        })
    return pd.DataFrame(records)


def build_phases_dataframe():
    """Build phases DataFrame."""
    wells = load_wells_json()
    records = []
    for w in wells:
        for ph in w.get('phases', []):
            phase_name = ph.get('phase', '')
            if not phase_name or phase_name == '0':
                continue
            # Determine hole size
            hole_size = ph.get('hole_size', 'N/A')
            if hole_size == 'N/A':
                hs_match = re.search(r'(\d+\.?\d*(?:\s*[-/]\s*\d+)?)\s*["\']', phase_name)
                if hs_match:
                    hole_size = hs_match.group(0).strip()

            records.append({
                "Well Name": w['well_name'],
                "Asset": w.get('asset', ''),
                "Field": w.get('field', ''),
                "Phase": phase_name,
                "Hole Size": hole_size,
                "Depth From (m)": ph.get('depth_from', 0),
                "Depth To (m)": ph.get('depth_to', 0),
                "Interval (m)": max(0, ph.get('depth_to', 0) - ph.get('depth_from', 0)),
                "Mud Type": normalize_mud_type(ph.get('mud_type', '')),
                "Mud Type Raw": ph.get('mud_type', ''),
                "Planned Days": ph.get('planned_days', 0),
                "Actual Days": ph.get('actual_days', 0),
                "Time Variance (days)": round(ph.get('planned_days', 0) - ph.get('actual_days', 0), 2),
                "Cost per Meter (INR)": ph.get('cost_per_meter', 0),
                "Cost per Barrel (INR)": ph.get('cost_per_barrel', 0),
                "Actual Cost (INR)": ph.get('actual_cost_inr', 0),
                "Volume Handled (bbl)": ph.get('volume_handled', 0),
            })
    return pd.DataFrame(records)


def build_complications_dataframe(comp_type="mud_loss"):
    """Build complications DataFrame for given type."""
    wells = load_wells_json()
    key = f"complications_{comp_type}"
    records = []
    for w in wells:
        for c in w.get(key, []):
            records.append({
                "Well Name": w['well_name'],
                "Asset": w.get('asset', ''),
                "Phase": c.get('phase', ''),
                "Date of Occurrence": c.get('date', ''),
                "Drill Depth (m)": c.get('drill_depth_m', 0),
                "Depth of Occurrence (m)": c.get('depth_occ_m', 0),
                "Mud System": c.get('mud_system', ''),
                "Operation in Brief": c.get('operation', ''),
                "Type of Loss/Stuck Up": c.get('type', ''),
                "Formation Info": c.get('formation', ''),
                "Layer": c.get('layer', ''),
                "Type of Pill/Action": c.get('pill_type', ''),
                "Mud Volume Lost (bbl)": c.get('volume_lost_bbl', 0),
                "NPT (Hrs)": c.get('npt_hrs', 0),
            })
    return pd.DataFrame(records) if records else pd.DataFrame()


def build_npt_summary_dataframe():
    """Build NPT summary per well."""
    wells = load_wells_json()
    records = []
    for w in wells:
        npt = w.get('npt', {})
        records.append({
            "Well Name": w['well_name'],
            "Asset": w.get('asset', ''),
            "Phase": "Overall",
            "Mud Loss (Hrs)": npt.get('mud_loss_hrs', 0),
            "Activity (Check-up Hrs)": npt.get('activity_hrs', 0),
            "Unplanned Waiting (Hrs)": npt.get('waiting_hrs', 0),
            "Stuck Up (Hrs)": npt.get('stuck_up_hrs', 0),
            "Total NPT (Hrs)": sum(npt.values()),
        })
    return pd.DataFrame(records)


def build_chemicals_dataframe():
    """Build chemical cost analysis DataFrame."""
    wells = load_wells_json()
    records = []
    for w in wells:
        for c in w.get('chemicals', []):
            records.append({
                "well_name": w['well_name'],
                "phase": c.get('phase', ''),
                "asset": w.get('asset', ''),
                "chemical_name": c.get('chemical_name', ''),
                "unit_size": c.get('unit_size', ''),
                "unit": c.get('unit', ''),
                "consumption_kg": c.get('consumption', 0),
                "actual_cost_inr": c.get('actual_cost_inr', 0),
                "price_per_unit_inr": c.get('price_per_unit_inr', 0),
            })
    return pd.DataFrame(records) if records else pd.DataFrame()


def build_cost_analysis_dataframe():
    """Build cost analysis table."""
    wells = load_wells_json()
    records = []
    for w in wells:
        row = {
            "Well Name": w['well_name'],
            "Asset": w.get('asset', ''),
            "Total Cost (INR)": w.get('total_cost_inr', 0),
            "Cost per Meter (INR)": round(w.get('cost_per_meter_inr', 0), 2),
            "Cost per Barrel (INR)": round(w.get('cost_per_barrel_inr', 0), 2),
        }
        for ph in w.get('phases', []):
            pname = ph.get('phase', '')
            if pname:
                row[f"{pname}_Cost_per_m"] = ph.get('cost_per_meter', 0)
                row[f"{pname}_Cost_per_bbl"] = ph.get('cost_per_barrel', 0)
                row[f"{pname}_Mud_Type"] = normalize_mud_type(ph.get('mud_type', ''))
        records.append(row)
    return pd.DataFrame(records)


def get_chemical_totals():
    """Get chemical consumption totals by phase for pie charts."""
    wells = load_wells_json()
    chem_by_name = {}
    for w in wells:
        for c in w.get('chemicals', []):
            name = c.get('chemical_name', '').strip().upper()
            phase = c.get('phase', 'Unknown')
            consumption = c.get('consumption', 0)

            # Normalize chemical name for grouping
            display_name = c.get('chemical_name', '').strip()
            if 'BENTONITE' in name:
                display_name = 'Bentonite'
            elif 'BARYTE' in name or 'BARITE' in name:
                display_name = 'Barite'
            elif 'XC POLYMER' in name:
                display_name = 'XC Polymer'
            elif 'PREGELAT' in name or 'POLYGEL' in name or 'STARCH' in name:
                display_name = 'Polygel (Pregelat. Starch)'
            elif 'PAC' in name and 'REGULAR' in name:
                display_name = 'PAC Regular'
            elif 'PAC' in name and 'LVG' in name:
                display_name = 'PAC LVG'
            elif 'CAUSTIC' in name:
                display_name = 'Caustic Soda'
            elif 'SODA ASH' in name:
                display_name = 'Soda Ash'
            elif 'POTASSIUM CHLORIDE' in name:
                display_name = 'Potassium Chloride'

            if display_name not in chem_by_name:
                chem_by_name[display_name] = {'total_kg': 0, 'phases': {}, 'unit_counts': {}, 'total_cost': 0}
            chem_by_name[display_name]['total_kg'] += consumption
            chem_by_name[display_name]['phases'][phase] = \
                chem_by_name[display_name]['phases'].get(phase, 0) + consumption

            # Track unit and cost
            unit = c.get('unit', '').strip().upper()
            if unit:
                chem_by_name[display_name]['unit_counts'][unit] = \
                    chem_by_name[display_name]['unit_counts'].get(unit, 0) + 1
            cost = c.get('actual_cost_inr', 0) or 0
            chem_by_name[display_name]['total_cost'] += cost

    # Determine primary unit for each chemical (most frequent)
    for chem_data in chem_by_name.values():
        uc = chem_data.get('unit_counts', {})
        if uc:
            chem_data['unit'] = max(uc, key=uc.get)
        else:
            chem_data['unit'] = ''
        del chem_data['unit_counts']

    return chem_by_name


def build_mud_parameters_dataframe():
    """Build mud parameters DataFrame with min/max/last ranges per phase per well."""
    wells = load_wells_json()
    records = []
    for w in wells:
        for mp in w.get('mud_parameters', []):
            rec = {
                "Well Name": w['well_name'],
                "Asset": w.get('asset', ''),
                "Field": w.get('field', ''),
                "Phase": mp.get('phase', ''),
                "Depth (m)": mp.get('depth', 0),
                "Mud System": mp.get('mud_system', ''),
                "Formation": mp.get('formation', ''),
                "Layer": mp.get('layer', ''),
                "Lithology": mp.get('lithology', ''),
            }
            # Add all mud parameter ranges
            MUD_PARAMS = [
                ('MW (PPG)', 'mud_weight_ppg'),
                ('FV (sec)', 'fv_sec'),
                ('PV (cP)', 'pv_cp'),
                ('YP (lb/100ft2)', 'yp_lb100ft2'),
                ('GEL0 (lb/100ft2)', 'gel0_lb100ft2'),
                ('GEL10 (lb/100ft2)', 'gel10_lb100ft2'),
                ('R6', 'r6'),
                ('R3', 'r3'),
                ('OWR Oil%', 'owr_oil_pct'),
                ('OWR Water%', 'owr_water_pct'),
                ('Solid%', 'solid_pct'),
                ('Chlorides (ppm)', 'chlorides_ppm'),
                ('HTHP F/L (mL)', 'hthp_fl_ml'),
                ('Ex. Lime (ppb)', 'ex_lime_ppb'),
                ('ES (V)', 'es_v'),
                ('WPS (ppm)', 'wps_ppm'),
                ('pH', 'ph'),
                ('FLT (°C)', 'flt_c'),
            ]
            for display_name, key_prefix in MUD_PARAMS:
                rec[f"{display_name} Min"] = mp.get(f"{key_prefix}_min", 0)
                rec[f"{display_name} Max"] = mp.get(f"{key_prefix}_max", 0)
                rec[f"{display_name} Last"] = mp.get(f"{key_prefix}_last", 0)
            records.append(rec)
    return pd.DataFrame(records) if records else pd.DataFrame()


def scan_for_new_wells(base_dir=WELL_CARDS_DIR):
    """Scan for new well card files not in current data."""
    wells = load_wells_json()
    known = {w['well_name'].upper().replace(' ', '').replace('-', '').replace('_', '') for w in wells}

    new_files = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if f.endswith('.xlsx') and not f.startswith('~$'):
                fp = os.path.join(root, f)
                # Simple name extraction
                name = f.replace('.xlsx', '').replace('well_card', '').replace('WELL CARD', '').strip()
                name_key = name.upper().replace(' ', '').replace('-', '').replace('_', '')
                if name_key not in known and 'calculation' not in f.lower():
                    new_files.append({'well_name': name, 'filepath': fp})

    return new_files[:20]  # Limit display


class WellCardWatcher:
    """File system watcher for new well card Excel files."""
    def __init__(self, watch_dir=WELL_CARDS_DIR, callback=None):
        self.watch_dir = watch_dir
        self.callback = callback
        self._running = False
        self._thread = None
        self._known_files = set()

    def start(self):
        self._known_files = set(
            f for f in glob.glob(os.path.join(self.watch_dir, "**", "*.xlsx"), recursive=True)
            if not os.path.basename(f).startswith("~$")
        )
        self._running = True
        self._thread = threading.Thread(target=self._watch_loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False

    def _watch_loop(self):
        while self._running:
            time.sleep(30)
            current = set(
                f for f in glob.glob(os.path.join(self.watch_dir, "**", "*.xlsx"), recursive=True)
                if not os.path.basename(f).startswith("~$")
            )
            new = current - self._known_files
            if new and self.callback:
                for f in new:
                    self.callback(f)
            self._known_files = current


# Cache
_cache = {}
_cache_time = {}
CACHE_TTL = 300


def get_cached(key, loader_fn):
    now = time.time()
    if key not in _cache or (now - _cache_time.get(key, 0)) > CACHE_TTL:
        _cache[key] = loader_fn()
        _cache_time[key] = now
    return _cache[key]


def invalidate_cache():
    _cache.clear()
    _cache_time.clear()


# Mud type colors for charts
MUD_TYPE_COLORS = {
    'OBM (SOBM)': '#E53935',
    'NDDF': '#1E88E5',
    'Gel Polymer': '#43A047',
    'KCl Polymer': '#FB8C00',
    'WBM': '#8E24AA',
    'Raw/Sea Water': '#00ACC1',
    'Unknown': '#9E9E9E',
}
