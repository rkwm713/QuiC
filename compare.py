"""compare.py – reusable engine that compares a SPIDAcalc
exchange JSON to a Katapult Pro job JSON.

Returns a Pandas DataFrame with:
    SCID, pole numbers, specs, existing/final loading %,
    Charter-drop flags, and simple match booleans.
"""

from __future__ import annotations
import json
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional

import pandas as pd

Coord = Tuple[float, float]              # (lat, lon) helper alias
EARTH_R = 6371000                        # metres – for overlap test

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _first_val(d: dict | None) -> Any | None:
    if isinstance(d, dict) and d:
        return next(iter(d.values()), None)
    return None

def _get_imported_val(d: dict | None) -> Any | None:
    """Extract value from Katapult attribute dict with -Imported key.
    
    Handles the pattern:
    "PL_number": {
        "-Imported": "actual_pole_number"  
    }
    
    Falls back to _first_val if -Imported key not found.
    """
    if isinstance(d, dict) and d:
        # First try the -Imported key specifically
        imported = d.get("-Imported")
        if imported is not None:
            return imported
        # Fall back to first value for compatibility
        return _first_val(d)
    return None

def _to_feet(raw) -> int | None:
    """Convert raw height to whole feet (int).

    Accepts:
        • dicts from SPIDA JSON, e.g. {"unit":"METRE","value":16.764}
        • numeric (assumed feet already)
        • strings like "50", "45'" or "13.7" (metres ⇒ feet)
    """
    if raw is None:
        return None

    # 1) Dict object with explicit unit/value --------------------------------
    if isinstance(raw, dict):
        unit = str(raw.get("unit", "")).lower()
        val = raw.get("value")
        try:
            val_f = float(val)
        except (TypeError, ValueError):
            return None

        if unit.startswith("met"):
            return int(round(val_f / 0.3048))
        # for anything else (FOOT / BLANK) assume already feet
        return int(round(val_f))

    # 2) Already numeric -----------------------------------------------------
    if isinstance(raw, (int, float)):
        return int(round(float(raw)))

    # 3) String parsing ------------------------------------------------------
    s = str(raw).strip()
    if s.isdigit():
        return int(s)
    if "'" in s:
        return int(s.split("'")[0])
    try:
        # assume metres string if contains decimal or unable above
        return int(round(float(s) / 0.3048))
    except ValueError:
        return None

def _fmt_pct(val: float | None) -> str | None:
    if val is None:
        return None
    if val > 1.01:  # already percent units (e.g. 65.4)
        return f"{val:.2f}%"
    return f"{val*100:.2f}%"

def _coords_from_spida_location(loc: dict) -> Optional[Coord]:
    """Extract coordinates from SPIDA location object.

    The newer SPIDAcalc JSON exports embed location coordinates in a GeoJSON
    object at ``location["geographicCoordinate"]["coordinates"]`` or
    sometimes ``location["mapLocation"]["coordinates"]``.  These are given in
    the standard GeoJSON order of ``[longitude, latitude]``.

    Older exports may still expose flat ``latitude``/``longitude`` keys (or the
    short forms ``lat``/``lon``).  This helper now supports all of these
    variants.
    """
    # 1. Simple flat keys ----------------------------------------------------
    lat = loc.get("latitude") or loc.get("lat")
    lon = loc.get("longitude") or loc.get("lon") or loc.get("long")
    if lat and lon:
        try:
            return (float(lat), float(lon))
        except (ValueError, TypeError):
            pass  # fall through to GeoJSON parsing

    # 2. GeoJSON coordinate arrays ------------------------------------------
    for key in ("geographicCoordinate", "mapLocation"):
        geo = loc.get(key)
        if isinstance(geo, dict):
            coords = geo.get("coordinates")
            if (isinstance(coords, (list, tuple)) and len(coords) == 2):
                lon_val, lat_val = coords  # GeoJSON order is [lon, lat]
                try:
                    return (float(lat_val), float(lon_val))
                except (ValueError, TypeError):
                    continue
    # If nothing worked, return None
    return None

def _coords_from_kat_node(node: dict) -> Optional[Coord]:
    """Extract coordinates from Katapult node object."""
    attrs = node.get("attributes", {})
    
    # Safely extract latitude
    lat_attr = attrs.get("latitude", {})
    if isinstance(lat_attr, dict):
        lat = lat_attr.get("-Imported")
    else:
        lat = lat_attr if lat_attr else None
    
    # Safely extract longitude
    lon_attr = attrs.get("longitude", {})
    if isinstance(lon_attr, dict):
        lon = lon_attr.get("-Imported")
    else:
        lon = lon_attr if lon_attr else None
    
    try:
        return (float(lat), float(lon)) if lat and lon else None
    except (ValueError, TypeError):
        return None

def _haversine_m(p1: Coord, p2: Coord) -> float:
    """Calculate distance between two coordinates in meters using Haversine formula."""
    from math import radians, sin, cos, sqrt, atan2
    lat1, lon1 = map(radians, p1)
    lat2, lon2 = map(radians, p2)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1)*cos(lat2)*sin(dlon/2)**2
    return 2 * EARTH_R * atan2(sqrt(a), sqrt(1-a))

# ---------------------------------------------------------------------------
# added helper
# ---------------------------------------------------------------------------

def _digits_only(txt: str | None) -> str | None:
    """Strip everything except 0-9; keep leading zeros."""
    if txt is None:
        return None
    digits = ''.join(ch for ch in str(txt) if ch.isdigit())
    return digits or None

# ---------------------------------------------------------------------------
# helpers for spatial fallback matching
# ---------------------------------------------------------------------------

from math import inf  # added for _nearest_scid


def _nearest_scid(sp_coord: Coord | None, kat_dict: Dict[str, dict], max_dist_m: float = 5.0) -> Optional[str]:
    """Return the Katapult SCID whose coordinate lies within *max_dist_m* metres of
    *sp_coord* – or *None* if no candidate is close enough.*"""
    if not sp_coord:
        return None

    nearest: Optional[str] = None
    best = max_dist_m  # keep *strictly* best so first within radius wins

    for k_scid, row in kat_dict.items():
        k_coord: Coord | None = row.get("Katapult Coord")
        if not k_coord:
            continue
        dist = _haversine_m(sp_coord, k_coord)
        if dist < best:
            best = dist
            nearest = k_scid
    return nearest

# ---------------------------------------------------------------------------
# main compare
# ---------------------------------------------------------------------------
def compare(spida_path: Path | str, kat_path: Path | str) -> pd.DataFrame:
    """Return DataFrame with merged comparison."""
    spida_path = Path(spida_path)
    kat_path = Path(kat_path)

    # ---------------- load SPIDA ----------------
    with spida_path.open("r", encoding="utf-8") as f:
        spida = json.load(f)

    sp_rows: List[dict] = []
    scid_counter = 0
    sp_charter_scids: set[str] = set()

    for lead in spida["leads"]:
        for loc in lead["locations"]:
            scid_counter += 1
            scid = f"{scid_counter:03d}"
            # Handle different pole label formats
            label_parts = loc["label"].split("-", 1)
            if len(label_parts) > 1:
                pole_num = label_parts[1]
            else:
                pole_num = loc["label"]  # Use the full label if no dash found

            measured    = next(d for d in loc["designs"] if d["layerType"] == "Measured")
            recommended = next(d for d in loc["designs"] if d["layerType"] == "Recommended")

            pole_struct = recommended["structure"]["pole"]
            sp_spec = _build_spida_spec(pole_struct) or ""

            # loading %
            def _get_load(design):
                for case in design.get("analysis", []):
                    for res in case.get("results", []):
                        if res.get("component") == "Pole":
                            return _fmt_pct(res.get("actual"))
                return None

            sp_exist = _get_load(measured)
            sp_final = _get_load(recommended)

            # Charter drop flag
            charter = any(
                item.get("owner", {}).get("industry") == "COMMUNICATION" and
                item.get("owner", {}).get("id") == "Charter" and
                item.get("clientItem", {}).get("type", "").lower().endswith("drop")
                for item in recommended["structure"].get("attachments", [])
            )
            if charter:
                sp_charter_scids.add(scid)

            # Extract coordinates
            coord = _coords_from_spida_location(loc)

            sp_rows.append(
                {
                    "SCID": scid,
                    "SPIDA Pole #": pole_num,
                    "SPIDA Spec": sp_spec,
                    "SPIDA Existing %": sp_exist,
                    "SPIDA Final %": sp_final,
                    "SPIDA Charter Drop": charter,
                    "SPIDA Coord": coord
                }
            )

    # ---------------- load Katapult ----------------
    with kat_path.open("r", encoding="utf-8") as f:
        kat = json.load(f)

    # Collect all birthmarks from the JSON
    birthmarks = {}
    _collect_birthmarks(kat, birthmarks)

    connections = kat.get("connections", {})
    section_to_conn: Dict[str, dict] = {}
    for conn in connections.values():
        for sid in conn.get("sections", {}):
            section_to_conn[sid] = conn

    kat_rows_by_scid: Dict[str, dict] = {}
    kat_scid_set: set[str] = set()
    kat_charter_scids: set[str] = set()
    kat_com_drop_scids: set[str] = set()  # Track poles with ANY service locations

    # First pass: Find all service locations and map them to poles
    # Method 1: Check service location nodes
    for node_id, node in kat["nodes"].items():
        attrs = node["attributes"]
        
        # Check if this is a Service Location node
        if attrs.get("node_type", {}).get("button_added") == "Service Location":
            owner = attrs.get("node_sub_type", {}).get("-Imported")
            
            for sec_id, measured in attrs.get("measured_attachments", {}).items():
                conn = section_to_conn.get(sec_id)
                if conn:
                    # Find the pole this service location is connected to
                    pole_node = conn["node_id_1"] if conn["node_id_2"] == node_id else conn["node_id_2"]
                    if pole_node in kat["nodes"]:
                        pole_attrs = kat["nodes"][pole_node]["attributes"]
                        pole_scid_raw = _first_val(pole_attrs.get("scid"))
                        pole_scid = pole_scid_raw if pole_scid_raw and pole_scid_raw.isdigit() else None
                        
                        if pole_scid:
                            # This pole has a service location (com drop)
                            kat_com_drop_scids.add(pole_scid)
                            
                            # If it's Charter and not measured, add to Charter list
                            if owner == "Charter" and measured is False:
                                kat_charter_scids.add(pole_scid)
    
    # Method 2: Check service drop connections
    for conn_id, conn in connections.items():
        conn_attrs = conn.get("attributes", {})
        conn_type = conn_attrs.get("connection_type", {}).get("button_added")
        
        if conn_type == "service drop":
            # node_id_2 is typically the pole with service drop
            pole_node_id = conn.get("node_id_2")
            if pole_node_id and pole_node_id in kat["nodes"]:
                pole_attrs = kat["nodes"][pole_node_id]["attributes"]
                pole_scid_raw = _first_val(pole_attrs.get("scid"))
                pole_scid = pole_scid_raw if pole_scid_raw and pole_scid_raw.isdigit() else None
                
                if pole_scid:
                    kat_com_drop_scids.add(pole_scid)

    # Second pass: Process main pole data
    # (constant set of node types we accept as actual poles)
    ALLOWED_NODE_TYPES: set[str] = {"pole", "Power", "Power Transformer", "Joint", "Joint Transformer"}
    for node_id, node in kat["nodes"].items():
        attrs = node["attributes"]
        scid_raw = _first_val(attrs.get("scid"))
        scid = scid_raw if scid_raw and scid_raw.isdigit() else None

        # Skip anything that isn't one of our allowed pole-type nodes
        node_type_attr = attrs.get("node_type")
        if isinstance(node_type_attr, dict):
            node_type_val = node_type_attr.get("button_added") or _first_val(node_type_attr)
        else:
            node_type_val = node_type_attr
        if node_type_val and str(node_type_val) not in ALLOWED_NODE_TYPES:
            continue

        # collect main pole data
        if not scid:
            continue

        # Extract pole number using correct Katapult field names
        # Primary field: DLOC_number
        dloc_data = attrs.get('DLOC_number', {})
        dloc_number = _first_val(dloc_data) if dloc_data else None
        pole_num = None
        if dloc_number and dloc_number != 'N/A':
            # Check if it already starts with PL to avoid double prefix
            pole_num = dloc_number if str(dloc_number).startswith('PL') else f"PL{dloc_number}"
        
        # Fallback: pole_tag.tagtext  
        if not pole_num:
            pole_tag_data = attrs.get('pole_tag', {})
            pole_tag_inner = _first_val(pole_tag_data) if pole_tag_data else {}
            if isinstance(pole_tag_inner, dict):
                tagtext = pole_tag_inner.get('tagtext')
                if tagtext and tagtext != 'N/A':
                    # Check if it already starts with PL to avoid double prefix
                    pole_num = tagtext if str(tagtext).startswith('PL') else f"PL{tagtext}"
        
        # Extract pole spec from birthmarks
        kat_spec = None
        # Look for birthmark reference in node attributes
        birthmark_ref = None
        for key in attrs.keys():
            if 'birthmark' in key.lower() or 'spec' in key.lower():
                # Use _get_imported_val to handle Katapult's nested attribute structure
                birthmark_ref = _get_imported_val(attrs.get(key))
                if birthmark_ref:
                    break
        
        if birthmark_ref and isinstance(birthmark_ref, str) and birthmark_ref in birthmarks:
            spec_data = birthmarks[birthmark_ref]
            height = spec_data.get('height')
            klass = spec_data.get('class')
            species = spec_data.get('species')
            if height and klass and species:
                kat_spec = f"{height}'-{klass} {species}"
        
        # Fallback to old method if birthmark not found
        if not kat_spec:
            height_raw = _first_val(attrs.get("pole_height")) or _first_val(attrs.get("poleLength")) or _first_val(attrs.get("Height"))
            klass      = _first_val(attrs.get("pole_class")) or _first_val(attrs.get("Class"))
            species    = _first_val(attrs.get("pole_species")) or _first_val(attrs.get("Species"))
            feet = _to_feet(height_raw)
            kat_spec = f"{feet}'-{klass} {species}" if all([feet, klass, species]) else None

        ex_pct = _first_val(attrs.get("existing_capacity_%"))
        fi_pct = _first_val(attrs.get("final_passing_capacity_%"))

        # Extract coordinates
        coord = _coords_from_kat_node(node)

        # Build row once so we can map it by multiple keys (SCID and digits-only).
        row_data = {
            "Katapult SCID #": scid,  # expose raw Katapult SCID as optional visible column
            "Katapult Pole #": pole_num,
            "Katapult Spec": kat_spec,
            "Katapult Existing %": f"{ex_pct}%" if ex_pct else None,
            "Katapult Final %": f"{fi_pct}%" if fi_pct else None,
            "Katapult Charter Drop": scid in kat_charter_scids,
            "Com Drop?": "Yes" if scid in kat_com_drop_scids else "No",
            "Katapult Coord": coord
        }

        # primary mapping by SCID
        kat_rows_by_scid[scid] = row_data
        # secondary mapping: if SCID is digits of pole number, allow lookup by that too
        digits_key = _digits_only(scid)
        if digits_key:
            kat_rows_by_scid[digits_key] = row_data

        # Track the official SCID once (avoid dup keys from digits mapping)
        kat_scid_set.add(scid)

    # ---------------- prepare list comparisons ----------------
    # Collect all SCIDs and pole numbers
    spida_scids = [row["SCID"] for row in sp_rows]
    spida_pole_nums = [row["SPIDA Pole #"] for row in sp_rows if row["SPIDA Pole #"]]
    
    katapult_scids = list(kat_scid_set)
    katapult_pole_nums = [row["Katapult Pole #"] for row in kat_rows_by_scid.values() if row["Katapult Pole #"]]
    
    # Create summary comparison data
    scids_only_in_spida = set(spida_scids) - set(katapult_scids)
    scids_only_in_katapult = set(katapult_scids) - set(spida_scids)
    scids_in_both = set(spida_scids) & set(katapult_scids)
    
    poles_only_in_spida = set(spida_pole_nums) - set(katapult_pole_nums)
    poles_only_in_katapult = set(katapult_pole_nums) - set(spida_pole_nums)
    poles_in_both = set(spida_pole_nums) & set(katapult_pole_nums)

    # ---------------- merge into df ----------------
    merged_rows: List[dict] = []
    for sp in sp_rows:
        scid = sp["SCID"]
        sp_pole_num = sp.get("SPIDA Pole #")
        sp_coord = sp.get("SPIDA Coord")  # needed for coord fallback

        # Tier 1 — SCID ↔ SCID ------------------------------------------------
        kdat = kat_rows_by_scid.get(scid)

        # Tier 2 — pole-ID equality -----------------------------------------
        if not kdat:
            kdat = kat_rows_by_scid.get(sp_pole_num)
        if not kdat:
            kdat = kat_rows_by_scid.get(_digits_only(sp_pole_num))

        # Tier 3 — nearest coordinate within radius -------------------------
        match_scid = None
        if not kdat:
            match_scid = _nearest_scid(sp_coord, kat_rows_by_scid, max_dist_m=5.0)
            if match_scid:
                kdat = kat_rows_by_scid.get(match_scid)

        krow = kdat or {}

        row = {**sp, **krow}
        
        # Add missing columns with defaults
        if "Katapult Pole #" not in row:
            row["Katapult Pole #"] = None
        if "Katapult SCID #" not in row:
            row["Katapult SCID #"] = None
        if "Katapult Spec" not in row:
            row["Katapult Spec"] = None
        if "Katapult Existing %" not in row:
            row["Katapult Existing %"] = None
        if "Katapult Final %" not in row:
            row["Katapult Final %"] = None
        if "Katapult Charter Drop" not in row:
            row["Katapult Charter Drop"] = False
        if "Com Drop?" not in row:
            row["Com Drop?"] = "No"
        if "Katapult Coord" not in row:
            row["Katapult Coord"] = None
            
        # Add comparison columns
        row["Spec Match"] = row.get("SPIDA Spec") == row.get("Katapult Spec")
        row["Existing % Match"] = row.get("SPIDA Existing %") == row.get("Katapult Existing %")
        row["Final % Match"] = row.get("SPIDA Final %") == row.get("Katapult Final %")
        row["Charter Drop Match"] = row["SPIDA Charter Drop"] == row["Katapult Charter Drop"]
        
        # Add list comparison information
        scid = row["SCID"]
        pole_num = row.get("SPIDA Pole #")
        
        # SCID comparison status
        if scid in scids_in_both:
            row["SCID Status"] = "In Both"
        elif scid in scids_only_in_spida:
            row["SCID Status"] = "SPIDA Only"
        else:
            row["SCID Status"] = "Unknown"
            
        # Pole number comparison status
        if pole_num and pole_num in poles_in_both:
            row["Pole # Status"] = "In Both"
        elif pole_num and pole_num in poles_only_in_spida:
            row["Pole # Status"] = "SPIDA Only"
        elif pole_num:
            row["Pole # Status"] = "Unknown"
        else:
            row["Pole # Status"] = "No Pole #"
        
        # Flag if this row was matched purely by coordinate proximity
        row["Matched by Coord"] = bool(match_scid)
        
        merged_rows.append(row)

    # Add rows for Katapult-only SCIDs
    for scid in scids_only_in_katapult:
        krow = kat_rows_by_scid[scid]
        row = {
            "SCID": scid,
            "SPIDA Pole #": None,
            "SPIDA Spec": None,
            "SPIDA Existing %": None,
            "SPIDA Final %": None,
            "SPIDA Charter Drop": False,
            "SPIDA Coord": None,
            **krow,
            "Matched by Coord": False,
            "Spec Match": False,
            "Existing % Match": False,
            "Final % Match": False,
            "Charter Drop Match": False,
            "SCID Status": "Katapult Only",
            "Pole # Status": "Katapult Only" if krow.get("Katapult Pole #") else "No Pole #"
        }
        merged_rows.append(row)

    return pd.DataFrame(merged_rows)

# Export haversine function for use in other modules
def haversine_m(p1: Coord, p2: Coord) -> float:
    """Public interface for distance calculation."""
    return _haversine_m(p1, p2)

# ---------------------------------------------------------------------------
# spec builder helper
# ---------------------------------------------------------------------------

def _build_spida_spec(pole_struct: dict) -> str | None:
    """Return a human-friendly pole spec string.

    Priority order:
    1. ``clientItemAlias`` if present – formats like ``50-2`` or ``45``.
    2. Height / class / species fields in ``clientItem``.
    """
    pb = pole_struct.get("clientItem", {})
    species = pb.get("species") or ""

    # -- 1. Alias handling --------------------------------------------------
    alias = pole_struct.get("clientItemAlias")
    if isinstance(alias, str) and alias.strip():
        alias_clean = alias.strip().replace("\u2032", "'")  # convert prime char if present
        parts = alias_clean.split("-", 1)
        length_ft_raw = parts[0].rstrip("'")
        if length_ft_raw.isdigit():
            length_ft = length_ft_raw  # numeric feet length
            if len(parts) == 2 and parts[1]:
                pole_class = parts[1]
                alias_fmt = f"{length_ft}'-{pole_class}"
            else:
                alias_fmt = f"{length_ft}'"
            return f"{alias_fmt} {species}".strip()
        # If alias isn't strictly numeric first, fall back further below

    # -- 2. Height / class / species fallback -------------------------------
    height_raw = pb.get("height")
    if isinstance(height_raw, dict):
        height_val = height_raw.get("value")
    else:
        height_val = height_raw
    feet = _to_feet(height_val)

    pole_class = pb.get("classOfPole") or pb.get("class") or ""

    if feet is not None and pole_class and species:
        return f"{feet}'-{pole_class} {species}"
    if feet is not None and species:
        return f"{feet}' {species}"
    if pole_class and species:
        return f"{pole_class} {species}"
    return species or None

def _collect_birthmarks(obj, out):
    """Recursively collect all birthmark sections from Katapult JSON."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "birthmark":
                out.update(v)  # Use update to merge all birthmarks into one dict
            else:
                _collect_birthmarks(v, out)
    elif isinstance(obj, list):
        for el in obj:
            _collect_birthmarks(el, out)
