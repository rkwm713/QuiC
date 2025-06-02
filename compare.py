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

        # Only convert if unit is actually metres
        if unit.startswith("met"):
            return int(round(val_f / 0.3048))
        else:
            # For FOOT or any other unit, assume already feet
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
    
    # Convert string to float if needed
    try:
        if isinstance(val, str):
            # Strip % sign if present and convert to float
            val_clean = val.strip().rstrip('%')
            val = float(val_clean)
        elif not isinstance(val, (int, float)):
            return None
    except (ValueError, TypeError):
        return None
        
    if val > 1.01:  # already percent units (e.g. 65.4)
        return f"{val:.2f}%"
    return f"{val*100:.2f}%"

def _coords_from_spida_location(loc: dict) -> Optional[Coord]:
    """
    Return (lat, lon) or None. Tries multiple sources:
    1. location.geographicCoordinate.coordinates   (GeoJSON order [lon, lat])
    2. location.mapLocation.coordinates            (same)
    3. flat latitude/longitude keys
    4. measured → geographicCoordinate fallback
    
    Includes heuristic coordinate swapping if values appear reversed.
    """
    def _geo_coords(block):
        """Extract coordinates from a GeoJSON-style block."""
        if isinstance(block, dict):
            coords = block.get("coordinates")
            if isinstance(coords, (list, tuple)) and len(coords) == 2:
                lon, lat = coords              # GeoJSON is lon,lat
                try:
                    return float(lat), float(lon)
                except (ValueError, TypeError):
                    pass
        return None

    # 1-2. Try recommended layer GeoJSON blocks first
    for key in ("geographicCoordinate", "mapLocation"):
        c = _geo_coords(loc.get(key))
        if c:
            return c

    # 3. Try flat latitude/longitude keys
    lat = loc.get("latitude") or loc.get("lat")
    lon = loc.get("longitude") or loc.get("lon") or loc.get("long")
    try:
        if lat and lon:
            lat_f, lon_f = float(lat), float(lon)
            # Heuristic swap if the file was exported in lat,lon order instead of GeoJSON lon,lat
            if abs(lat_f) < 5 and abs(lon_f) > 20:   # looks like lat,lon was swapped to lon,lat
                lat_f, lon_f = lon_f, lat_f
            return lat_f, lon_f
    except (TypeError, ValueError):
        pass

    # 4. Fallback to measured design if recommended layer missing coords
    designs = loc.get("designs", [])
    if designs:
        # Measured is typically index 0, but let's be explicit
        measured_design = None
        for design in designs:
            if design.get("layerType") == "Measured":
                measured_design = design
                break
        
        if measured_design:
            # Try structure.poleLocation first
            pole_location = (
                measured_design.get("structure", {})
                .get("poleLocation", {})
            )
            c = _geo_coords(pole_location)
            if c:
                return c
            
            # Also try the same paths as recommended layer
            for key in ("geographicCoordinate", "mapLocation"):
                struct_geo = measured_design.get("structure", {}).get(key)
                c = _geo_coords(struct_geo)
                if c:
                    return c

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
# advanced pole normalization and matching helpers
# ---------------------------------------------------------------------------

def _clean_digits(txt: str | None) -> str | None:
    """Extract digits only and strip leading zeros for SCID comparison."""
    if txt is None:
        return None
    digits = ''.join(ch for ch in str(txt) if ch.isdigit())
    if not digits:
        return None
    return digits.lstrip('0') or '0'  # preserve single '0' if all zeros

def _normalize_pole_num(pole_num: str | None) -> str | None:
    """Normalize pole number by extracting digits and stripping leading zeros."""
    if not pole_num:
        return None
    # Remove common prefixes and extract digits
    clean = str(pole_num).upper().replace('PL', '')
    digits = ''.join(ch for ch in clean if ch.isdigit())
    if not digits:
        return None
    return digits.lstrip('0') or '0'

def _extract_spec_components(spec: str | None) -> tuple[int | None, str | None, str | None]:
    """Extract height, class, species from a pole spec string like '45-3 Southern Pine'."""
    if not spec:
        return None, None, None
    
    spec = str(spec).strip()
    
    # Try to parse format like "45-3 Southern Pine" or "45' Southern Pine"
    import re
    
    # Pattern: height (with optional '), optional dash, class, species
    pattern = r"^(\d+)[']*[-\s]*([A-Z0-9]*)\s*(.*)$"
    match = re.match(pattern, spec, re.IGNORECASE)
    
    if match:
        height_str, class_str, species_str = match.groups()
        
        try:
            height = int(height_str)
        except ValueError:
            height = None
            
        pole_class = class_str.strip() if class_str.strip() else None
        species = species_str.strip() if species_str.strip() else None
        
        return height, pole_class, species
    
    return None, None, None

def _specs_match(spida_spec: str | None, kat_spec: str | None, height_tolerance_ft: int = 1) -> bool:
    """Compare two pole specs for compatibility within tolerance."""
    if not spida_spec or not kat_spec:
        return False
        
    sp_height, sp_class, sp_species = _extract_spec_components(spida_spec)
    kat_height, kat_class, kat_species = _extract_spec_components(kat_spec)
    
    # Height check (within tolerance)
    if sp_height is not None and kat_height is not None:
        if abs(sp_height - kat_height) > height_tolerance_ft:
            return False
    elif sp_height != kat_height:  # both None or one None
        return False
    
    # Class check (exact match when both present)
    if sp_class and kat_class:
        if sp_class.upper() != kat_class.upper():
            return False
    elif sp_class != kat_class:  # both None or one None
        return False
    
    # Species check is optional - we don't fail on species mismatch
    # but we could add logging here if needed
    
    return True

def _build_lookup_tables(kat_rows_by_scid: dict) -> tuple[dict, dict, dict]:
    """Build optimized lookup tables for different matching tiers."""
    scid_lookup = {}
    pole_num_lookup = {}
    coord_lookup = {}
    
    for scid, row_data in kat_rows_by_scid.items():
        # SCID lookup (clean digits)
        clean_scid = _clean_digits(scid)
        if clean_scid:
            scid_lookup[clean_scid] = row_data
        
        # Pole number lookup (normalized)
        pole_num = row_data.get("Katapult Pole #")
        norm_pole = _normalize_pole_num(pole_num)
        if norm_pole:
            pole_num_lookup[norm_pole] = row_data
        
        # Coordinate lookup (for spatial queries)
        coord = row_data.get("Katapult Coord")
        if coord:
            # Round to 6 decimal places for lookup key (~0.1m precision)
            coord_key = (round(coord[0], 6), round(coord[1], 6))
            coord_lookup[coord_key] = row_data
    
    return scid_lookup, pole_num_lookup, coord_lookup

def _find_closest_poles(sp_coord: Coord | None, kat_rows_by_scid: dict, max_dist_m: float = 5.0) -> list[tuple[str, dict, float]]:
    """Find all Katapult poles within max_dist_m, sorted by distance."""
    if not sp_coord:
        return []
    
    candidates = []
    for k_scid, row in kat_rows_by_scid.items():
        k_coord = row.get("Katapult Coord")
        if k_coord:
            dist = _haversine_m(sp_coord, k_coord)
            if dist <= max_dist_m:
                candidates.append((k_scid, row, dist))
    
    # Sort by distance
    return sorted(candidates, key=lambda x: x[2])

# ---------------------------------------------------------------------------
# main compare with tiered matching
# ---------------------------------------------------------------------------
def compare(spida_path: Path | str, kat_path: Path | str) -> pd.DataFrame:
    """Return DataFrame with merged comparison."""
    spida_path = Path(spida_path)
    kat_path = Path(kat_path)

    # ---------------- load SPIDA ----------------
    with spida_path.open("r", encoding="utf-8") as f:
        spida = json.load(f)

    # Quick sanity-check: verify Charter attachments are found
    owners = _owners_table(spida)
    test_hits = [
        (att.get("catalog", {}).get("code", "N/A"), att.get("usageGroup", "N/A"))
        for lead in spida["leads"]
        for loc  in lead["locations"]
        for des  in loc["designs"]
        if des["layerType"] == "Recommended"
        for att  in _iter_all_attachments(des["structure"])
        if "charter" in (att.get("owner", {}).get("id") or owners.get(att.get("ownerId",""), "")).lower()
    ]
    if test_hits:
        print(f"✅ {len(test_hits)} Charter attachments found in SPIDA file")
    else:
        print("⚠️  No Charter attachments found in SPIDA file")

    # ---------------- build SPIDA alias table ----------------
    alias_table = {}
    client_data = spida.get("clientData", {})
    poles = client_data.get("poles", [])
    
    for pole in poles:
        # Get pole specifications
        height_raw = pole.get("height", {})
        if isinstance(height_raw, dict):
            height_val = height_raw.get("value")
            if height_val:
                height_ft = round(height_val / 0.3048)  # Convert meters to feet
            else:
                height_ft = None
        else:
            height_ft = _to_feet(height_raw)
            
        pole_class = pole.get("classOfPole") or pole.get("class", "")
        species = pole.get("species", "")
        
        # Build the full spec string
        if height_ft and pole_class and species:
            full_spec = f"{height_ft}'-{pole_class} {species}"
        elif height_ft and species:
            full_spec = f"{height_ft}' {species}"
        elif pole_class and species:
            full_spec = f"{pole_class} {species}"
        else:
            full_spec = species or None
            
        # Map all aliases for this pole to the full spec
        for alias_obj in pole.get("aliases", []):
            alias_id = alias_obj.get("id")
            if alias_id and full_spec:
                alias_table[alias_id] = full_spec

    print(f"📋 Built alias table with {len(alias_table)} pole specifications")

    # ---------------- process SPIDA locations ----------------
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
            sp_spec = _build_spida_spec(pole_struct, alias_table) or ""

            # loading %
            def _get_load(design):
                for case in design.get("analysis", []):
                    for res in case.get("results", []):
                        if res.get("component") == "Pole":
                            return _fmt_pct(res.get("actual"))
                return None

            sp_exist = _get_load(measured)
            sp_final = _get_load(recommended)

            # Charter drop flag - comprehensive detection using helper functions
            charter = any(
                _is_charter_service(att, owners)
                for att in _iter_all_attachments(recommended["structure"])
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
        
        # Extract pole spec from attributes - try direct pole_spec first
        kat_spec = None
        
        # ⬇️ NEW: Direct check for pole_spec before birthmark logic
        spec_raw = _get_imported_val(attrs.get("pole_spec"))
        if spec_raw:              # already a finished spec like "45-3 Southern Pine"
            kat_spec = str(spec_raw)
        else:
            # ⬇️ EXISTING: Look for birthmark reference in node attributes
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
            "Katapult Existing %": _fmt_pct(ex_pct),
            "Katapult Final %": _fmt_pct(fi_pct),
            "Katapult Charter Drop": scid in kat_com_drop_scids,  # True if ANY service location exists
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

    # ---------------- build optimized lookup tables ----------------
    scid_lookup, pole_num_lookup, coord_lookup = _build_lookup_tables(kat_rows_by_scid)

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

    # ---------------- merge into df with tiered matching ----------------
    merged_rows: List[dict] = []
    match_stats = {
        'scid': 0,
        'pole_num': 0, 
        'coord_direct': 0,
        'coord_spec_verified': 0,
        'unmatched': 0
    }
    
    for sp in sp_rows:
        scid = sp["SCID"]
        sp_pole_num = sp.get("SPIDA Pole #")
        sp_coord = sp.get("SPIDA Coord")
        sp_spec = sp.get("SPIDA Spec")
        
        kdat = None
        match_tier = None
        match_distance = None
        
        # ==================== TIER 1: EXACT SCID MATCH ====================
        clean_spida_scid = _clean_digits(scid)
        if clean_spida_scid and clean_spida_scid in scid_lookup:
            kdat = scid_lookup[clean_spida_scid]
            match_tier = 'scid'
            match_stats['scid'] += 1
        
        # ==================== TIER 2: POLE NUMBER MATCH ====================
        if not kdat:
            norm_spida_pole = _normalize_pole_num(sp_pole_num)
            if norm_spida_pole and norm_spida_pole in pole_num_lookup:
                kdat = pole_num_lookup[norm_spida_pole]
                match_tier = 'pole_num'
                match_stats['pole_num'] += 1
        
        # ==================== TIER 3 & 4: COORDINATE + SPEC MATCHING ====================
        if not kdat and sp_coord:
            closest_poles = _find_closest_poles(sp_coord, kat_rows_by_scid, max_dist_m=5.0)
            
            for kat_scid, candidate_data, distance in closest_poles:
                # Tier 3a: Direct match if < 1m
                if distance < 1.0:
                    kdat = candidate_data
                    match_tier = 'coord_direct'
                    match_distance = distance
                    match_stats['coord_direct'] += 1
                    break
                
                # Tier 3b + 4: Candidate match (1-5m) requires spec verification
                elif distance <= 5.0:
                    kat_spec = candidate_data.get("Katapult Spec")
                    if _specs_match(sp_spec, kat_spec):
                        kdat = candidate_data
                        match_tier = 'coord_spec_verified'
                        match_distance = distance
                        match_stats['coord_spec_verified'] += 1
                        break
        
        # ==================== HANDLE UNMATCHED POLES ====================
        if not kdat:
            match_tier = 'unmatched'
            match_stats['unmatched'] += 1

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
            
        # ==================== ADD MATCH METADATA ====================
        row["Match Tier"] = match_tier
        row["Match Distance (m)"] = f"{match_distance:.2f}" if match_distance is not None else None
        
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
        
        # Legacy flag for backward compatibility
        row["Matched by Coord"] = match_tier in ['coord_direct', 'coord_spec_verified']
        
        merged_rows.append(row)

    # ==================== ADD KATAPULT-ONLY POLES ====================
    matched_katapult_scids = set()
    for row in merged_rows:
        if row.get("Katapult SCID #"):
            matched_katapult_scids.add(row["Katapult SCID #"])
    
    for scid in kat_scid_set:
        if scid not in matched_katapult_scids:
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
                "Match Tier": "katapult_only",
                "Match Distance (m)": None,
                "Matched by Coord": False,
                "Spec Match": False,
                "Existing % Match": False,
                "Final % Match": False,
                "Charter Drop Match": False,
                "SCID Status": "Katapult Only",
                "Pole # Status": "Katapult Only" if krow.get("Katapult Pole #") else "No Pole #"
            }
            merged_rows.append(row)
    
    # ==================== REPORT MATCH STATISTICS ====================
    total_spida_poles = len(sp_rows)
    total_matches = sum(match_stats[key] for key in ['scid', 'pole_num', 'coord_direct', 'coord_spec_verified'])
    match_rate = (total_matches / total_spida_poles * 100) if total_spida_poles > 0 else 0
    
    print(f"\n📊 Tiered Matching Results:")
    print(f"   🎯 Tier 1 (SCID): {match_stats['scid']} poles")
    print(f"   🏷️  Tier 2 (Pole #): {match_stats['pole_num']} poles")
    print(f"   📍 Tier 3a (Coord <1m): {match_stats['coord_direct']} poles")
    print(f"   🔍 Tier 3b+4 (Coord+Spec): {match_stats['coord_spec_verified']} poles")
    print(f"   ❌ Unmatched: {match_stats['unmatched']} poles")
    print(f"   ✅ Overall match rate: {match_rate:.1f}% ({total_matches}/{total_spida_poles})")

    return pd.DataFrame(merged_rows)

# Export haversine function for use in other modules
def haversine_m(p1: Coord, p2: Coord) -> float:
    """Public interface for distance calculation."""
    return _haversine_m(p1, p2)

# ---------------------------------------------------------------------------
# spec builder helper
# ---------------------------------------------------------------------------

def _build_spida_spec(pole_struct: dict, alias_table: dict) -> str | None:
    """Return a human-friendly pole spec string following the exact SPIDA algorithm.

    Priority order:
    1. ``clientItemAlias`` if present – resolved via alias_table OR parsed directly
    2. Height / class / species fields in ``clientItem``.
    """
    # -- 1. Alias handling with table lookup --------------------------------------------------
    alias = pole_struct.get("clientItemAlias")
    if alias and alias in alias_table:
        return alias_table[alias]  # Fast path: direct lookup from alias table
    
    # -- 2. Direct alias parsing (improved) --------------------------------------------------
    if isinstance(alias, str) and alias.strip():
        pb = pole_struct.get("clientItem", {})
        species = pb.get("species") or ""
        
        alias_clean = alias.strip().replace("\u2032", "'")  # convert prime char if present
        
        # Check if alias has the standard "XX-Y" format
        if "-" in alias_clean:
            parts = alias_clean.split("-", 1)
            height_part = parts[0].rstrip("'")
            class_part = parts[1] if len(parts) > 1 else ""
            
            if height_part.isdigit() and class_part:
                # Standard format like "45-3"
                return f"{height_part}′-{class_part} {species}".strip()
        
        # If not standard format, fall back to raw field construction

    # -- 3. Height / class / species construction from raw fields -------------------------------
    pb = pole_struct.get("clientItem", {})
    species = pb.get("species") or ""
    
    height_raw = pb.get("height")
    if isinstance(height_raw, dict):
        height_val = height_raw.get("value")
        unit = str(height_raw.get("unit", "")).lower()
        
        # Convert to feet using the fixed _to_feet function
        feet = _to_feet(height_raw)
    else:
        feet = _to_feet(height_raw)

    pole_class = pb.get("classOfPole") or pb.get("class") or ""

    # Build spec string in the standard format
    if feet is not None and pole_class and species:
        return f"{feet}′-{pole_class} {species}"
    elif feet is not None and species:
        return f"{feet}′ {species}"
    elif pole_class and species:
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

# ---------------------------------------------------------------------------
# Charter service drop detection helpers
# ---------------------------------------------------------------------------

# --- SPIDA look-ups that are sometimes needed -----------------
def _owners_table(spida_json: dict) -> dict[str,str]:
    """Map ownerId→ownerName (Charter, AT&T, etc.)."""
    owners = {}
    for lead in spida_json.get("leads", []):
        for o in lead.get("owners", []):
            owners[o["id"]] = o.get("name", o["id"])
    return owners

def _iter_all_attachments(struct: dict):
    """Attachments, wires *and* spans – v11 uses all three."""
    for key in ("attachments", "wires", "spans"):
        yield from struct.get(key, [])
    for node in struct.get("nodes", []):
        for key in ("attachments", "wires", "spans"):
            yield from node.get(key, [])

def _is_charter_service(att: dict, owners: dict[str,str]) -> bool:
    """Check if attachment is a Charter service drop using broad heuristics."""
    owner_id = att.get("owner", {}).get("id") or owners.get(att.get("ownerId",""), "")
    owner_id = owner_id.lower()
    if "charter" not in owner_id:
        return False                    # not Charter

    # broad service-drop heuristics
    ug   = str(att.get("usageGroup", "")).lower()
    ctyp = str(att.get("clientItem", {}).get("type", "")).lower()
    code = str(att.get("catalog", {}).get("code", "")).upper()

    return (
        "service" in ug                 # COMMUNICATION_SERVICE, …_SERVICE_DROP, etc.
        or ctyp.endswith("drop")        # older clientItem types
        or "FSV0250" in code            # Charter's 0.25-inch fibre
        or att.get("serviceDrop") is True  # v11 boolean flag
    )
