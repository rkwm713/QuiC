"""
spida_writer.py – given a loaded SPIDA JSON, a SCID,
the column name, and the user's new value, patch the JSON
in-place so it's ready to dump back to disk.
"""

from typing import Tuple

def apply_edit(spida: dict, scid: str, column: str, new_val: str):
    """Mutate *spida* so that column on that SCID equals new_val."""
    scid_counter = 0
    
    for lead in spida.get("leads", []):
        for loc in lead.get("locations", []):
            scid_counter += 1
            if f"{scid_counter:03d}" != scid:
                continue

            # Find the recommended design
            rec = None
            for design in loc.get("designs", []):
                if design.get("layerType") == "Recommended":
                    rec = design
                    break
            
            if not rec:
                continue
                
            pole = rec.get("structure", {}).get("pole", {}).get("clientItem", {})

            if column in ("SPIDA Spec", "SPIDA Pole Spec"):
                _update_pole_spec(pole, new_val)
            elif column in ("SPIDA Existing %",):
                _set_loading(loc, "Measured", float(new_val.strip("%")) / 100)
            elif column in ("SPIDA Final %",):
                _set_loading(loc, "Recommended", float(new_val.strip("%")) / 100)
            elif column in ("SPIDA Charter Drop", "Com Drop? (SPIDA)"):
                _toggle_charter(rec, new_val.lower().startswith("t"))

            return  # once patched → done

def _update_pole_spec(pole: dict, new_val: str):
    """Parse and update pole specification string like "40' H1 Southern Pine"."""
    try:
        # Normalize prime character to plain apostrophe for easier splitting
        cleaned = new_val.replace("\u2032", "'")
        parts = cleaned.strip().split("'", 1)
        if len(parts) < 2:
            return
            
        feet = float(parts[0].strip())
        rest = parts[1].strip()
        
        if not rest:
            return
            
        # Split class and species
        rest_parts = rest.split(" ", 1)
        if len(rest_parts) < 2:
            return
            
        klass = rest_parts[0].strip()
        species = rest_parts[1].strip()
        
        # Update pole data
        if "height" not in pole:
            pole["height"] = {}
        pole["height"]["value"] = feet * 0.3048  # Convert feet to meters
        
        pole["classOfPole"] = klass.replace("-", "")
        pole["species"] = species
        
    except (ValueError, IndexError) as e:
        print(f"Error parsing pole spec '{new_val}': {e}")

def _set_loading(location_block: dict, layer_name: str, pct: float):
    """Set the loading percentage for a specific design layer."""
    design = None
    for d in location_block.get("designs", []):
        if d.get("layerType") == layer_name:
            design = d
            break
    
    if not design:
        return
        
    for case in design.get("analysis", []):
        for res in case.get("results", []):
            if res.get("component") == "Pole":
                res["actual"] = pct

def _toggle_charter(rec_design: dict, want: bool):
    """Add or remove Charter drop attachment based on want flag."""
    structure = rec_design.get("structure", {})
    atts = structure.setdefault("attachments", [])
    
    # Check if Charter drop already exists
    present = any(
        a.get("owner", {}).get("id") == "Charter"
        and a.get("clientItem", {}).get("type", "").lower().endswith("drop")
        for a in atts
    )
    
    if want and not present:
        # Add Charter drop
        atts.append({
            "owner": {"industry": "COMMUNICATION", "id": "Charter"},
            "clientItem": {"type": "ServiceDrop"},
            "attachmentHeight": 18.0
        })
    elif not want and present:
        # Remove Charter drop
        structure["attachments"] = [
            a for a in atts if not (
                a.get("owner", {}).get("id") == "Charter" and
                a.get("clientItem", {}).get("type", "").lower().endswith("drop"))
        ]
