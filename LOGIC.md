**SCID**

SPIDAcalc:

In SPIDAcalc’s “exchange” JSON there isn’t a field called scid tucked away in each pole object—instead the SCID is simply the position of that pole in the project’s structures array (zero-padded to three digits).  In other words:

1. Open your JSON (often under something like project.structures or spidadevices.structures).
2. Loop through that array with an index.
3. Your SCID is (index + 1) formatted as a three-digit string (e.g. "001", "002", …).

Here’s a quick Python example:

```python
import json

# load your SPIDAcalc JSON
with open("spidacalc_project.json") as f:
    data = json.load(f)

# adjust this path if your JSON nests it differently
structures = data["project"]["structures"]

for idx, pole in enumerate(structures):
    scid = str(idx + 1).zfill(3)
    pole_id = pole.get("id", "<no-id>")
    print(f"Pole ID: {pole_id}  →  SCID: {scid}")
```

If you ever need to match back to Katapult or flag mismatches, use that generated scid string to join your datasets.

Katapult:

In a Katapult Pro export the “SCID” for each pole lives in the nodes object under each node’s attributes.scid["-Imported"].  Here’s the pattern:

1. Open your JSON and grab nodes

import json

with open("job.json") as f:

kat = json.load(f)

nodes = kat["nodes"]

1. Iterate through each node and read the SCID

for node_id, node in nodes.items():

attrs = node.get("attributes", {})

# primary SCID field

scid_field = attrs.get("scid", {})

# the "-Imported" key holds the actual SCID string

scid = scid_field.get("-Imported")

# if there’s no scid or it’s a reference (e.g. "002.A"), you can skip or flag:

if scid and scid.isdigit():

print(f"Node {node_id} → SCID {scid}")

else:

# either missing or a reference (like "002.A")

print(f"Node {node_id} → no main SCID (got {scid})")

1. Filtering out non-main poles
If you only want “main” poles, require scid to be all digits (e.g. "001", "002"), and skip ones ending in letters ("002.A", etc.).

That will print each Katapult node’s SCID exactly as Katapult imported it.

**Pole Number**

SPIDAcalc:

In a SPIDAcalc “exchange” JSON, the pole number is simply the id field of each structure object.  You’ll find all of the poles in the array at something like data["project"]["structures"], and each entry’s .id is the Pole Number you see in the UI (and what SPIDAcalc will match when you upload).

Here’s how you can pull them out in Python:

```python
import json

# load your SPIDAcalc JSON
with open("spidacalc_project.json") as f:
    data = json.load(f)

# adjust this path if your JSON nests it differently
structures = data["project"]["structures"]

for pole in structures:
    # the pole number is the 'id' property
    pole_number = pole["id"]
    print(f"Pole Number: {pole_number}")
```

If for some reason your import used another field, you can also fall back on externalId:

```python
pole_number = pole.get("id") or pole.get("externalId")
```

But in most SPIDAexchange files it’s always the .id on each structure.

Katapult:

In a Katapult Pro export the “Pole Number” isn’t its JSON key but lives alongside the SCID in each node’s attributes.  You’ll usually see it as either PL_number["-Imported"] or (in some versions) PoleNumber["-Imported"].

Here’s the pattern:

import json

# 1) load your Katapult job JSON

with open("job_export.json") as f:

kat = json.load(f)

# 2) grab the nodes dict

nodes = kat["nodes"]

# 3) iterate and pull out pole numbers

for node_id, node in nodes.items():

attrs = node.get("attributes", {})

# first try the PL_number field…

pole_num_field = attrs.get("PL_number", {})

pole_number = pole_num_field.get("-Imported")

# …or fall back to the PoleNumber field

if not pole_number:

pole_num_field = attrs.get("PoleNumber", {})

pole_number = pole_num_field.get("-Imported")

print(f"Node {node_id} → Pole Number: {pole_number}")

Notes:

- If you only want “main” poles, you can skip any nodes where pole_number is missing or doesn’t match your expected pattern.
- You can then join this to your SCID (from attrs["scid"]["-Imported"]) to build a full map of Katapult → SPIDAcalc identifiers.

**Pole Spec**

Here’s how SPIDAexchange encodes the “pole spec” and how you can pull it out and format it as

`50’-2 Southern Pine`

# **1. Where to find it in the JSON**

Each pole lives in the array at

```json
data["project"]["structures"]
```

Within each structure you’ll usually see both a measuredDesign and a recommendedDesign block.  The pole spec you want is under one of those, e.g.:

```json
"recommendedDesign": {
  "pole": {
    "length": 15.85,           // in meters
    "class": "2",              // class code
    "species": "Southern Pine" // timber species
  },
  …  
}
```

*(If your JSON uses feet instead of meters, you’ll see “length” already in feet.)*

# **2. Converting & formatting in Python**

1. Load your JSON
2. Drill in to structures `→ recommendedDesign → pole`
3. Convert `meters → total feet (if needed)`
4. Split into feet/inches (optional—your example omits inches)
5. Assemble the string:

{feet}’-{class} {species}

```python
import json

# 1) load
with open("spidacalc_project.json") as f:
    data = json.load(f)

# 2) grab your poles array
structures = data["project"]["structures"]

for struct in structures:
    pole_spec = struct["recommendedDesign"]["pole"]
    
    # 3) convert meters → feet
    length_m = pole_spec["length"]
    total_feet = length_m / 0.3048
    
    # 4) split into whole feet and inches (if you want the inches)
    feet = int(total_feet)
    inches = int(round((total_feet - feet) * 12))
    
    # 5) pull class & species
    class_code = pole_spec["class"]
    species    = pole_spec["species"]
    
    # format two ways:
    # a) just feet & class & species
    spec_simple = f"{feet}’-{class_code} {species}"
    # b) with inches included
    spec_full   = f"{feet}’-{inches}\"-{class_code} {species}"
    
    print(spec_simple)   # → 50’-2 Southern Pine
    # print(spec_full)    # → 50’-2"-2 Southern Pine
```

- If your JSON already stores length in feet, skip the / 0.3048 step and just use total_feet = pole_spec["length"].
- Swap in struct["measuredDesign"]["pole"] if you need the existing spec instead of the recommended one.

That will give you exactly the “50’-2 Southern Pine”–style output you’re after.

Katapult:

First, you’ll need to see exactly what your Katapult JSON calls the “length”, “class” and “species” fields.  The quickest way is to dump out the attribute-keys for one of your pole nodes:

import json

with open("your_katapult_export.json") as f:

kat = json.load(f)

# peek at the first node’s attribute keys

first_node = next(iter(kat["nodes"].values()))

print(first_node["attributes"].keys())

Look in that list for things like

- "poleLength" or "Height"
- "poleClass" or "Class"
- "poleSpecies" or "Species"

Once you know the exact names, here’s a recipe you can adapt:

import json

import re

def to_feet_inches(raw):

"""

Turn either a string "50'-2" or a metric value in string/number form

into a (feet, inches) tuple.

"""

if isinstance(raw, str) and "'" in raw:

# e.g. "50'-2" or "50'-2\""

m = re.match(r"(\d+)'-?(\d+)", raw)

return int(m.group(1)), int(m.group(2))

else:

# assume meters if numeric or numeric‐string

m = float(raw)

total_ft = m / 0.3048

ft = int(total_ft)

inch = int(round((total_ft - ft) * 12))

return ft, inch

# --- load Katapult JSON ---

with open("your_katapult_export.json") as f:

kat = json.load(f)

for node_id, node in kat["nodes"].items():

attrs = node.get("attributes", {})

# — adjust these to whatever your keys actually are —

raw_len     = attrs.get("poleLength",     {}).get("-Imported") \

or attrs.get("Height",         {}).get("-Imported")

pole_class  = attrs.get("poleClass",      {}).get("-Imported")

species     = attrs.get("poleSpecies",    {}).get("-Imported")

# convert to feet/inches

feet, inches = to_feet_inches(raw_len)

# format as “50’-2 Southern Pine”

spec = f"{feet}’-{pole_class} {species}"

print(f"Node {node_id}: {spec}")

How it works:

1. Discover your field names by printing out attributes.keys().
2. Extract the raw length, class and species via attrs.get(...).
3. Convert the raw length to feet & inches (handles both a pre-formatted string or a metric number).
4. Assemble the string with f"{feet}’-{class} {species}".

That will give you exactly the same "50’-2 Southern Pine"-style pole spec for every pole in your Katapult JSON.

**Loading percentages**

SPIDAcalc:

To get both the existing (measured) and final (recommended) loading percentages out of your SPIDAexchange JSON and display them as xx.xx%, you’ll generally do two things:

1. Locate the loading values in the JSON
2. Format them as percentages in Python

# **1. Locating “existing” vs. “final” loading in the JSON**

In a SPIDAexchange file you’ll have a top‐level array called analysisAssets (sometimes named analysisResults), which contains one asset per design layer.  Each asset has a "structures" list of results—one per pole—where each result object includes an "actual" field:

```json
{

"analysisAssets": [

{

"designName": "Measured Design",            // your “existing” design

"structures": [

{ "structureId": "PL398491", "actual": 0.6534, "allowable": 1.0 },

{ "structureId": "PL398492", "actual": 0.8121, "allowable": 1.0 },

…

]

},

{

"designName": "Recommended Design",         // your “final” design

"structures": [

{ "structureId": "PL398491", "actual": 0.7123, "allowable": 1.0 },

{ "structureId": "PL398492", "actual": 0.8857, "allowable": 1.0 },

…

]

}

],

"project": {

"structures": [ /* your pole list, same order */ ]

}

}
```

Here, actual is “the loading percentage or safety factor of that component under that analysis case”  .

# **2. Pulling & formatting in Python**

Once you’ve loaded your JSON, pick out the two assets (measured vs. recommended), then for each pole index extract .actual and format it:

```python
import json

# 1) Load JSON

with open("spidacalc_project.json") as f:

data = json.load(f)

# 2) Grab structures & assets

structures   = data["project"]["structures"]

assets       = data["analysisAssets"]

measured     = next(a for a in assets if a["designName"] == "Measured Design")

recommended  = next(a for a in assets if a["designName"] == "Recommended Design")

# 3) Loop & format

for idx, pole in enumerate(structures):

pole_id           = pole["id"]

existing_loading  = measured["structures"][idx]["actual"]

final_loading     = recommended["structures"][idx]["actual"]

# 4) Format as xx.xx% (multiplies fraction by 100 and appends %)

existing_pct = f"{existing_loading:.2%}"

final_pct    = f"{final_loading:.2%}"

print(f"Pole {pole_id}: Existing Loading = {existing_pct}, Final Loading = {final_pct}")

- f"{value:.2%}" will multiply value by 100, round to two decimal places, and add a % sign—exactly the xx.xx% format you want .
- If your JSON’s actual is already in percent form (e.g. 65.34 for 65.34 %), swap to .2f and append %:

existing_pct = f"{existing_loading:.2f}%"
```

This approach cleanly gives you both the existing and final loading for every pole in your SPIDAcalc project.

Katapult:

In your exported JSON each pole (or “node”) has an attributes object that holds a bunch of the make-ready data.  In there you’ll find two fields that map IDs to percentage strings:

- existing_capacity_%

"existing_capacity_%": {

"<attributeId>": "95.35"

}

- This is your existing loading percentage (in this example, 95.35 %)
- final_passing_capacity_%

"final_passing_capacity_%": {

"<attributeId>": "95.28"

}

- This is your final loading percentage (in this example, 95.28 %)

**Example (Python)**

# assuming `data` is your parsed JSON and

# that your poles are top-level keys or under data["nodes"]

for node_id, node in data.items():

attrs = node.get("attributes", {})

existing_dict = attrs.get("existing_capacity_%", {})

final_dict    = attrs.get("final_passing_capacity_%", {})

# pull out the one percentage value in each:

existing_pct = next(iter(existing_dict.values()), None)

final_pct    = next(iter(final_dict.values()),    None)

print(f"Pole {node_id}: existing={existing_pct}%, final={final_pct}%")

Just swap in whatever key under which your poles actually live (e.g. data["nodes"] or directly in data).

**Com Drop?**

SPIDAcalc:

In your SPIDAexchange JSON, each pole’s “recommended” layer lives under

```json
data["project"]["structures"][…]["recommendedDesign"]
```

and within that there’s an "attachments" (or sometimes called "items") array. To pull out only the COMMUNICATION drops owned by Charter, you can do something like this:

```python
import json

# 1) load your exchange file

with open("spidacalc_project.exchange.spida") as f:

data = json.load(f)

# 2) grab the list of poles

structures = data["project"]["structures"]

# 3) loop through poles and filter attachments

for pole in structures:

pole_id = pole["id"]

rec = pole.get("recommendedDesign", {})

attachments = rec.get("attachments", [])

# find Charter drops

charter_drops = [

att for att in attachments

if att.get("owner", {}).get("industry") == "COMMUNICATION"

and att.get("owner", {}).get("id") == "Charter"

and att.get("clientItem", {}).get("type", "").lower().endswith("drop")

]

for drop in charter_drops:

drop_type    = drop["clientItem"]["type"]

drop_height  = drop.get("attachmentHeight", "<no height>")

print(f"Pole {pole_id}: {drop_type} at {drop_height} owned by Charter")
```

**How it works**

1. structures
Your poles live in an array at data["project"]["structures"].
2. recommendedDesign.attachments
Under each pole:

```python
rec = pole.get("recommendedDesign", {})

attachments = rec.get("attachments", [])
```

Filter by owner + type
We pick out attachments where

```python
att["owner"]["industry"] == "COMMUNICATION"

and att["owner"]["id"]       == "Charter"
```

1. and whose clientItem.type ends with “Drop” (so you only get service‐drops).
2. Output
For each matching drop we print the pole’s id, the drop’s type, and its attachment height.

You can adapt that list comprehension to collect these into a structure or Excel sheet instead of printing, but that’s the core pattern.

Katapult:

In Katapult’s Pro JSON, service‐drop (“Com Drop”) attachments live as their own Service Location nodes and are flagged in each node’s measured_attachments attribute:

1. Identify Service Location nodes
Every drop shows up as a node whose

attributes.node_type["button_added"] == "Service Location"

1. (i.e. it’s a “Service Location” icon on the map)
2. Look at its measured_attachments map
In that same node’s attributes you’ll find

"measured_attachments": {

"<sectionId1>": true,

"<sectionId2>": false,

…

}

1. – a boolean per attachment. A value of false means “not yet measured” (i.e. proposed)
2. (Optional) Filter by owner
If you only want Charter’s drops, check

attributes.node_sub_type["-Imported"] == "Charter"

1. on the same node
2. Tie back to the pole
The keys in measured_attachments are section‐IDs. To see which pole it’s on, find the connection in your top‐level
data["connections"][connId]["sections"]
that contains that sectionId, then look at that connection’s node_id_1 / node_id_2.

**Example in Python**

# assume `data` is your parsed JSON (e.g. from “ryantest123.json”)

proposed_com_drops = []

for node_id, node in data["nodes"].items():

attrs = node.get("attributes", {})

# 1) only Service Location nodes

if attrs.get("node_type", {}).get("button_added") != "Service Location":

continue

owner = attrs.get("node_sub_type", {}).get("-Imported")

for section_id, measured in attrs.get("measured_attachments", {}).items():

# 2) false → proposed

if not measured:

# find which connection includes this section

for conn_id, conn in data["connections"].items():

if section_id in conn.get("sections", {}):

# pole is the other node on that connection

pole_id = (conn["node_id_1"] if conn["node_id_2"] == node_id

else conn["node_id_2"])

proposed_com_drops.append({

"pole_node": pole_id,

"drop_node": node_id,

"connection_id": conn_id,

"section_id": section_id,

"owner": owner

})

break

# now `proposed_com_drops` holds all your proposed (Com) drops

This will give you every proposed communication service drop—filter further on owner=="Charter" if you only want Charter’s.