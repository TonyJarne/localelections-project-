"""
Generates an interactive HTML map for each local authority showing
ward-level 2026 election results coloured by winning party.

Output: Councils/{CouncilName}/maps/{CouncilName}_Ward_Map.html

Requirements: pip install folium shapely pandas openpyxl
"""

import os, io, glob, re, ssl, urllib.request, json, time
import pandas as pd
from shapely import wkt
from shapely.geometry import mapping
import folium

ROOT     = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"
COUNCILS = os.path.join(ROOT, "Councils")

# ── Party colours (hex, no #) ─────────────────────────────────────────────────
PARTY_COLOURS = {
    "Reform UK":    "#12B6CF",
    "Labour":       "#E4003B",
    "Green":        "#02A95B",
    "Conservative": "#0087DC",
    "Lib Dem":      "#FAA61A",
    "Independent":  "#888888",
    "Other":        "#AAAAAA",
}

def norm_party(raw):
    if not raw: return "Other"
    r = str(raw).strip().lower()
    if "labour"      in r: return "Labour"
    if "green"       in r: return "Green"
    if "reform"      in r: return "Reform UK"
    if "conserv"     in r: return "Conservative"
    if "liberal dem" in r or r == "ld": return "Lib Dem"
    if "independent" in r: return "Independent"
    return "Other"

# ── Download / cache source data ──────────────────────────────────────────────

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode    = ssl.CERT_NONE

def http_get(url, retries=3):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
                return r.read()
        except Exception as e:
            if i == retries-1: raise
            time.sleep(3 * (i+1))

CACHE_DIR  = os.path.join(ROOT, "_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

def cached_csv(url, fname):
    path = os.path.join(CACHE_DIR, fname)
    if not os.path.exists(path):
        print(f"  Downloading {fname}...")
        data = http_get(url)
        with open(path, "wb") as f:
            f.write(data)
    return pd.read_csv(path, low_memory=False)

print("Loading ward boundary data...")
planning_df = cached_csv(
    "https://files.planning.data.gov.uk/dataset/ward.csv",
    "planning_wards.csv"
)
# Drop wards with no geometry
planning_df = planning_df[planning_df["geometry"].notna()].copy()
print(f"  {len(planning_df):,} wards with geometry")

print("Loading ONS 2026 ward lookup...")
lookup_df = cached_csv(
    "https://hub.arcgis.com/api/v3/datasets/7447015a1f2f4332807d7341a636f95d_0/downloads/data?format=csv&spatialRefId=4326&where=1%3D1",
    "ons_ward_lookup_2026.csv"
)
print(f"  {len(lookup_df):,} ward-LAD mappings")

# Build: reference -> LAD name
code_to_lad = lookup_df.set_index("WD26CD")["LAD26NM"].to_dict()
# Build: planning entity -> LAD name (via majority vote)
planning_df["LAD26NM"] = planning_df["reference"].map(code_to_lad)

# ── Helper: normalise ward name for fuzzy matching ────────────────────────────
def clean(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

# ── Build geometry lookup: (lad_name, ward_name) -> shapely geometry ──────────
print("Building geometry index...")
geo_index = {}   # key: (clean_lad, clean_ward)  value: geometry WKT

# For wards with a known LAD (matched via 2026 code)
for _, row in planning_df[planning_df["LAD26NM"].notna()].iterrows():
    key = (clean(row["LAD26NM"]), clean(row["name"]))
    geo_index[key] = row["geometry"]

# For wards WITHOUT a matched LAD, index by (org_entity, ward_name)
org_index = {}   # key: (org_entity, clean_ward)  value: geometry WKT
for _, row in planning_df[planning_df["LAD26NM"].isna()].iterrows():
    key = (str(row["organisation-entity"]), clean(row["name"]))
    org_index[key] = row["geometry"]

print(f"  Geometry index: {len(geo_index):,} entries")

# ── Build org_entity lookup from organisations ────────────────────────────────
# Use the planning LAD-matched wards to find which org_entity = which LAD
lad_to_entity = {}
for _, row in planning_df[planning_df["LAD26NM"].notna()].iterrows():
    lad  = row["LAD26NM"]
    ent  = str(row["organisation-entity"])
    lad_to_entity.setdefault(lad, {})
    lad_to_entity[lad][ent] = lad_to_entity[lad].get(ent, 0) + 1

# Pick the most common entity per LAD
lad_to_entity = {lad: max(ents, key=ents.get) for lad, ents in lad_to_entity.items()}

# ── Process each council ──────────────────────────────────────────────────────

def get_winner(ward_results):
    """Return the row with most votes that is elected, else most votes."""
    elected = ward_results[ward_results["Elected"] == "Yes"]
    if not elected.empty:
        top = elected.sort_values("Votes", ascending=False).iloc[0]
    else:
        top = ward_results.sort_values("Votes", ascending=False).iloc[0]
    return top["Party Group"] if "Party Group" in top.index else norm_party(top["Party"])

def make_map(council_folder, council_name):
    # Find Excel file
    excel_files = glob.glob(os.path.join(council_folder, "*_Elections_2026.xlsx"))
    if not excel_files:
        return False, "no Excel"

    # Load election results
    try:
        df = pd.read_excel(excel_files[0], sheet_name="1. Raw Results")
    except Exception as e:
        return False, f"Excel read error: {e}"

    if "Party" not in df.columns or "Ward" not in df.columns:
        return False, "missing columns"

    # Normalise party
    df["Party Group"] = df["Party"].apply(norm_party)

    # Get winner per ward
    winners = {}
    for ward, grp in df.groupby("Ward"):
        winners[ward] = get_winner(grp)

    if not winners:
        return False, "no wards"

    # Try to find LAD name match in lookup
    lad_matches = lookup_df[lookup_df["LAD26NM"].str.lower() == council_name.lower().replace("_"," ")]
    lad_name = lad_matches["LAD26NM"].iloc[0] if not lad_matches.empty else council_name.replace("_"," ")
    org_entity = lad_to_entity.get(lad_name)

    # Build GeoJSON features
    features = []
    unmatched = []
    for ward_name, winner_party in winners.items():
        # Try exact code-based match first
        key = (clean(lad_name), clean(ward_name))
        geom_wkt = geo_index.get(key)

        # Fallback: try via org_entity
        if geom_wkt is None and org_entity:
            key2 = (org_entity, clean(ward_name))
            geom_wkt = org_index.get(key2)

        # Fallback: fuzzy search in planning_df by name only within org_entity wards
        if geom_wkt is None and org_entity:
            candidates = planning_df[
                planning_df["organisation-entity"].astype(str) == org_entity
            ]
            cname = clean(ward_name)
            scores = candidates["name"].apply(lambda n: len(set(clean(n)) & set(cname)) / max(len(clean(n)), len(cname), 1))
            if scores.max() > 0.7:
                best = candidates.loc[scores.idxmax()]
                geom_wkt = best["geometry"]

        if geom_wkt is None:
            unmatched.append(ward_name)
            continue

        try:
            geom = wkt.loads(geom_wkt)
        except Exception:
            unmatched.append(ward_name)
            continue

        # Get vote details for tooltip
        ward_df   = df[df["Ward"] == ward_name].sort_values("Votes", ascending=False)
        elected_r = ward_df[ward_df["Elected"] == "Yes"]
        elected_str = ", ".join(
            f"{r['Candidate']} ({r['Party Group']}, {int(r['Votes']):,})"
            for _, r in elected_r.iterrows()
        ) if not elected_r.empty else "N/A"

        colour = PARTY_COLOURS.get(winner_party, "#AAAAAA")
        features.append({
            "type": "Feature",
            "geometry": mapping(geom),
            "properties": {
                "ward":     ward_name,
                "winner":   winner_party,
                "colour":   colour,
                "elected":  elected_str,
                "votes_top": int(ward_df.iloc[0]["Votes"]) if len(ward_df) else 0,
            }
        })

    if not features:
        return False, f"no geometry matched (unmatched: {unmatched[:5]})"

    geojson = {"type": "FeatureCollection", "features": features}

    # Centre map on council
    import statistics
    all_coords = []
    for f in features:
        g = f["geometry"]
        if g["type"] == "MultiPolygon":
            for poly in g["coordinates"]:
                all_coords.extend(poly[0])
        elif g["type"] == "Polygon":
            all_coords.extend(g["coordinates"][0])
    if not all_coords:
        return False, "no coords"

    lat = statistics.median(c[1] for c in all_coords)
    lon = statistics.median(c[0] for c in all_coords)

    # Party legend HTML
    seen_parties = sorted({f["properties"]["winner"] for f in features})
    legend_html = """
    <div style="position:fixed;bottom:30px;left:30px;z-index:9999;
                background:white;padding:10px 14px;border-radius:8px;
                border:1px solid #ccc;font-family:Arial,sans-serif;font-size:13px;
                box-shadow:2px 2px 6px rgba(0,0,0,.2)">
      <b>Winning party</b><br>"""
    for p in seen_parties:
        c = PARTY_COLOURS.get(p, "#AAAAAA")
        legend_html += f'<span style="display:inline-block;width:14px;height:14px;background:{c};margin-right:6px;vertical-align:middle;border-radius:2px"></span>{p}<br>'
    legend_html += "</div>"

    # Build map
    m = folium.Map(location=[lat, lon], zoom_start=12, tiles="CartoDB positron")

    folium.GeoJson(
        geojson,
        name="Wards",
        style_function=lambda f: {
            "fillColor":   f["properties"]["colour"],
            "color":       "white",
            "weight":      1.5,
            "fillOpacity": 0.75,
        },
        tooltip=folium.GeoJsonTooltip(
            fields=["ward", "winner", "elected"],
            aliases=["Ward", "Winner", "Elected"],
            style="font-family:Arial;font-size:13px;"
        ),
    ).add_to(m)

    # Title
    title_html = f"""
    <div style="position:fixed;top:10px;left:50%;transform:translateX(-50%);z-index:9999;
                background:white;padding:8px 18px;border-radius:8px;
                border:1px solid #ccc;font-family:Arial;font-size:15px;font-weight:bold;
                box-shadow:2px 2px 6px rgba(0,0,0,.2)">
      {council_name.replace("_"," ")} — 2026 Local Elections
    </div>"""
    m.get_root().html.add_child(folium.Element(title_html))
    m.get_root().html.add_child(folium.Element(legend_html))

    # Save
    maps_dir = os.path.join(council_folder, "maps")
    os.makedirs(maps_dir, exist_ok=True)
    out = os.path.join(maps_dir, f"{council_name.replace(' ','_')}_Ward_Map.html")
    m.save(out)
    matched = len(features)
    total   = len(winners)
    return True, f"{matched}/{total} wards mapped" + (f" (unmatched: {unmatched[:3]})" if unmatched else "")


# ── Run for all councils ──────────────────────────────────────────────────────

# Collect all council folders (skip sub-groupings like County Councils)
council_folders = []
for entry in os.listdir(COUNCILS):
    full = os.path.join(COUNCILS, entry)
    if not os.path.isdir(full):
        continue
    # Check if it contains an Excel directly, or recurse one level (County Councils subfolder)
    if glob.glob(os.path.join(full, "*_Elections_2026.xlsx")):
        council_folders.append((entry, full))
    else:
        for sub in os.listdir(full):
            sub_full = os.path.join(full, sub)
            if os.path.isdir(sub_full) and glob.glob(os.path.join(sub_full, "*_Elections_2026.xlsx")):
                council_folders.append((sub, sub_full))

print(f"\nGenerating maps for {len(council_folders)} councils...\n")

ok = err = 0
for name, folder in sorted(council_folders):
    success, msg = make_map(folder, name)
    status = "OK" if success else "SKIP"
    print(f"  [{status}] {name:45s} {msg}")
    if success: ok += 1
    else:       err += 1

print(f"\nDone. {ok} maps created, {err} skipped.")
print(f"Maps saved to: Councils/{{council}}/maps/{{council}}_Ward_Map.html")
