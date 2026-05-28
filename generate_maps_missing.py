"""
Generates maps for the 26 councils skipped in generate_maps.py.
Uses ONS December 2022 ward boundaries (FeatureServer) as fallback,
with fuzzy name matching to 2026 election ward names.
"""

import os, glob, re, ssl, urllib.request, json, time, statistics
import pandas as pd
import folium
from shapely.geometry import shape, mapping

ROOT     = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"
COUNCILS = os.path.join(ROOT, "Councils")

PARTY_COLOURS = {
    "Reform UK":    "#12B6CF",
    "Labour":       "#E4003B",
    "Green":        "#02A95B",
    "Conservative": "#0087DC",
    "Lib Dem":      "#FAA61A",
    "Independent":  "#888888",
    "Other":        "#AAAAAA",
}

ONS_BASE = (
    "https://services1.arcgis.com/ESMARspQHYMw9BZ9/arcgis/rest/services/"
    "Wards_December_2022_Boundaries_UK_BGC/FeatureServer/0/query"
)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode    = ssl.CERT_NONE

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

def clean(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def fetch_ons_wards(lad_name, retries=3):
    """Fetch all 2022 ward boundaries for a given LAD name from ONS."""
    # ONS uses different LAD name formats; try the council name as-is
    where = f"LAD22NM = '{lad_name}'"
    url = (f"{ONS_BASE}?where={urllib.request.quote(where)}"
           f"&outFields=WD22CD,WD22NM,LAD22CD,LAD22NM"
           f"&outSR=4326&f=geojson&resultRecordCount=200")
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
                return json.loads(r.read())
        except Exception as e:
            if attempt == retries - 1:
                return None
            time.sleep(2 * (attempt + 1))

def fuzzy_match(name26, names22_list):
    """Return best matching 2022 name for a 2026 ward name."""
    c26   = clean(name26)
    best_name  = None
    best_score = 0
    for n22 in names22_list:
        c22 = clean(n22)
        # Jaccard similarity on character bigrams
        b26 = set(c26[i:i+2] for i in range(len(c26)-1)) if len(c26) > 1 else set(c26)
        b22 = set(c22[i:i+2] for i in range(len(c22)-1)) if len(c22) > 1 else set(c22)
        if not b26 or not b22:
            continue
        score = len(b26 & b22) / len(b26 | b22)
        if score > best_score:
            best_score = score
            best_name  = n22
    return best_name, best_score

# Known LAD name overrides (2026 name → 2022 name in ONS data)
LAD_OVERRIDES = {
    "City of Lincoln":              "Lincoln",
    "London Borough of Tower Hamlets": "Tower Hamlets",
    "North Tyneside Council":       "North Tyneside",
    "West Northamptonshire":        "Northampton",  # approx fallback
    "Newport":                      None,   # Wales - skip
    "Powys":                        None,   # Wales - skip
}

SKIPPED = [
    "Barnsley", "Bradford", "Calderdale", "City of Lincoln", "Coventry",
    "East Surrey", "Gateshead", "Kingston_upon_Hull", "Kirklees",
    "London Borough of Tower Hamlets", "Milton_Keynes", "Newcastle_upon_Tyne",
    "Newport", "North Tyneside Council", "Sandwell", "Sefton", "Solihull",
    "South_Tyneside", "Sunderland", "Swindon", "Thurrock", "Wakefield",
    "Walsall", "West Northamptonshire", "West Surrey",
]

ok = err = 0

for folder_name in SKIPPED:
    # Resolve folder path
    folder = os.path.join(COUNCILS, folder_name)
    if not os.path.exists(folder):
        # Try county subfolder
        folder = os.path.join(COUNCILS, "County Councils", folder_name)
    if not os.path.exists(folder):
        print(f"  [SKIP] {folder_name}: folder not found")
        err += 1
        continue

    excel_files = glob.glob(os.path.join(folder, "*_Elections_2026.xlsx"))
    if not excel_files:
        print(f"  [SKIP] {folder_name}: no Excel")
        err += 1
        continue

    # Load election results
    df = pd.read_excel(excel_files[0], sheet_name="1. Raw Results")
    df["Party Group"] = df["Party"].apply(norm_party)

    winners = {}
    for ward, grp in df.groupby("Ward"):
        elected = grp[grp["Elected"] == "Yes"]
        if not elected.empty:
            top = elected.sort_values("Votes", ascending=False).iloc[0]
        else:
            top = grp.sort_values("Votes", ascending=False).iloc[0]
        winners[ward] = top["Party Group"]

    # Determine LAD name to use for ONS query
    display_name = folder_name.replace("_", " ")
    lad_name     = LAD_OVERRIDES.get(display_name, display_name)

    if lad_name is None:
        print(f"  [SKIP] {display_name}: Wales/other - skipping")
        err += 1
        continue

    # Fetch ONS 2022 boundaries
    time.sleep(0.5)
    geojson = fetch_ons_wards(lad_name)
    if not geojson or not geojson.get("features"):
        # Try stripping "Council" etc.
        alt = lad_name.replace(" Council","").replace(" Borough","").replace(" City","").strip()
        if alt != lad_name:
            geojson = fetch_ons_wards(alt)

    if not geojson or not geojson.get("features"):
        print(f"  [SKIP] {display_name}: ONS returned no features for '{lad_name}'")
        err += 1
        continue

    features_22 = geojson["features"]
    names22     = {f["properties"]["WD22NM"]: f for f in features_22}

    # Match 2026 ward names → 2022 ward names
    features_out = []
    unmatched    = []
    THRESHOLD    = 0.35

    for ward26, party in winners.items():
        # Exact match first
        if ward26 in names22:
            f22 = names22[ward26]
        else:
            best, score = fuzzy_match(ward26, list(names22.keys()))
            if score < THRESHOLD:
                unmatched.append(ward26)
                continue
            f22 = names22[best]

        geom   = f22["geometry"]
        colour = PARTY_COLOURS.get(party, "#AAAAAA")

        # Vote detail tooltip
        ward_df     = df[df["Ward"] == ward26].sort_values("Votes", ascending=False)
        elected_str = ", ".join(
            f"{r['Candidate']} ({r['Party Group']}, {int(r['Votes']):,})"
            for _, r in ward_df[ward_df["Elected"] == "Yes"].iterrows()
        ) or "N/A"

        features_out.append({
            "type": "Feature",
            "geometry": geom,
            "properties": {
                "ward":    ward26,
                "winner":  party,
                "colour":  colour,
                "elected": elected_str,
            }
        })

    if not features_out:
        print(f"  [SKIP] {display_name}: no wards matched (tried {len(winners)} wards)")
        err += 1
        continue

    fc = {"type": "FeatureCollection", "features": features_out}

    # Centre
    all_coords = []
    for f in features_out:
        g = f["geometry"]
        coords = g.get("coordinates", [])
        if g["type"] == "MultiPolygon":
            for poly in coords:
                all_coords.extend(poly[0])
        elif g["type"] == "Polygon":
            all_coords.extend(coords[0])

    lat = statistics.median(c[1] for c in all_coords)
    lon = statistics.median(c[0] for c in all_coords)

    # Legend
    seen = sorted({f["properties"]["winner"] for f in features_out})
    legend = """<div style="position:fixed;bottom:30px;left:30px;z-index:9999;
                background:white;padding:10px 14px;border-radius:8px;
                border:1px solid #ccc;font-family:Arial,sans-serif;font-size:13px;
                box-shadow:2px 2px 6px rgba(0,0,0,.2)"><b>Winning party</b><br>"""
    for p in seen:
        c = PARTY_COLOURS.get(p, "#AAAAAA")
        legend += f'<span style="display:inline-block;width:14px;height:14px;background:{c};margin-right:6px;vertical-align:middle;border-radius:2px"></span>{p}<br>'
    legend += "</div>"

    note = ""
    if unmatched:
        note = f"""<div style="position:fixed;bottom:30px;right:30px;z-index:9999;
                background:#fff3cd;padding:8px 12px;border-radius:8px;
                border:1px solid #ffc107;font-family:Arial;font-size:12px;max-width:260px">
                <b>Note:</b> {len(unmatched)} new ward(s) not shown (boundary review 2026):<br>
                {', '.join(unmatched[:6])}{'...' if len(unmatched)>6 else ''}
                </div>"""

    m = folium.Map(location=[lat, lon], zoom_start=12, tiles="CartoDB positron")

    folium.GeoJson(
        fc,
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

    title = f"""<div style="position:fixed;top:10px;left:50%;transform:translateX(-50%);
                z-index:9999;background:white;padding:8px 18px;border-radius:8px;
                border:1px solid #ccc;font-family:Arial;font-size:15px;font-weight:bold;
                box-shadow:2px 2px 6px rgba(0,0,0,.2)">
                {display_name} — 2026 Local Elections
                {'<br><span style=\"font-size:11px;color:#888;font-weight:normal\">(approx. 2022 ward boundaries)</span>' if unmatched else ''}
                </div>"""

    m.get_root().html.add_child(folium.Element(title))
    m.get_root().html.add_child(folium.Element(legend))
    if note:
        m.get_root().html.add_child(folium.Element(note))

    maps_dir = os.path.join(folder, "maps")
    os.makedirs(maps_dir, exist_ok=True)
    safe = folder_name.replace(" ", "_")
    out  = os.path.join(maps_dir, f"{safe}_Ward_Map.html")
    m.save(out)

    matched = len(features_out)
    total   = len(winners)
    suffix  = f" (unmatched: {unmatched[:3]})" if unmatched else ""
    print(f"  [OK]   {display_name:45s} {matched}/{total} wards mapped{suffix}")
    ok += 1

print(f"\nDone. {ok} additional maps created, {err} skipped.")
