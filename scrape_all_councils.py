"""
Scrapes ALL 2026 local election results for England from the Democracy Club API.
Creates one Excel file per council (same format as Sandwell_Elections_2026.xlsx).
Updates 00_Master_Sources.xlsx with collection status and sources.

Run: python scrape_all_councils.py
Estimated time: 5-10 minutes (API rate limits permitting)
"""

import urllib.request, ssl, json, os, time, re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ROOT    = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"
API_BASE = "https://candidates.democracyclub.org.uk/api/next"

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def api_get(url):
    for attempt in range(5):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait = 10 * (2 ** attempt)
                print(f"\n    [429 rate limit] waiting {wait}s...", end=" ", flush=True)
                time.sleep(wait)
            else:
                raise
    raise RuntimeError(f"Failed after 5 retries: {url}")

# ── PARTY NORMALISATION ───────────────────────────────────────────────────────

def norm_party(raw):
    r = raw.strip().lower()
    if "labour"    in r: return "Labour"
    if "green"     in r: return "Green"
    if "reform"    in r: return "Reform UK"
    if "conserv"   in r: return "Conservative"
    if "liberal dem" in r or r == "ld": return "Lib Dem"
    if "independent" in r: return "Independent"
    return "Other"

def bloc(party):
    if party in ("Labour","Green"):           return "Left"
    if party in ("Reform UK","Conservative"): return "Right"
    if party == "Lib Dem":                    return "Lib Dem"
    return "Other/Independent"

# ── EXCEL STYLING HELPERS ─────────────────────────────────────────────────────

BORDER = Border(*[Side(style="thin", color="CCCCCC")]*4)
PARTY_COLOURS = {
    "Reform UK":"12B6CF","Labour":"E4003B","Green":"02A95B",
    "Conservative":"0087DC","Lib Dem":"FAA61A","Independent":"888888","Other":"AAAAAA",
}
BLOC_COLOURS = {"Left":"C0392B","Right":"1A8CB0","Lib Dem":"FAA61A","Other/Independent":"888888"}

def style_header(ws):
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = "A2"

def style_rows(ws):
    alt = PatternFill("solid", fgColor="EBF1F8")
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        f = alt if i % 2 == 0 else PatternFill()
        for cell in row:
            cell.fill = f
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")

def auto_width(ws, max_w=40):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(w + 4, max_w)

def slug_to_folder(council_slug):
    """Convert API slug (e.g. 'barking-and-dagenham') to folder name."""
    return council_slug.replace("-", "_").title().replace("_", " ")

# ── STEP 1: FETCH ALL BALLOTS ─────────────────────────────────────────────────

print("Fetching all 2026 ballot data from Democracy Club API...")
print("(This may take a few minutes)\n")

all_ballots = []
url = f"{API_BASE}/ballots/?election_date=2026-05-07&limit=200"
page = 1

while url:
    print(f"  Page {page}...", end=" ", flush=True)
    try:
        data = api_get(url)
        all_ballots.extend(data["results"])
        print(f"{len(data['results'])} ballots (total so far: {len(all_ballots)})")
        url  = data.get("next")
        page += 1
        time.sleep(1.0)
    except Exception as e:
        print(f"ERROR: {e}")
        break

print(f"\nTotal ballots fetched: {len(all_ballots)}")

# ── STEP 2: GROUP BY COUNCIL ──────────────────────────────────────────────────

councils = {}
for b in all_ballots:
    eid   = b["election"]["election_id"]          # e.g. local.sandwell.2026-05-07
    slug  = eid.replace("local.", "").replace(".2026-05-07", "")  # e.g. sandwell
    name  = b["election"]["name"].replace(" local election", "").replace(" Local Election", "")
    ward  = b["post"]["label"]
    seats = b.get("winner_count", 1)

    candidates = []
    for c in b.get("candidacies", []):
        res = c.get("result") or {}
        candidates.append({
            "Ward":      ward,
            "Candidate": c.get("person", {}).get("name", "Unknown"),
            "Party":     c.get("party_name", "Unknown"),
            "Party Group": norm_party(c.get("party_name", "")),
            "Votes":     res.get("num_ballots") or 0,
            "Elected":   "Yes" if res.get("elected") else "No",
        })

    if slug not in councils:
        councils[slug] = {"name": name, "slug": slug, "wards": {}}
    councils[slug]["wards"][ward] = {"candidates": candidates, "seats": seats}

print(f"Councils found: {len(councils)}")

# ── STEP 3: CREATE EXCEL FOR EACH COUNCIL ────────────────────────────────────

MASTER_PATH = os.path.join(ROOT, "00_Master_Sources.xlsx")
master_updates = {}

for slug, council in councils.items():
    name   = council["name"]
    wards  = council["wards"]

    # Build flat dataframe
    rows = []
    for ward_name, wd in wards.items():
        for c in wd["candidates"]:
            c["Ward"] = ward_name
            rows.append(c)

    if not rows:
        continue

    df = pd.DataFrame(rows)
    df["Bloc"] = df["Party Group"].apply(bloc)

    # ── Sheet 1: Raw Results ──────────────────────────────────────────────────
    raw = df[["Ward","Candidate","Party","Votes","Elected"]].sort_values(
        ["Ward","Votes"], ascending=[True,False]).reset_index(drop=True)

    # ── Sheet 2: Party Totals ─────────────────────────────────────────────────
    grand_total = df["Votes"].sum()
    party_totals = (df.groupby("Party Group")["Votes"].sum()
                      .rename("Total Votes").reset_index()
                      .sort_values("Total Votes", ascending=False))
    if grand_total > 0:
        party_totals["% of All Votes"] = (party_totals["Total Votes"]/grand_total*100).round(2)
    else:
        party_totals["% of All Votes"] = 0
    seats = (df[df["Elected"]=="Yes"].groupby("Party Group").size()
               .rename("Seats Won").reset_index())
    party_totals = party_totals.merge(seats, on="Party Group", how="left")
    party_totals["Seats Won"] = party_totals["Seats Won"].fillna(0).astype(int)

    # ── Sheet 3: Blocs ────────────────────────────────────────────────────────
    bloc_totals = (df.groupby("Bloc")["Votes"].sum()
                     .rename("Total Votes").reset_index()
                     .sort_values("Total Votes", ascending=False))
    if grand_total > 0:
        bloc_totals["% of All Votes"] = (bloc_totals["Total Votes"]/grand_total*100).round(2)
    else:
        bloc_totals["% of All Votes"] = 0

    # ── Sheet 4: Ward Analysis ────────────────────────────────────────────────
    ward_records = []
    for ward_name, wd in wards.items():
        wdf = df[df["Ward"]==ward_name]
        tot = wdf["Votes"].sum()

        def grp_votes(g):
            return wdf[wdf["Party Group"]==g]["Votes"].sum()

        reform = grp_votes("Reform UK")
        labour = grp_votes("Labour")
        green  = grp_votes("Green")
        con    = grp_votes("Conservative")
        libdem = grp_votes("Lib Dem")
        ind    = grp_votes("Independent")
        other  = grp_votes("Other")
        left   = labour + green
        right  = reform + con

        elected = wdf[wdf["Elected"]=="Yes"]["Party Group"].value_counts().to_dict()
        split   = (left > reform) and (elected.get("Reform UK",0) > 0)

        ward_records.append({
            "Ward":                    ward_name,
            "Reform UK Votes":         reform,
            "Labour Votes":            labour,
            "Green Votes":             green,
            "Left (Lab+Green) Votes":  left,
            "Conservative Votes":      con,
            "Right (Ref+Con) Votes":   right,
            "Lib Dem Votes":           libdem,
            "Independent Votes":       ind,
            "Other Votes":             other,
            "Total Candidate Votes":   tot,
            "Reform % of Votes":       round(reform/tot*100,1) if tot else 0,
            "Labour % of Votes":       round(labour/tot*100,1) if tot else 0,
            "Green % of Votes":        round(green/tot*100,1)  if tot else 0,
            "Left % of Votes":         round(left/tot*100,1)   if tot else 0,
            "Right % of Votes":        round(right/tot*100,1)  if tot else 0,
            "Reform Seats":            elected.get("Reform UK",0),
            "Labour Seats":            elected.get("Labour",0),
            "Green Seats":             elected.get("Green",0),
            "Conservative Seats":      elected.get("Conservative",0),
            "Lib Dem Seats":           elected.get("Lib Dem",0),
            "Independent Seats":       elected.get("Independent",0),
            "Split Vote Evident?":     "YES" if split else "no",
        })

    ward_df = pd.DataFrame(ward_records)

    # ── Sheet 5: Flourish Data ────────────────────────────────────────────────
    flourish_df = ward_df[[
        "Ward","Reform UK Votes","Labour Votes","Green Votes","Conservative Votes",
        "Lib Dem Votes","Left (Lab+Green) Votes","Right (Ref+Con) Votes",
        "Reform % of Votes","Labour % of Votes","Green % of Votes",
        "Left % of Votes","Right % of Votes",
        "Reform Seats","Labour Seats","Green Seats","Conservative Seats","Split Vote Evident?"
    ]].copy()

    # ── Save Excel ────────────────────────────────────────────────────────────
    # Create folder (try name from our COUNCILS list, fall back to slug)
    folder_name = name  # democracy club name
    folder_path = os.path.join(ROOT, folder_name)
    if not os.path.exists(folder_path):
        # Try slug-based alternatives
        alt = slug.replace("-"," ").title()
        alt_path = os.path.join(ROOT, alt)
        if os.path.exists(alt_path):
            folder_path = alt_path
            folder_name = alt
        else:
            os.makedirs(folder_path, exist_ok=True)

    safe_name = re.sub(r'[<>:"/\\|?*]', '', name)
    excel_path = os.path.join(folder_path, f"{safe_name}_Elections_2026.xlsx")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        raw.to_excel(writer,         sheet_name="1. Raw Results",    index=False)
        party_totals.to_excel(writer,sheet_name="2. Party Totals",   index=False)
        bloc_totals.to_excel(writer, sheet_name="3. Left-Right Blocs",index=False)
        ward_df.to_excel(writer,     sheet_name="4. Ward Analysis",   index=False)
        flourish_df.to_excel(writer, sheet_name="5. Flourish Data",   index=False)

    # Style
    wb = load_workbook(excel_path)
    ws1 = wb["1. Raw Results"]
    style_header(ws1); style_rows(ws1); auto_width(ws1)
    for row in ws1.iter_rows(min_row=2):
        pg = norm_party(row[2].value or "")
        col = PARTY_COLOURS.get(pg)
        if col:
            row[2].fill = PatternFill("solid", fgColor=col)
            row[2].font = Font(color="FFFFFF", bold=True)
        if row[4].value == "Yes":
            row[4].fill = PatternFill("solid", fgColor="00B050")
            row[4].font = Font(color="FFFFFF", bold=True)

    for sh in ["2. Party Totals","3. Left-Right Blocs","4. Ward Analysis","5. Flourish Data"]:
        ws = wb[sh]
        style_header(ws); style_rows(ws); auto_width(ws)

    # Colour party column in sheet 2
    ws2 = wb["2. Party Totals"]
    for row in ws2.iter_rows(min_row=2):
        col = PARTY_COLOURS.get(row[0].value)
        if col:
            row[0].fill = PatternFill("solid", fgColor=col)
            row[0].font = Font(color="FFFFFF", bold=True)

    # Colour bloc column in sheet 3
    ws3 = wb["3. Left-Right Blocs"]
    for row in ws3.iter_rows(min_row=2):
        col = BLOC_COLOURS.get(row[0].value)
        if col:
            row[0].fill = PatternFill("solid", fgColor=col)
            row[0].font = Font(color="FFFFFF", bold=True)

    # Highlight split vote in sheet 4
    ws4 = wb["4. Ward Analysis"]
    split_idx = None
    for cell in ws4[1]:
        if cell.value == "Split Vote Evident?":
            split_idx = cell.column; break
    if split_idx:
        for row in ws4.iter_rows(min_row=2):
            cell = row[split_idx-1]
            if cell.value == "YES":
                cell.fill = PatternFill("solid", fgColor="FF0000")
                cell.font = Font(color="FFFFFF", bold=True)

    wb.save(excel_path)

    api_url = f"https://candidates.democracyclub.org.uk/elections/local.{slug}.2026-05-07/"
    master_updates[name] = {
        "url":       api_url,
        "collected": "Yes",
        "wards":     len(wards),
        "candidates":len(df),
        "total_votes":grand_total,
    }
    print(f"  OK {name:35s} -- {len(wards)} wards, {len(df)} candidates")

# ── STEP 4: UPDATE MASTER SOURCES EXCEL ──────────────────────────────────────

print("\nUpdating 00_Master_Sources.xlsx...")
try:
    df_master = pd.read_excel(MASTER_PATH, sheet_name="Council Sources")
    for i, row in df_master.iterrows():
        cname = str(row["Council"])
        for api_name, info in master_updates.items():
            if cname.lower() in api_name.lower() or api_name.lower() in cname.lower():
                df_master.at[i, "Results Page URL"]  = info["url"]
                df_master.at[i, "Data Collected"]     = info["collected"]
                df_master.at[i, "Notes"] = f"{info['wards']} wards, {info['candidates']} candidates, {info['total_votes']:,} total votes"
                break

    with pd.ExcelWriter(MASTER_PATH, engine="openpyxl", mode="w") as writer:
        df_master.to_excel(writer, sheet_name="Council Sources", index=False)
    print("  Master sources updated.")
except Exception as e:
    print(f"  Could not update master: {e}")

print(f"\n{'='*60}")
print(f"DONE. Excel files created for {len(master_updates)} councils.")
print(f"Location: {ROOT}")
