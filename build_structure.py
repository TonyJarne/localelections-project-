"""
Creates the full folder structure for localelections_project
and generates the master source Excel.
Run from: C:\\Users\\antoj\\OneDrive\\Escritorio\\localelections_project
"""
import os, re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ROOT = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"

# ── COMPLETE COUNCIL LIST ─────────────────────────────────────────────────────
# Source: opencouncildata.co.uk/elections.php  (verified May 2026)

COUNCILS = [
    # name, type, seats, control
    # County Councils
    ("East Sussex",             "County",       50,  "Reform UK"),
    ("Essex",                   "County",       78,  "Reform UK"),
    ("Hampshire",               "County",       78,  "Conservative"),
    ("Norfolk",                 "County",       84,  "TBC"),
    ("Suffolk",                 "County",       70,  "Reform UK"),
    ("West Sussex",             "County",       70,  "Lib Dem / Green"),
    # District Councils
    ("Adur",                    "District",     29,  "Labour"),
    ("Basildon",                "District",     42,  "Con / Reform / Ind"),
    ("Basingstoke and Deane",   "District",     54,  "Ind / Lib Dem"),
    ("Brentwood",               "District",     39,  "Lib Dem / Labour"),
    ("Broxbourne",              "District",     30,  "Conservative"),
    ("Burnley",                 "District",     45,  "Ind / Lib Dem / Green"),
    ("Cambridge",               "District",     42,  "TBC"),
    ("Cannock Chase",           "District",     36,  "Reform UK"),
    ("Cheltenham",              "District",     40,  "Lib Dem"),
    ("Cherwell",                "District",     48,  "Lib Dem / Green / Ind"),
    ("Chorley",                 "District",     42,  "Labour"),
    ("Colchester",              "District",     51,  "TBC"),
    ("Crawley",                 "District",     36,  "Labour"),
    ("Eastleigh",               "District",     39,  "Lib Dem"),
    ("Epping Forest",           "District",     54,  "TBC"),
    ("Exeter",                  "District",     39,  "Labour"),
    ("Fareham",                 "District",     32,  "Conservative"),
    ("Gosport",                 "District",     28,  "Conservative"),
    ("Harlow",                  "District",     33,  "Conservative"),
    ("Hart",                    "District",     33,  "Lib Dem / Ind"),
    ("Hastings",                "District",     32,  "Green"),
    ("Havant",                  "District",     36,  "Lab / Green / Lib Dem"),
    ("Huntingdonshire",         "District",     52,  "Lib Dem / Ind / Green"),
    ("Hyndburn",                "District",     35,  "TBC"),
    ("Ipswich",                 "District",     48,  "Labour"),
    ("Lincoln",                 "District",     33,  "Labour"),
    ("Newcastle-under-Lyme",    "District",     44,  "Reform UK"),
    ("Norwich",                 "District",     39,  "Green"),
    ("Nuneaton and Bedworth",   "District",     38,  "Reform UK"),
    ("Oxford",                  "District",     48,  "Labour"),
    ("Pendle",                  "District",     33,  "Ind / Lib Dem"),
    ("Preston",                 "District",     48,  "Labour"),
    ("Redditch",                "District",     27,  "Con / Ind"),
    ("Rochford",                "District",     39,  "Reform UK"),
    ("Rugby",                   "District",     42,  "Lab / Lib Dem"),
    ("Rushmoor",                "District",     39,  "Labour"),
    ("South Cambridgeshire",    "District",     45,  "Lib Dem"),
    ("St Albans",               "District",     56,  "Lib Dem"),
    ("Stevenage",               "District",     39,  "Labour"),
    ("Tamworth",                "District",     30,  "Labour"),
    ("Three Rivers",            "District",     39,  "Lib Dem"),
    ("Tunbridge Wells",         "District",     39,  "Lib Dem"),
    ("Watford",                 "District",     36,  "Lib Dem (Mayor)"),
    ("Welwyn Hatfield",         "District",     48,  "Lab / Lib Dem"),
    ("West Lancashire",         "District",     45,  "Conservative"),
    ("West Oxfordshire",        "District",     49,  "Lib Dem / Lab / Green"),
    ("Winchester",              "District",     45,  "Lib Dem"),
    ("Worthing",                "District",     37,  "Labour"),
    # London Boroughs
    ("Barking and Dagenham",    "London",       51,  "Labour"),
    ("Barnet",                  "London",       63,  "Labour"),
    ("Bexley",                  "London",       45,  "Conservative"),
    ("Brent",                   "London",       57,  "Labour"),
    ("Bromley",                 "London",       58,  "Conservative"),
    ("Camden",                  "London",       55,  "Labour"),
    ("Croydon",                 "London",       70,  "Conservative (Mayor)"),
    ("Ealing",                  "London",       70,  "Labour"),
    ("Enfield",                 "London",       63,  "TBC"),
    ("Greenwich",               "London",       55,  "Labour"),
    ("Hackney",                 "London",       57,  "Green (Mayor)"),
    ("Hammersmith and Fulham",  "London",       50,  "Labour"),
    ("Haringey",                "London",       57,  "Green"),
    ("Harrow",                  "London",       55,  "Conservative"),
    ("Havering",                "London",       55,  "Reform UK"),
    ("Hillingdon",              "London",       53,  "Conservative"),
    ("Hounslow",                "London",       62,  "Labour"),
    ("Islington",               "London",       51,  "Labour"),
    ("Kensington and Chelsea",  "London",       50,  "Conservative"),
    ("Kingston upon Thames",    "London",       48,  "Lib Dem"),
    ("Lambeth",                 "London",       63,  "TBC"),
    ("Lewisham",                "London",       54,  "Green (Mayor)"),
    ("Merton",                  "London",       57,  "Labour"),
    ("Newham",                  "London",       66,  "Labour (Mayor)"),
    ("Redbridge",               "London",       63,  "Labour"),
    ("Richmond upon Thames",    "London",       54,  "Lib Dem"),
    ("Southwark",               "London",       63,  "TBC"),
    ("Sutton",                  "London",       55,  "Lib Dem"),
    ("Tower Hamlets",           "London",       45,  "Aspire (Mayor)"),
    ("Waltham Forest",          "London",       60,  "Green"),
    ("Wandsworth",              "London",       58,  "TBC"),
    ("Westminster",             "London",       54,  "Conservative"),
    # Metropolitan Boroughs
    ("Barnsley",                "Metropolitan", 63,  "Reform UK"),
    ("Birmingham",              "Metropolitan", 101, "TBC"),
    ("Bolton",                  "Metropolitan", 60,  "Lab / Con / LD / Ind"),
    ("Bradford",                "Metropolitan", 90,  "Reform UK"),
    ("Bury",                    "Metropolitan", 51,  "Labour"),
    ("Calderdale",              "Metropolitan", 54,  "Reform UK"),
    ("Coventry",                "Metropolitan", 54,  "Labour"),
    ("Dudley",                  "Metropolitan", 72,  "Conservative"),
    ("Gateshead",               "Metropolitan", 66,  "Reform UK"),
    ("Kirklees",                "Metropolitan", 69,  "TBC"),
    ("Knowsley",                "Metropolitan", 45,  "Labour"),
    ("Leeds",                   "Metropolitan", 99,  "Labour"),
    ("Manchester",              "Metropolitan", 96,  "Labour"),
    ("Newcastle upon Tyne",     "Metropolitan", 78,  "TBC"),
    ("North Tyneside",          "Metropolitan", 60,  "Labour (Mayor)"),
    ("Oldham",                  "Metropolitan", 60,  "TBC"),
    ("Rochdale",                "Metropolitan", 60,  "Labour"),
    ("Salford",                 "Metropolitan", 60,  "Labour (Mayor)"),
    ("Sandwell",                "Metropolitan", 72,  "Reform UK"),
    ("Sefton",                  "Metropolitan", 66,  "Labour"),
    ("Sheffield",               "Metropolitan", 84,  "Lab / Green"),
    ("Solihull",                "Metropolitan", 51,  "Conservative"),
    ("South Tyneside",          "Metropolitan", 54,  "Reform UK"),
    ("St. Helens",              "Metropolitan", 48,  "Reform UK"),
    ("Stockport",               "Metropolitan", 63,  "Lib Dem"),
    ("Sunderland",              "Metropolitan", 75,  "Reform UK"),
    ("Tameside",                "Metropolitan", 57,  "Labour"),
    ("Trafford",                "Metropolitan", 63,  "Labour"),
    ("Wakefield",               "Metropolitan", 63,  "Reform UK"),
    ("Walsall",                 "Metropolitan", 60,  "Reform UK"),
    ("Wigan",                   "Metropolitan", 75,  "Labour"),
    ("Wolverhampton",           "Metropolitan", 60,  "Labour"),
    # Unitary Authorities
    ("Blackburn with Darwen",   "Unitary",      51,  "Labour"),
    ("Halton",                  "Unitary",      54,  "Labour"),
    ("Hartlepool",              "Unitary",      36,  "Reform UK / Con"),
    ("Isle of Wight",           "Unitary",      39,  "Reform UK"),
    ("Kingston upon Hull",      "Unitary",      57,  "Lib Dem"),
    ("Milton Keynes",           "Unitary",      60,  "Lib Dem / Labour"),
    ("North East Lincolnshire", "Unitary",      42,  "Reform UK"),
    ("Peterborough",            "Unitary",      60,  "Lab / Lib Dem"),
    ("Plymouth",                "Unitary",      57,  "Labour"),
    ("Portsmouth",              "Unitary",      42,  "Lib Dem"),
    ("Reading",                 "Unitary",      48,  "Labour"),
    ("Southampton",             "Unitary",      51,  "Labour"),
    ("Southend-on-Sea",         "Unitary",      51,  "Lab / Green / LD"),
    ("Swindon",                 "Unitary",      57,  "Conservative"),
    ("Thurrock",                "Unitary",      49,  "Reform UK"),
    ("Wokingham",               "Unitary",      54,  "Lib Dem"),
]

# Known results page URLs (manually verified)
KNOWN_URLS = {
    "Sandwell":  "https://www.sandwell.gov.uk/election-results/candidates-standing-2026-local-elections",
    "Sutton":    "https://www.sutton.gov.uk/w/local-elections-results-7-may-2026",
}

def slug(name):
    """Convert council name to safe folder name."""
    return re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_').replace('.', '')

# ── 1. CREATE FOLDER STRUCTURE ────────────────────────────────────────────────

print("Creating folder structure...")
for name, ctype, seats, control in COUNCILS:
    folder = os.path.join(ROOT, slug(name))
    os.makedirs(folder, exist_ok=True)

print(f"  Created {len(COUNCILS)} council folders")

# ── 2. MASTER SOURCE EXCEL ────────────────────────────────────────────────────

print("Building master source Excel...")

rows = []
for name, ctype, seats, control in COUNCILS:
    url = KNOWN_URLS.get(name, "")
    # Build a best-guess search URL for councils without known URLs
    search_hint = f"https://www.google.com/search?q={name.replace(' ', '+')}+council+election+results+2026"
    rows.append({
        "Council":           name,
        "Type":              ctype,
        "Seats (2026)":      seats,
        "Control After Election": control,
        "Results Page URL":  url if url else "— see search hint →",
        "Search Hint":       "" if url else search_hint,
        "Data Collected":    "Yes" if name == "Sandwell" else "No",
        "Source":            "opencouncildata.co.uk/elections.php",
        "Notes":             "",
    })

df_master = pd.DataFrame(rows)

master_path = os.path.join(ROOT, "00_Master_Sources.xlsx")
with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
    df_master.to_excel(writer, sheet_name="Council Sources", index=False)

# Style it
wb = load_workbook(master_path)
ws = wb.active

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_FILL    = PatternFill("solid", fgColor="EBF1F8")
BORDER      = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
TYPE_COLOURS = {
    "County":       "D6E4F0", "District":     "D5F5E3",
    "London":       "FDEBD0", "Metropolitan": "E8DAEF",
    "Unitary":      "FDFFD6",
}
CONTROL_COLOURS = {
    "Reform UK": "12B6CF", "Labour": "E4003B", "Green": "02A95B",
    "Conservative": "0087DC", "Lib Dem": "FAA61A",
}

for cell in ws[1]:
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = BORDER

for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
    ctype_val = row[1].value
    fill_col  = TYPE_COLOURS.get(ctype_val, "FFFFFF")
    for cell in row:
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="left", wrap_text=False)
    row[1].fill = PatternFill("solid", fgColor=fill_col)

    ctrl = row[3].value or ""
    for party, col in CONTROL_COLOURS.items():
        if party in ctrl:
            row[3].fill = PatternFill("solid", fgColor=col)
            row[3].font = Font(color="FFFFFF", bold=True)
            break

    # Hyperlink for known URLs
    url_cell = row[4]
    if url_cell.value and url_cell.value.startswith("http"):
        url_cell.hyperlink = url_cell.value
        url_cell.font = Font(color="0563C1", underline="single")

    # Yes/No colouring
    collected = row[6]
    if collected.value == "Yes":
        collected.fill = PatternFill("solid", fgColor="00B050")
        collected.font = Font(color="FFFFFF", bold=True)

ws.freeze_panes = "A2"
col_widths = {"A":28,"B":14,"C":13,"D":22,"E":38,"F":50,"G":16,"H":30,"I":20}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

# Type legend
ws_leg = wb.create_sheet("Legend")
ws_leg["A1"] = "Council Type"
ws_leg["B1"] = "Colour"
for i, (t, c) in enumerate(TYPE_COLOURS.items(), start=2):
    ws_leg[f"A{i}"] = t
    ws_leg[f"B{i}"].fill = PatternFill("solid", fgColor=c)

wb.save(master_path)
print(f"  Saved: {master_path}")

# ── 3. REFORM UK FOLDER + EXCEL ───────────────────────────────────────────────

print("Building Reform UK councils Excel...")
reform_folder = os.path.join(ROOT, "Reform_UK_Analysis")
os.makedirs(reform_folder, exist_ok=True)

reform_councils = [
    (name, ctype, seats, control)
    for name, ctype, seats, control in COUNCILS
    if "Reform" in control
]

df_reform = pd.DataFrame([{
    "Council":      name,
    "Type":         ctype,
    "Seats":        seats,
    "Control":      control,
    "Results URL":  KNOWN_URLS.get(name, ""),
    "Notes":        "Full ward-level data available" if name == "Sandwell" else "To be scraped",
} for name, ctype, seats, control in reform_councils])

reform_path = os.path.join(reform_folder, "Reform_UK_Councils_2026.xlsx")
with pd.ExcelWriter(reform_path, engine="openpyxl") as writer:
    df_reform.to_excel(writer, sheet_name="Reform UK Results", index=False)

wb2 = load_workbook(reform_path)
ws2 = wb2.active
for cell in ws2[1]:
    cell.fill = PatternFill("solid", fgColor="12B6CF")
    cell.font = Font(color="FFFFFF", bold=True)
    cell.border = BORDER
for row in ws2.iter_rows(min_row=2):
    for cell in row:
        cell.border = BORDER
ws2.freeze_panes = "A2"
for col, w in zip("ABCDEF", [28,14,8,22,45,30]):
    ws2.column_dimensions[col].width = w
wb2.save(reform_path)
print(f"  Saved: {reform_path}")
print(f"  Reform UK won/led in {len(reform_councils)} councils")

print("\nDone. Summary:")
print(f"  Council folders created: {len(COUNCILS)}")
print(f"  Master Excel:  00_Master_Sources.xlsx")
print(f"  Reform Excel:  Reform_UK_Analysis/Reform_UK_Councils_2026.xlsx")
print(f"\nNOTE: Delete the old 'sandwell_project' folder manually in Windows Explorer.")
