"""
Creates the Restore/ folder with full ward-level results for every ward
where Restore Britain (via 'Great Yarmouth First') stood in May 2026.
"""
import os, glob
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ROOT       = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"
OUT_FOLDER = os.path.join(ROOT, "Restore")
os.makedirs(OUT_FOLDER, exist_ok=True)

BORDER = Border(*[Side(style="thin", color="CCCCCC")] * 4)

PARTY_COLOURS = {
    "Reform UK":    "12B6CF",
    "Labour":       "E4003B",
    "Green":        "02A95B",
    "Conservative": "0087DC",
    "Lib Dem":      "FAA61A",
    "Independent":  "888888",
    "Other":        "AAAAAA",
    "Restore Britain / Great Yarmouth First": "8E44AD",
}

RESTORE_WARDS = {
    "Great Yarmouth": ["Caister South"],
    "Norfolk": [
        "Breydon", "Gorleston", "Lothingland", "Magdalen",
        "North Caister & Ormesby", "South Caister & Bure",
        "The Fleggs", "Yarmouth Nelson & Southtown", "Yarmouth North & Central",
    ],
}

def find_excel(council_name):
    for p in glob.glob(os.path.join(ROOT, "*", "*_Elections_2026.xlsx")):
        folder = os.path.basename(os.path.dirname(p))
        if council_name.lower() in folder.lower():
            return p
    return None

def style_header(ws):
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="4A235A")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = "A2"

def style_rows(ws):
    alt = PatternFill("solid", fgColor="F5EEF8")
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        f = alt if i % 2 == 0 else PatternFill()
        for cell in row:
            cell.fill = f
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")

def auto_width(ws, max_w=45):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(w + 4, max_w)

# ── Collect all ward data ─────────────────────────────────────────────────────

all_rows = []   # full raw candidate rows
summary_rows = []  # one row per ward

for council_name, wards in RESTORE_WARDS.items():
    excel_path = find_excel(council_name)
    if not excel_path:
        print(f"  WARNING: no Excel found for {council_name}")
        continue

    df = pd.read_excel(excel_path, sheet_name="1. Raw Results")

    for ward in wards:
        wdf = df[df["Ward"] == ward].copy()
        wdf["Council"] = council_name
        wdf["Council Type"] = "County Council" if council_name == "Norfolk" else "Borough Council"
        all_rows.append(wdf)

        # Per-ward summary
        tot = wdf["Votes"].sum()
        restore = wdf[wdf["Party"].str.contains("Great Yarmouth First", na=False)]["Votes"].sum()
        reform  = wdf[wdf["Party"].str.contains("Reform", na=False)]["Votes"].sum()
        labour  = wdf[wdf["Party"].str.contains("Labour", na=False)]["Votes"].sum()
        con     = wdf[wdf["Party"].str.contains("Conserv", na=False)]["Votes"].sum()
        libdem  = wdf[wdf["Party"].str.contains("Liberal|Lib Dem", na=False)]["Votes"].sum()
        green   = wdf[wdf["Party"].str.contains("Green", na=False)]["Votes"].sum()
        ind     = wdf[wdf["Party"].str.contains("Independent", na=False)]["Votes"].sum()
        other   = tot - restore - reform - labour - con - libdem - green - ind

        restore_elected = wdf[(wdf["Party"].str.contains("Great Yarmouth First", na=False)) &
                               (wdf["Elected"] == "Yes")].shape[0]

        summary_rows.append({
            "Council":              council_name,
            "Council Type":         "County Council" if council_name == "Norfolk" else "Borough Council",
            "Ward":                 ward,
            "Restore (GYF) Votes":  restore,
            "Reform UK Votes":      reform,
            "Labour Votes":         labour,
            "Conservative Votes":   con,
            "Lib Dem Votes":        libdem,
            "Green Votes":          green,
            "Independent Votes":    ind,
            "Other Votes":          other if other > 0 else 0,
            "Total Votes":          tot,
            "Restore % of Votes":   round(restore / tot * 100, 1) if tot else 0,
            "Reform % of Votes":    round(reform  / tot * 100, 1) if tot else 0,
            "Labour % of Votes":    round(labour  / tot * 100, 1) if tot else 0,
            "Restore Seats Won":    restore_elected,
        })

raw_df     = pd.concat(all_rows, ignore_index=True)
summary_df = pd.DataFrame(summary_rows)

# Reorder raw columns
col_order = ["Council", "Council Type", "Ward", "Candidate", "Party", "Votes", "Elected"]
raw_df = raw_df[[c for c in col_order if c in raw_df.columns]]
raw_df = raw_df.sort_values(["Council", "Ward", "Votes"], ascending=[True, True, False]).reset_index(drop=True)

# ── Party overview ────────────────────────────────────────────────────────────

party_map = {
    "Great Yarmouth First": "Restore Britain / Great Yarmouth First",
}

raw_df["Party Group"] = raw_df["Party"].apply(
    lambda p: party_map.get(p, p) if "Great Yarmouth First" in str(p) else p
)

party_overview = (
    raw_df.groupby("Party").agg(
        Candidates=("Candidate", "count"),
        Total_Votes=("Votes", "sum"),
        Elected=("Elected", lambda x: (x == "Yes").sum()),
    )
    .rename(columns={"Total_Votes": "Total Votes", "Elected": "Seats Won"})
    .sort_values("Total Votes", ascending=False)
    .reset_index()
)
grand = raw_df["Votes"].sum()
party_overview["% of All Votes"] = (party_overview["Total Votes"] / grand * 100).round(2) if grand else 0

# ── Summary sheet ─────────────────────────────────────────────────────────────

meta = pd.DataFrame([{
    "Field":  "Party",              "Value": "Restore Britain (endorsed 'Great Yarmouth First')"},
    {"Field": "Registered",         "Value": "13 February 2026 (Electoral Commission PP18382)"},
    {"Field": "Strategy",           "Value": "Did not stand under own name — endorsed local affiliate Great Yarmouth First"},
    {"Field": "Wards contested",    "Value": len(summary_rows)},
    {"Field": "Councils contested", "Value": len(RESTORE_WARDS)},
    {"Field": "Total seats won",    "Value": int(summary_df["Restore Seats Won"].sum())},
    {"Field": "Win rate",           "Value": "100% (10/10)"},
    {"Field": "Geographic focus",   "Value": "Great Yarmouth area only (Norfolk CC + Great Yarmouth BC)"},
    {"Field": "Source",             "Value": "Electoral Commission + Democracy Club API + Hope Not Hate"},
])

# ── Write Excel ───────────────────────────────────────────────────────────────

OUT_PATH = os.path.join(OUT_FOLDER, "Restore_Britain_2026.xlsx")
with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
    meta.to_excel(writer,         sheet_name="1. Party Summary",  index=False)
    raw_df.drop(columns=["Party Group"], errors="ignore").to_excel(
                               writer, sheet_name="2. Raw Results",    index=False)
    party_overview.to_excel(writer, sheet_name="3. Party Totals",  index=False)
    summary_df.to_excel(writer,  sheet_name="4. Ward Analysis",    index=False)

# ── Style ─────────────────────────────────────────────────────────────────────

wb = load_workbook(OUT_PATH)

for sh in wb.sheetnames:
    ws = wb[sh]
    style_header(ws)
    style_rows(ws)
    auto_width(ws)

# Colour Party column in Raw Results
ws2 = wb["2. Raw Results"]
party_col = None
for cell in ws2[1]:
    if cell.value == "Party":
        party_col = cell.column; break

if party_col:
    for row in ws2.iter_rows(min_row=2):
        cell = row[party_col - 1]
        val  = str(cell.value or "")
        if "Great Yarmouth First" in val:
            cell.fill = PatternFill("solid", fgColor="8E44AD")
            cell.font = Font(color="FFFFFF", bold=True)
        # colour elected column green
        elected_cell = row[col_order.index("Elected")]
        if elected_cell.value == "Yes":
            elected_cell.fill = PatternFill("solid", fgColor="00B050")
            elected_cell.font = Font(color="FFFFFF", bold=True)

# Highlight Restore rows in Ward Analysis
ws4 = wb["4. Ward Analysis"]
for row in ws4.iter_rows(min_row=2):
    seats_cell = row[list(summary_df.columns).index("Restore Seats Won")]
    if seats_cell.value and int(seats_cell.value) > 0:
        seats_cell.fill = PatternFill("solid", fgColor="8E44AD")
        seats_cell.font = Font(color="FFFFFF", bold=True)

wb.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
print(f"Wards included: {len(summary_rows)}")
print(f"Total candidates across all wards: {len(raw_df)}")
print(f"Restore seats won: {int(summary_df['Restore Seats Won'].sum())}/10")
