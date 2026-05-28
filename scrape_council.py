"""
scrape_council.py — General UK council election results scraper.

Usage (auto-scrape mode):
  python scrape_council.py "Sutton" "https://www.sutton.gov.uk/w/local-elections-results-7-may-2026"

Usage (JSON input mode — when Claude pre-extracted the data via WebFetch):
  python scrape_council.py "Chorley" "https://..." --json data.json

The script:
  1. Fetches the main results page and discovers ward sub-pages
  2. Scrapes each ward page (handles "table with Candidate/Party/Votes/Elected" format)
  3. Normalises party names into standard groups
  4. Builds a 5-sheet styled Excel at Councils/<Name>/<Name>_Elections_2026.xlsx
  5. Updates 00_Master_Sources.xlsx (marks Data Collected = Yes)
  6. Prints a summary

If auto-scraping fails (unusual HTML structure), Claude should collect the data
via WebFetch and pass it as JSON. See --help for the expected JSON schema.
"""

import sys, os, re, json, time, warnings, argparse
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

warnings.filterwarnings("ignore")  # suppress SSL warnings common on Windows

BASE_DIR = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project"
MASTER   = os.path.join(BASE_DIR, "00_Master_Sources.xlsx")

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# ── PARTY NORMALISATION ──────────────────────────────────────────────────────

PARTY_MAP = {
    "labour party":                   "Labour",
    "labour":                         "Labour",
    "green party":                    "Green",
    "the green party":                "Green",
    "green":                          "Green",
    "reform uk":                      "Reform UK",
    "reform":                         "Reform UK",
    "conservative party":             "Conservative",
    "conservative and unionist party":"Conservative",
    "conservative":                   "Conservative",
    "local conservatives":            "Conservative",
    "the conservative party":         "Conservative",
    "liberal democrats":              "Lib Dem",
    "liberal democrat":               "Lib Dem",
    "lib dem":                        "Lib Dem",
    "lib dems":                       "Lib Dem",
    "independent":                    "Independent",
    "workers party of britain":       "Other",
    "workers party":                  "Other",
    "tusc":                           "Other",
    "trade unionist and socialist coalition": "Other",
    "ukip":                           "Other",
    "brexit party":                   "Other",
    "social democratic party":        "Other",
    "sdp":                            "Other",
    "heritage party":                 "Other",
    "english democrats":              "Other",
    "advance uk":                     "Other",
    "britain first":                  "Other",
    "reclaim":                        "Other",
    "for britain":                    "Other",
    "community":                      "Other",
}

def normalise_party(raw: str) -> str:
    key = raw.strip().lower()
    return PARTY_MAP.get(key, "Other")

def bloc(party: str) -> str:
    if party in ("Labour", "Green"):         return "Left"
    if party in ("Reform UK", "Conservative"): return "Right"
    if party == "Lib Dem":                   return "Lib Dem"
    return "Other/Independent"

PARTY_COLOURS = {
    "Reform UK":    "12B6CF",
    "Labour":       "E4003B",
    "Green":        "02A95B",
    "Conservative": "0087DC",
    "Lib Dem":      "FAA61A",
    "Independent":  "888888",
    "Other":        "AAAAAA",
}

# ── WEB SCRAPING ──────────────────────────────────────────────────────────────

def fetch(url: str, retries=3, delay=1.5) -> tuple[BeautifulSoup, str]:
    """Returns (soup, final_url_after_redirects)."""
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20, verify=False)
            r.raise_for_status()
            return BeautifulSoup(r.text, "lxml"), r.url
        except Exception as e:
            if attempt == retries - 1:
                raise
            time.sleep(delay)

def find_ward_links(soup: BeautifulSoup, final_url: str) -> list[tuple[str, str]]:
    """Return list of (ward_name, full_url) pairs."""
    parsed = urlparse(final_url)
    main_path = parsed.path.rstrip("/")
    origin = f"{parsed.scheme}://{parsed.netloc}"

    ward_links = []
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        # Only sub-paths of the main results page (after redirect)
        if href.startswith(main_path + "/"):
            full = urljoin(origin, href).split("#")[0]
            name = a.get_text(strip=True)
            if full not in seen and name and len(name) < 80:
                seen.add(full)
                ward_links.append((name, full))
    return ward_links

def parse_ward_page(ward_name: str, url: str) -> dict | None:
    """
    Try to extract candidate data from a ward page.
    Handles the 4-column table format: Candidate | Party | Votes | Elected/Not Elected
    Returns a dict or None if parsing failed.
    """
    try:
        soup, _ = fetch(url)
    except Exception as e:
        print(f"  WARNING: Could not fetch {url}: {e}")
        return None

    candidates = []

    # Strategy 1: look for a <table> whose header contains "Party" and "Votes"
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        header_cells = [c.get_text(strip=True).lower() for c in rows[0].find_all(["th", "td"])]
        if "party" in header_cells and any("vote" in h for h in header_cells):
            # Find column indices
            try:
                party_i  = next(i for i, h in enumerate(header_cells) if "party" in h)
                votes_i  = next(i for i, h in enumerate(header_cells) if "vote" in h)
                name_i   = 0
                # Elected column: specifically "elected/not elected" or the LAST column with "elect"
                # Avoid matching "election candidate" (first column)
                elect_candidates = [i for i, h in enumerate(header_cells)
                                    if "elect" in h and i != name_i and "candidate" not in h]
                elect_i = elect_candidates[-1] if elect_candidates else None
            except StopIteration:
                continue

            for row in rows[1:]:
                cells = row.find_all(["th", "td"])
                if len(cells) < max(party_i, votes_i) + 1:
                    continue
                raw_name    = cells[name_i].get_text(strip=True)
                raw_party   = cells[party_i].get_text(strip=True)
                raw_votes   = cells[votes_i].get_text(strip=True).replace(",", "").strip()
                raw_elected = cells[elect_i].get_text(strip=True) if elect_i is not None else ""

                if not raw_name or not raw_votes:
                    continue
                try:
                    votes = int(re.sub(r"[^\d]", "", raw_votes))
                except ValueError:
                    continue

                elected = "elected" in raw_elected.lower() and "not" not in raw_elected.lower()

                candidates.append({
                    "name":    raw_name,
                    "party":   normalise_party(raw_party),
                    "votes":   votes,
                    "elected": elected,
                })
            if candidates:
                break

    # Extract electorate/turnout if present (some councils include this)
    text = soup.get_text(separator=" ")
    electorate = _extract_number(text, r"electorate[:\s]+([0-9,]+)")
    ballots    = _extract_number(text, r"(?:total votes|ballots cast|votes cast|total ballot)[:\s]+([0-9,]+)")
    turnout    = _extract_float(text,  r"turnout[:\s]+([\d.]+)\s*%")

    if not candidates:
        return None

    return {
        "ward":       ward_name,
        "electorate": electorate,
        "ballots":    ballots,
        "turnout":    turnout,
        "candidates": candidates,
    }

def _extract_number(text, pattern):
    m = re.search(pattern, text, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    return None

def _extract_float(text, pattern):
    m = re.search(pattern, text, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None

# ── EXCEL BUILDING ────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_FILL    = PatternFill("solid", fgColor="EBF1F8")
BORDER_SIDE = Side(style="thin", color="CCCCCC")
THIN_BORDER = Border(left=BORDER_SIDE, right=BORDER_SIDE, top=BORDER_SIDE, bottom=BORDER_SIDE)

def style_sheet(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = ALT_FILL if i % 2 == 0 else PatternFill()
        for cell in row:
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 42)
    ws.freeze_panes = "A2"

def build_excel(council_name: str, wards: list[dict], output_path: str):
    rows = []
    for w in wards:
        for c in w["candidates"]:
            rows.append({
                "Ward":      w["ward"],
                "Candidate": c["name"],
                "Party":     c["party"],
                "Votes":     c["votes"],
                "Elected":   "Yes" if c["elected"] else "No",
                "Bloc":      bloc(c["party"]),
            })

    df = pd.DataFrame(rows)
    if df.empty:
        raise ValueError("No candidate data to write")

    grand_total = df["Votes"].sum()

    # Sheet 1 – raw results
    raw = df[["Ward","Candidate","Party","Votes","Elected"]].sort_values(["Ward","Votes"], ascending=[True,False])

    # Sheet 2 – party totals
    pt = (df.groupby("Party")["Votes"].sum().rename("Total Votes")
            .reset_index().sort_values("Total Votes", ascending=False))
    pt["% of All Votes"] = (pt["Total Votes"] / grand_total * 100).round(2)
    seats = df[df["Elected"]=="Yes"].groupby("Party").size().rename("Seats Won").reset_index()
    pt = pt.merge(seats, on="Party", how="left")
    pt["Seats Won"] = pt["Seats Won"].fillna(0).astype(int)

    # Sheet 3 – blocs
    bt = (df.groupby("Bloc")["Votes"].sum().rename("Total Votes")
            .reset_index().sort_values("Total Votes", ascending=False))
    bt["% of All Votes"] = (bt["Total Votes"] / grand_total * 100).round(2)

    # Sheet 4 – ward analysis
    ward_records = []
    for w in wards:
        wdf = df[df["Ward"] == w["ward"]]
        def pvotes(p):  return int(wdf[wdf["Party"]==p]["Votes"].sum())
        def pbloc(b):   return int(wdf[wdf["Bloc"]==b]["Votes"].sum())
        def seats_by(p): return int((wdf[(wdf["Party"]==p) & (wdf["Elected"]=="Yes")]).shape[0])

        reform  = pvotes("Reform UK")
        labour  = pvotes("Labour")
        green   = pvotes("Green")
        con     = pvotes("Conservative")
        libdem  = pvotes("Lib Dem")
        ind     = pvotes("Independent")
        other   = pvotes("Other")
        left    = labour + green
        right   = reform + con
        total   = int(wdf["Votes"].sum())

        split_evident = (left > reform) and (seats_by("Reform UK") > 0)
        ward_records.append({
            "Ward":                  w["ward"],
            "Electorate":            w.get("electorate") or "",
            "Ballots Cast":          w.get("ballots") or "",
            "Turnout %":             w.get("turnout") or "",
            "Reform UK Votes":       reform,
            "Labour Votes":          labour,
            "Green Votes":           green,
            "Left (Lab+Green) Votes": left,
            "Conservative Votes":    con,
            "Right (Ref+Con) Votes": right,
            "Lib Dem Votes":         libdem,
            "Independent Votes":     ind,
            "Other Votes":           other,
            "Total Candidate Votes": total,
            "Reform % of Votes":     round(reform/total*100,1) if total else 0,
            "Labour % of Votes":     round(labour/total*100,1) if total else 0,
            "Green % of Votes":      round(green/total*100,1)  if total else 0,
            "Left % of Votes":       round(left/total*100,1)   if total else 0,
            "Reform Seats":          seats_by("Reform UK"),
            "Labour Seats":          seats_by("Labour"),
            "Green Seats":           seats_by("Green"),
            "Conservative Seats":    seats_by("Conservative"),
            "Independent Seats":     seats_by("Independent"),
            "Split Vote Evident?":   "YES" if split_evident else "no",
            "Left Votes > Reform?":  "YES" if left > reform else "no",
        })
    ward_df = pd.DataFrame(ward_records)

    # Sheet 5 – Flourish data
    flourish = ward_df[[
        "Ward","Reform UK Votes","Labour Votes","Green Votes",
        "Conservative Votes","Lib Dem Votes","Independent Votes",
        "Left (Lab+Green) Votes","Right (Ref+Con) Votes",
        "Reform % of Votes","Labour % of Votes","Green % of Votes","Left % of Votes",
        "Reform Seats","Labour Seats","Green Seats","Conservative Seats",
        "Split Vote Evident?",
    ]].copy()

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        raw.to_excel(writer,     sheet_name="1. Raw Results",    index=False)
        pt.to_excel(writer,      sheet_name="2. Party Totals",   index=False)
        bt.to_excel(writer,      sheet_name="3. Left-Right Blocs", index=False)
        ward_df.to_excel(writer, sheet_name="4. Ward Analysis",  index=False)
        flourish.to_excel(writer,sheet_name="5. Flourish Data",  index=False)

    # Style
    wb = load_workbook(output_path)
    for sh in wb.sheetnames:
        style_sheet(wb[sh])

    ws1 = wb["1. Raw Results"]
    for row in ws1.iter_rows(min_row=2):
        party_val = row[2].value
        colour = PARTY_COLOURS.get(party_val)
        if colour:
            row[2].fill = PatternFill("solid", fgColor=colour)
            row[2].font = Font(color="FFFFFF", bold=True)
        if row[4].value == "Yes":
            row[4].fill = PatternFill("solid", fgColor="00B050")
            row[4].font = Font(color="FFFFFF", bold=True)

    ws4 = wb["4. Ward Analysis"]
    split_col = next((c.column for c in ws4[1] if c.value == "Split Vote Evident?"), None)
    if split_col:
        for row in ws4.iter_rows(min_row=2):
            cell = row[split_col - 1]
            if cell.value == "YES":
                cell.fill = PatternFill("solid", fgColor="FF0000")
                cell.font = Font(color="FFFFFF", bold=True)

    wb.save(output_path)

def update_master(council_name: str, url: str):
    if not os.path.exists(MASTER):
        print("  SKIP: Master sources file not found")
        return
    import openpyxl as xl
    wb = xl.load_workbook(MASTER)
    ws = wb.active
    headers = {c.value: c.column for c in ws[1] if c.value}
    url_col       = headers.get("Results Page URL")
    collected_col = headers.get("Data Collected")
    council_col   = headers.get("Council")
    if not all([url_col, collected_col, council_col]):
        print("  SKIP: Master sources columns not found")
        return
    for row in ws.iter_rows(min_row=2):
        cell_val = row[council_col - 1].value
        if cell_val and council_name.lower() in str(cell_val).lower():
            row[url_col - 1].value = url
            row[collected_col - 1].value = "Yes"
            print(f"  Master sources updated for: {cell_val}")
            break
    wb.save(MASTER)

# ── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Scrape UK council election results")
    parser.add_argument("council", help="Council name (e.g. Sutton)")
    parser.add_argument("url",     help="Main results page URL")
    parser.add_argument("--json",  metavar="FILE", help="Load pre-scraped data from JSON instead of fetching")
    parser.add_argument("--workers", type=int, default=6, help="Parallel ward fetches (default 6)")
    args = parser.parse_args()

    council = args.council
    url     = args.url
    safe    = re.sub(r"[^\w\s-]", "", council).strip().replace(" ", "_")

    # Determine output folder (Councils/ or County Councils/)
    county_keywords = ["essex","suffolk","hampshire","hertfordshire","gloucestershire",
                       "east sussex","west sussex","norfolk","kent","surrey","warwickshire"]
    folder = "County Councils" if any(k in council.lower() for k in county_keywords) else "Councils"
    output_path = os.path.join(BASE_DIR, folder, safe, f"{council}_Elections_2026.xlsx")

    # ── Collect ward data ────────────────────────────────────────────────────
    if args.json:
        print(f"Loading pre-scraped data from {args.json}...")
        with open(args.json, encoding="utf-8") as f:
            wards = json.load(f)
        # Normalise parties in case Claude used raw party names
        for w in wards:
            for c in w["candidates"]:
                c["party"] = normalise_party(c["party"])
    else:
        print(f"Fetching main page: {url}")
        main_soup, final_url = fetch(url)
        if final_url != url:
            print(f"  Redirected to: {final_url}")

        ward_links = find_ward_links(main_soup, final_url)
        if not ward_links:
            print("ERROR: Could not find ward sub-pages. Try --json mode (see --help).")
            sys.exit(1)

        print(f"Found {len(ward_links)} ward pages. Scraping in parallel ({args.workers} workers)...")

        wards = []
        failed = []
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = {pool.submit(parse_ward_page, name, wurl): name
                       for name, wurl in ward_links}
            for fut in as_completed(futures):
                ward_name = futures[fut]
                result = fut.result()
                if result:
                    wards.append(result)
                    print(f"  ✓ {ward_name} ({len(result['candidates'])} candidates)")
                else:
                    failed.append(ward_name)
                    print(f"  ✗ {ward_name} — parse failed")

        if failed:
            print(f"\nWARNING: {len(failed)} wards could not be parsed: {failed}")
        if not wards:
            print("ERROR: No ward data collected. Use --json mode.")
            sys.exit(1)

        wards.sort(key=lambda w: w["ward"])

    # ── Build Excel ──────────────────────────────────────────────────────────
    print(f"\nBuilding Excel: {output_path}")
    build_excel(council, wards, output_path)
    print(f"Saved.")

    # ── Update master sources ────────────────────────────────────────────────
    update_master(council, url)

    # ── Summary ──────────────────────────────────────────────────────────────
    all_candidates = [c for w in wards for c in w["candidates"]]
    total_votes    = sum(c["votes"] for c in all_candidates)
    seats          = [c for c in all_candidates if c["elected"]]

    from collections import Counter
    seats_by_party = Counter(c["party"] for c in seats)
    votes_by_party = {}
    for c in all_candidates:
        votes_by_party[c["party"]] = votes_by_party.get(c["party"], 0) + c["votes"]

    split_wards = [w for w in wards
                   if sum(c["votes"] for c in w["candidates"] if c["party"] in ("Labour","Green"))
                      > sum(c["votes"] for c in w["candidates"] if c["party"] == "Reform UK")
                   and any(c["elected"] and c["party"] == "Reform UK" for c in w["candidates"])]

    print(f"\n{'='*50}")
    print(f"COUNCIL: {council}")
    print(f"Wards scraped: {len(wards)}")
    print(f"Total candidates: {len(all_candidates)}")
    print(f"Total votes: {total_votes:,}")
    print(f"\nSeats won:")
    for party, n in sorted(seats_by_party.items(), key=lambda x: -x[1]):
        pct = votes_by_party.get(party,0) / total_votes * 100
        print(f"  {party:20s}: {n:3d} seats  ({pct:.1f}% of votes)")
    print(f"\nSplit-vote evidence (Left > Reform but Reform won seats): {len(split_wards)} wards")
    for w in split_wards:
        print(f"  - {w['ward']}")
    print(f"\nOutput: {output_path}")


if __name__ == "__main__":
    main()
