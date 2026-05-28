# Scrape Local Election Results

Scrape UK local council election results from a public results page and produce a structured Excel file.

## Usage

```
/scrape-elections <council_name> <results_url>
```

**Example:**
```
/scrape-elections Sutton https://www.sutton.gov.uk/w/local-elections-results-7-may-2026
```

## What this skill does

1. **Tries the Python scraper first** — runs `scrape_council.py` which auto-fetches all ward pages and builds the Excel.
2. **Falls back to WebFetch** — if the HTML structure is non-standard (JavaScript-rendered, no table, etc.), Claude manually fetches each page and extracts data as JSON, then passes it to the script.
3. **Produces** an Excel file at `localelections_project/Councils/<Name>/<Name>_Elections_2026.xlsx` with these sheets:
   - `1. Raw Results` — every candidate row, styled by party colour
   - `2. Party Totals` — aggregate votes + % + seats per party
   - `3. Left-Right Blocs` — Left (Lab+Green) vs Right (Reform+Con) totals
   - `4. Ward Analysis` — per-ward breakdown with split-vote flag
   - `5. Flourish Data` — cleaned columns ready for Flourish/Datawrapper
4. **Updates** `localelections_project/00_Master_Sources.xlsx` — sets the Results URL and marks Data Collected = Yes.
5. **Reports** a summary: council name, wards scraped, total candidates, total votes, seats by party, and whether split-vote evidence was found.

## Step 1 — Try the Python scraper

Run this command first:

```powershell
python -X utf8 scrape_council.py "CouncilName" "https://results-url"
```

**The `-X utf8` flag is required on Windows** (avoids UnicodeEncodeError with emoji characters).

If it prints `Found N ward pages` and `Saved.` → you're done. Skip to Step 5 (report).

If it prints `ERROR: Could not find ward sub-pages` → proceed to Step 2.

## Step 2 — Fetch main results page (fallback only)

Use WebFetch on the URL. Ask it to return all hyperlinks to individual ward/division result pages.

If the page redirects, use the redirected URL for step 3.

## Step 3 — Fetch all ward pages in parallel (fallback only)

Use WebFetch on each ward URL simultaneously. For each page extract:
- Ward/division name
- For each candidate: name, party, votes (integer), elected (boolean)
- Electorate, total ballots cast, turnout % (if available — set to null if not)

## Step 4 — Build via JSON input (fallback only)

Write the extracted data to `data.json` in this format:

```json
[
  {
    "ward": "Ward Name",
    "electorate": 10000,
    "ballots": 3500,
    "turnout": 35.0,
    "candidates": [
      {"name": "Candidate Name", "party": "Reform UK", "votes": 1234, "elected": true},
      {"name": "Another Name", "party": "Labour Party", "votes": 987, "elected": false}
    ]
  }
]
```

Then run:

```powershell
python -X utf8 scrape_council.py "CouncilName" "https://results-url" --json data.json
```

Party names are normalised automatically. Use the raw names from the website.

## Step 5 — Report back

Print a summary: council name, wards scraped, total candidates, total votes, seats by party, and whether split-vote evidence was found. The script prints this automatically.

## Party normalisation (built into the script)

The script maps raw party names to standard groups. Common mappings:
- "Labour Party", "Labour" → Labour
- "Green Party", "The Green Party" → Green
- "Reform UK" → Reform UK
- "Conservative Party", "Conservative and Unionist Party", "Local Conservatives" → Conservative
- "Liberal Democrats", "Liberal Democrat" → Lib Dem
- "Independent" → Independent
- Everything else → Other

## Notes

- **SSL errors on Windows**: handled internally with `verify=False`. No action needed.
- **Numbered ward pages (Sandwell-style)**: some councils use `/2`, `/3` numeric paths. The auto-scraper won't detect these — use WebFetch fallback.
- **County councils**: if the council is a county council (Essex, Hampshire, Hertfordshire, etc.), the output goes to `County Councils/<Name>/` instead of `Councils/<Name>/`.
- If a page returns 403 or is behind authentication, note it in the master Excel and skip.
- Run time for auto-scrape: ~15–30 seconds for 20 wards with 6 parallel workers.
