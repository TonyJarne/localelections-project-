---
name: run-localelections-project
description: Run, scrape, and verify the localelections_project UK election analysis pipeline. Use when asked to scrape a council, run the scraper, test the pipeline, generate charts, or screenshot the charts.
---

# Run: localelections-project

This project has two executable pipelines:

1. **Scraper** (`scrape_council.py`) — fetches a UK council election results website and produces a 5-sheet styled Excel file.
2. **Chart generator** (`Councils/Sandwell/generate_charts.py`) — produces 5 standalone D3.js HTML charts from hardcoded Sandwell data.

## Prerequisites

```powershell
pip install requests beautifulsoup4 lxml pandas openpyxl
```

Python 3.14 is installed at `C:\Python314\python.exe`. All packages verified present.

## Run: Scraper (agent path)

```powershell
python -X utf8 scrape_council.py "CouncilName" "https://council-results-url.gov.uk/path"
```

**Verified working on Sutton (2026-05-29):**
```powershell
python -X utf8 scrape_council.py "Sutton" "https://www.sutton.gov.uk/w/local-elections-results-7-may-2026"
```

Output:
- `Councils/<Name>/<Name>_Elections_2026.xlsx` — 5-sheet Excel with styled party colours
- Updates `00_Master_Sources.xlsx` — sets Results Page URL and Data Collected = Yes
- Prints summary: wards scraped, candidates, votes, seats by party, split-vote wards

The scraper handles URL redirects automatically and disables SSL verification (needed on Windows). Run time: ~15–30 seconds for 20 wards.

## Run: JSON fallback mode (councils with unusual HTML)

When the auto-scraper fails (non-table HTML, JavaScript-rendered content, or unusual layouts), collect data via WebFetch and pass pre-extracted JSON:

```powershell
python -X utf8 scrape_council.py "CouncilName" "https://..." --json data.json
```

JSON format (`data.json`):
```json
[
  {
    "ward": "Ward Name",
    "electorate": 10000,
    "ballots": 3500,
    "turnout": 35.0,
    "candidates": [
      {"name": "Candidate Name", "party": "Reform UK", "votes": 1234, "elected": true},
      {"name": "Another Candidate", "party": "Labour Party", "votes": 987, "elected": false}
    ]
  }
]
```

`electorate`, `ballots`, and `turnout` can be `null` if not available. `party` values are auto-normalised (e.g. "The Green Party" → "Green", "Local Conservatives" → "Conservative").

## Run: Chart generator

```powershell
cd Councils\Sandwell
python generate_charts.py
```

Output: 5 HTML files in `Councils/Sandwell/charts/`. Open in any browser — no internet needed, D3.js is embedded.

## Gotchas

- **`-X utf8` is required on Windows** — without it, the checkmark emoji in print statements causes `UnicodeEncodeError`.
- **SSL verification** — Windows Python 3.14 can't verify some UK council SSL certificates. The scraper uses `verify=False` internally and suppresses the resulting warning.
- **Redirect detection** — some council URLs (`/w/short-slug`) redirect to a longer path. The scraper follows redirects and uses the final URL to match ward sub-pages. Always pass the URL as listed in `00_Master_Sources.xlsx`.
- **"Election Candidate" header collision** — councils using "Election Candidate" as their column 0 header would match "elect" before the "Elected/Not Elected" column. The scraper excludes column 0 from the elected-column search.
- **Numbered ward pages (Sandwell-style)** — some councils (e.g. Sandwell) use `/2`, `/3` numeric sub-pages instead of slug sub-pages. The scraper's link detector does not handle these; use `--json` mode or the council-specific `sandwell_elections.py`.
- **Parallel fetch order** — wards are fetched with 6 concurrent threads; order in output may vary. The Excel is sorted alphabetically by ward before writing.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `ERROR: Could not find ward sub-pages` | The main URL path doesn't match ward sub-links. Try the redirected URL directly, or use `--json`. |
| `UnicodeEncodeError: 'charmap'` | Add `-X utf8` flag: `python -X utf8 scrape_council.py ...` |
| `SSLError: CERTIFICATE_VERIFY_FAILED` | Already handled internally with `verify=False`. If you see this, the scraper has a bug — file an issue. |
| `0 seats shown in summary` | The "Elected" column header contains "candidate" — the parser avoided the wrong column. Check the actual HTML table headers. |
| `build_excel ValueError: No candidate data` | All wards returned `None` from parsing — the HTML structure is not table-based. Use `--json` mode. |
