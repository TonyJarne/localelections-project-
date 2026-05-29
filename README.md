# UK Local Elections 2026 — Analysis Toolkit

UK Local Elections 2026 toolkit: scrape any council's results from its official website, auto-generate a styled 5-sheet Excel (votes, seats, left/right blocs, split-vote analysis), and visualise with interactive D3.js charts. Includes a Claude `/scrape-elections` command and full Sandwell worked example.

---

## The story behind this project

UK Elections 2026: how the split vote on the left favoured Reform

A few weeks ago, there were local elections in the United Kingdom.

In these elections, the Reform Party won a majority in several local authorities. For example, in Sandwell (West Midlands). How did this happen?

The answer is the First Past The Post system, applied here in its
 multi-member ward variant. 

 Each ward elects three councillors. Each voter has up to three votes and can distribute them freely across any candidates from any party. In practice, most voters use all three votes for the three candidates of their preferred party.

Sandwell had always been an area where Labour had been in power. However, the decline in Labour’s vote, with an increase in votes for the Greens, meant that Reform became the leading party and won 41 of the 72 council seats.

The split in the vote on the left, with a strong only party in the right -Reform- while Tories just got over 10%, meant that in Sandwell, where a larger percentage of people voted for the left, they would have a right-wing council.  

The idea is to show, ward by ward, exactly where and by how much the split vote cost the left seats it should have won on raw vote share alone. 
The result is a set of data tools for scraping council election results, running the analysis, and visualising the gap between votes and seats in each local authority where this happened.

---

## What's included

| File / Folder | Purpose |
|---|---|
| `scrape_council.py` | General scraper: fetches any council's results page and builds the Excel |
| `build_structure.py` | Project scaffolding: creates council folders and master sources Excel |
| `00_Master_Sources.xlsx` | Full list of 2026 councils with known results URLs |
| `Councils/Sandwell/` | Complete worked example: data, scripts, charts |
| `.claude/commands/scrape-elections.md` | Claude `/scrape-elections` command |
| `.claude/skills/run-localelections-project/SKILL.md` | Claude run skill with verified commands |

---

## Quickstart

**Requirements:** Python 3.x, `pip install requests beautifulsoup4 lxml pandas openpyxl`

**Scrape any council:**
```powershell
python -X utf8 scrape_council.py "Sutton" "https://www.sutton.gov.uk/w/local-elections-results-7-may-2026"
```

**Regenerate Sandwell charts:**
```powershell
cd Councils/Sandwell
python generate_charts.py
```

Output Excel has five sheets: Raw Results, Party Totals, Left-Right Blocs, Ward Analysis, and Flourish-ready data for visualisation tools.

---

## Claude integration

If you use [Claude Code](https://claude.ai/code), clone this repo and the `/scrape-elections` command becomes available automatically. It tries the Python scraper first and falls back to WebFetch for councils with non-standard HTML.
