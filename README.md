# UK Local Elections 2026 — Analysis Toolkit

UK Local Elections 2026 toolkit: scrape any council's results from its official website, auto-generate a styled 5-sheet Excel (votes, seats, left/right blocs, split-vote analysis), and visualise with interactive D3.js charts. Includes a Claude `/scrape-elections` command and full Sandwell worked example.

---

## The story behind this project

On 7 May 2026, Reform UK swept Sandwell Metropolitan Borough Council, winning 41 of 72 seats and taking control of a council that had been Labour-held for decades. The result made national headlines. But the headline number hid something striking.

When you add up every vote cast across all 24 wards, the left — Labour and the Green Party combined — got **50.6% of the vote**. Reform UK got **36.3%**. The left outpolled Reform by more than 14 percentage points. And yet Reform walked away with 41 seats to Labour's 28 and the Greens' 2.

How does a bloc that wins the majority of votes end up with fewer than half the seats?

The answer is England's **First Past The Post** system, applied here in its multi-member ward variant. Each ward elects three councillors. Each voter has up to three votes and can distribute them freely across any candidates from any party. In practice, most voters use all three votes for the three candidates of their preferred party.

This is where the left's problem becomes structural. A Reform voter faces no dilemma: they vote for all three Reform candidates and move on. A left-leaning voter, however, has to choose — vote for all three Labour candidates and ignore the Greens, vote for all three Greens and ignore Labour, or split their votes across both parties. Whatever they decide, the left's total vote is being divided across six candidates instead of three. Each individual Labour and Green candidate receives fewer votes than they would if the two parties were running as one. Reform's three candidates, even with a minority of the total votes cast, consistently rank above them.

In 9 of Sandwell's 24 wards, Labour and Green votes combined exceeded Reform's total. Reform still won seats in every one of them.

This project started as an attempt to quantify that dynamic precisely — to move beyond "Reform won big" and show, ward by ward, exactly where and by how much the split vote cost the left seats it should have won on raw vote share alone. The result is a set of data tools for scraping council election results, running the analysis, and visualising the gap between votes and seats.

Sandwell is the worked example. The tools work for any council.

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
