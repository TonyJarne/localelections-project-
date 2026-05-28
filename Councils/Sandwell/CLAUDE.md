# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Analysis of the May 2026 Sandwell local elections — specifically investigating whether the split of left-wing votes between Labour and the Green Party enabled Reform UK to win.

Data source: https://www.sandwell.gov.uk/election-results/candidates-standing-2026-local-elections (24 ward pages at `/2` through `/25`)

## Commands

```bash
# Re-generate the Excel file from scratch
python sandwell_elections.py
```

Requires Python 3 with `pandas` and `openpyxl`:
```bash
pip install pandas openpyxl
```

## Output

`Sandwell_Elections_2026.xlsx` — five sheets:

| Sheet | Contents |
|---|---|
| 1. Raw Results | Every candidate, party, votes, elected status |
| 2. Party Totals | Aggregate votes + % + seats per party |
| 3. Left-Right Blocs | Left (Lab+Green) vs Right (Reform+Con) totals |
| 4. Ward Analysis | Per-ward breakdown with split-vote flag |
| 5. Flourish Data | Cleaned columns ready for copy-paste into Flourish |

## Key Findings (2026)

- **Overall**: Left bloc (Labour + Green) = **50.6%** of all candidate votes; Right bloc (Reform + Conservative) = **46.9%**
- **Seats**: Reform 41, Labour 28, Green 2, Independent 1
- **Split vote**: In **9 of 24 wards**, the combined Labour + Green vote exceeded Reform UK's total, yet Reform won seats — the clearest evidence of vote splitting costing the left councillors

## Data Conventions

- "Votes" = individual candidate votes (each voter can vote for up to 3 candidates in 3-seat wards)
- "Ballots Cast" = number of voters who turned out (≠ sum of all candidate votes)
- Party groupings: Left = Labour + Green; Right = Reform UK + Conservative; Lib Dem and Independents tracked separately
- Fringe parties (TUSC, Yeshua, Advance UK, Socialist and Trade Union) are grouped as "Other"
