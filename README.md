# Season Card Collection

A community-friendly framework for building an affordable,
season-by-season historical baseball card collection for any team.

The project focuses on the players who contributed to each
season---including obscure and forgotten players---rather than
completing a manufacturer's set or maximizing card value. The priorities
are historical accuracy, consistency, affordability, and the enjoyment
of collecting.

## Getting Started

1.  Open [`Instructions.pdf`](Instructions.pdf) for the full collection
    rules, player-selection criteria, card-selection preferences, and
    binder organization guidance.
2.  Open the checklist for the team and season you want to collect.
3.  Use the checklist to identify the players and arrange cards in the
    listed order. 
4.  Track your collection using the accompanying spreadsheet, if one is
    provided.
5.  Leave a checklist slot empty until you find a suitable card. You do
    not need to fill every slot immediately.

The way I organized this is that I created 3 binders, each with about 20 years of seasons. 1970-1990, 1991-2010, 2011-Current.

I then printed my team's checklist and use full protector sheet sleeves to use as the first page for each of my team's seasons.

## Collection Rules at a Glance

The complete rules are in **[Instructions.pdf](Instructions.pdf)**. In
summary:

-   Each season has 27 player slots: 18 position players and 9 pitchers.
-   Position players with at least one non-pitcher appearance are ranked by
    games played, with plate appearances (PA) used to break ties. A player
    with a token pitching appearance is still treated as a position player
    when Lahman records them at another position.
-   Pitchers are ranked by innings pitched, with strikeouts (SO) used to
    break ties.
-   Use regular-season statistics accumulated with the team being
    documented; exclude postseason statistics.
-   Cards from any manufacturer, set, type, or release year are
    acceptable. Certain card characteristics are preferences, not
    requirements.
-   A player may appear in multiple seasons, but the same physical card
    should not be reused for that player across seasons.
-   Document corrections and exceptions in the master spreadsheet, with
    supporting sources where possible.

Refer to `Instructions.pdf` for the full wording and details. If this
README and the PDF ever differ, please report the discrepancy so the
project documentation can be corrected.

## Reporting an Error or Suggesting an Improvement

Historical and statistical accuracy matter. If you find an incorrect
ranking, player detail, team record, or other issue, please open an
issue or submit a pull request. 


## `card_checklists.py`

`card_checklists.py` generates printable, three-page PDF checklists using
regular-season data from the Lahman Baseball Database. Data is loaded through
the `pylahman` package, so an internet connection may be required the first
time the script downloads the data.

### Setup

From this directory, install the required Python packages:

```bash
python -m pip install pylahman pandas openpyxl reportlab
```

Then run the script:

```bash
python card_checklists.py
```

The default command generates a PDF for every MLB team in the Lahman data and
writes the files to the `checklists/` directory. Existing files with the same
name may be replaced.

### Examples

Generate checklists for one team and a specific range of seasons:

```bash
python card_checklists.py --team-id NYY --start 1970 --end 2025
```

Generate all teams for a specific range of seasons:

```bash
python card_checklists.py --start 1970 --end 2025
```

Write the generated PDFs to a different directory:

```bash
python card_checklists.py --team-id PIT --start 1970 --end 2025 --output output
```

Available options:

- `--team-id`: Lahman MLB team ID, such as `NYY` or `PIT`. Defaults to all MLB teams.
- `--start`: Earliest season to include. Defaults to each team's first available season.
- `--end`: Latest season to include. Defaults to each team's latest available season.
- `--output`: Destination directory. Defaults to `checklists`.

The generated PDFs contain 27 checklist slots per season: 18 position players
and 9 pitchers. Position-player slots include players with a recorded
non-pitcher appearance; players are ranked using the rules described in
[`Instructions.pdf`](Instructions.pdf). Postseason statistics are not included.
Season sheets are ordered from the latest included season back to the earliest.
When the requested range includes 1970 through 1979, the PDF also ends with an
alternate `1970-79 Season` sheet containing aggregated numbers for that period.