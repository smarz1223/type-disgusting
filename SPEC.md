# Type Disgusting Fantasy Hockey: site spec

Repo `smarz1223/type-disgusting` | Site `typedisgusting.redplanetanalytics.com` | Yahoo league 24054
Repo code is the source of truth. If this file disagrees with `fetch_data.py`, the code wins.

## League format (2026-27)
- 8 teams, Yahoo head-to-head points, keeper league. No median game.
- Regular season weeks 1-23. Playoffs: top 4, weeks 24-25, reseeding on, 3rd place game (no payout).
- Standings tiebreaker: points for. Playoff tiebreaker (Yahoo): better regular season record.
- Payouts: $100 fee, $800 pool. Champion $500, 2nd $100, regular season 1st $200.
- Scoring: G 15, A 10, +/- 5, PIM 2.5, PPP 5, SOG 2 | W 20, GA -5, SV 1, SHO 10.

## Owners (display names, site shows owners only)
| Owner | 2026-27 Yahoo team |
|---|---|
| PHIL | Cat Plark |
| MARZ | Soft Dump in Corner |
| ADAM B | Wild Hogs |
| PEDRO | Panarin Bread |
| PAT CLARK | H DUBZ ON THA TRACK |
| SCOTTY | Bros Before Hossas |
| FUR | Ladies...Beverages |
| MEANY | Sloppy Seconds (first season) |

History mapping: Scott Ratto = SCOTTY, current owners uppercased, former managers shown as-is
(no merging, e.g. "Dave Burt" and "davemuthafuckinburt" stay separate). Charlie is inactive.
Seasons display as start year to end year (history year 2025 = 2025-26).

## Data sources (published Google Sheets .xlsx)
- Stats workbook tabs used:
  - `Free For All`: Week, Team (Yahoo team name), PF, PA. Typed by hand. Opponents are inferred
    by mirrored PF/PA. Team names map to owners via `TEAM_TO_OWNER` in `fetch_data.py`.
  - `SKATER PULL`, `GOALIE PULL`: IMPORTHTML of each team's Yahoo team log (season to date,
    active lineup only, refreshes daily). One 12-column block per owner, label in row 1.
  - `Data Pull`: Yahoo league stat totals through completed matchups. Used only to verify
    Free For All scores.
- Ignored tabs: Summary, Skaters, Goalies, Games Played, TAKEN DEFENSEMEN, TAKEN FWDD.
- History workbook: `Data` tab only (Years, Manager, Status, 1st, 2nd, 3rd, Wins, Losses, Ties, Games, Playoffs).

## Data check (badge in the header)
- PASS: Data Pull points tie to cumulative Free For All PF for every team through the last week entered.
- PARTIAL: ties through an earlier week (Yahoo table hasn't caught up).
- CHECK: ties for no week (likely a typo in a score).
- FAIL: opponent not matched, unknown team name, or parsed log rows don't equal Yahoo Totals rows.
- Team log points run ahead of Free For All during a live week. That gap is shown as "Current week so far" in `data/recon_report.md`.

## Pages
Standings, Weekly, Stat Categories, Positions (Forwards = C/LW/RW, Defensemen, Goalies),
Teams, League History (with Records), Rules (Yahoo settings cited).

## Design
NHL '94 Sega Genesis: navy arena header with logo + LED jumbotron, rink-line stripe,
Press Start 2P headings/nav, VT323 scoreboard digits, Chakra Petch body, light ice background
(heat maps stay readable), hard pixel shadows. Palette: arena #0B1650, blue #2547D0,
red #E3261B, lamp #FFC21A, ice #EAF1F9.

## Rollover checklist (each new season)
- Update `SEASON`, `OWNERS`, `TEAM_TO_OWNER`, week counts, and scoring if Yahoo settings changed.
- Point `STATS_URL` at the new season's published workbook.
- Add the finished season to the history workbook `Data` tab.
- Recheck the Rules page text against Yahoo settings and payouts.
