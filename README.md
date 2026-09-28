# Trading Journal Analytics

SQL and dashboard analysis of my own trading journal. 81 crypto futures trades
(BTC/ETH/SOL), Feb-Jun 2026, logged in Excel with direction, result, R-multiple,
a setup grade (A+ to D) assigned before entry, five rule-violation flags, and
trade type.

## Stats

Overall:
- 81 trades, net -6.7R
- Win rate 28.0% (excl. break-even)
- Avg win +2.06R, avg loss -0.92R (payoff 2.2)
- Expectancy -0.08R per trade
- Max drawdown -18.5R, longest losing streak 7

By rule compliance:
- 0 violations, 51 trades: +17.1R total, +0.34R per trade
- 1+ violations, 30 trades: -23.8R total, -0.79R per trade

By setup grade (assigned at entry, before the outcome was known):
- A, 6 trades: +8.6R, +1.44R per trade, 67% win rate
- B, 44 trades: -8.8R, -0.20R per trade, 28%
- C, 26 trades: -2.0R, -0.08R per trade, 24%
- D, 5 trades: -4.5R, -0.90R per trade, 0%

By trade type:
- Breakout, 51 trades: +0.03R total
- Range-bound, 30 trades: -6.7R total

## What the numbers say

The losses didn't come from the strategy. Trades where I followed my rules were
profitable on their own. The damage came from the 30 trades where I broke a rule,
and from volume in setups I had already graded as mediocre: only 6 of 81 trades
were A-grade, and A was the only grade with positive expectancy. Nearly all of the
net loss sits in range-bound trades; breakouts roughly broke even.

Extraction also caught a mistyped date in the workbook (a trade entered as May
that belonged to April). Before the fix it inflated the losing-streak stat from 7
to 8 and put one trade in the wrong month.

## Definitions

Trades are taken in order of date, then trade number within the day. R is a
trade's result as a multiple of the amount risked.

- **Net R**: the sum of R over all trades.
- **Win rate (excl. break-even)**: wins / (wins + losses) x 100. Break-even trades
  are left out of both counts.
- **Expectancy**: net R / number of trades, break-even trades included.
- **Payoff ratio**: average R of winning trades / absolute average R of losing
  trades.
- **Max drawdown**: keep a running total of R, starting from a balance of 0R. The
  peak is the highest running total so far, also starting at 0R, so losses on the
  very first trades count. Drawdown is running total minus peak; max drawdown is
  its lowest value.
- **Longest losing streak**: the most losses in a row. A win or a break-even ends
  a streak.

## Files

| File | What it does |
|---|---|
| `journal.csv` | Per-trade extract from the workbook |
| `analysis.sql` | 10 SQLite queries: headline stats, payoff, compliance split, grade expectancy, monthly breakdowns, equity/drawdown (window functions), losing streaks |
| `update_data.py` | Reads the workbook's Master sheet, validates dates, rewrites the CSV and the dashboard's embedded data; `--db-only` builds `journal.db` from the CSV |
| `dashboard.html` | Single-file dashboard (Chart.js + SheetJS); opens with the data embedded, can load the original .xlsx or a CSV |

## Run it

Open `dashboard.html` in a browser. To load fresh data, click *Load journal* - it
takes the original `.xlsx` workbook or a CSV with columns:
`date, direction, result (Win/Loss/BE), r, violations, type, grade`.

For the SQL, build `journal.db` from the committed `journal.csv` (no workbook
needed), then run the queries. Both lines work in PowerShell, cmd and bash:

```
python update_data.py --db-only
sqlite3 -header -column journal.db ".read analysis.sql"
```

Use `pip install -r requirements.txt` first. `sqlite3` is the SQLite command-line
tool (sqlite.org/download.html).

---
*Nathan Yap Hao Kang · nathan7657@gmail.com*
