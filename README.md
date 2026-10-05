# Trading Journal Analytics

SQL and dashboard analysis of my own trading journal. 101 crypto futures trades
(BTC/ETH/SOL), Feb-Oct 2026, logged in Excel with direction, result, R-multiple,
a setup grade (A+ to D) assigned before entry, five rule-violation flags, and
trade type.

## Stats

Overall:
- 101 trades, net -4.1R
- Win rate 31.5% (excl. break-even)
- Avg win +1.88R, avg loss -0.93R (payoff 2.0)
- Expectancy -0.04R per trade
- Max drawdown -20.2R, longest losing streak 7

By rule compliance:
- 0 violations, 59 trades: +24.8R total, +0.42R per trade
- 1+ violations, 42 trades: -28.9R total, -0.69R per trade

By setup grade (assigned at entry, before the outcome was known):
- A, 11 trades: +13.0R, +1.18R per trade, 70% win rate
- B, 50 trades: -6.9R, -0.14R per trade, 32%
- C, 34 trades: -4.6R, -0.14R per trade, 25%
- D, 6 trades: -5.5R, -0.92R per trade, 0%

By trade type:
- Breakout, 58 trades: +2.5R total
- Range-bound, 43 trades: -6.5R total

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
- **Average time in trade / in drawdown**: the mean of "D. in trade" / "D. in
  drawdown" in minutes, for winners and for losers separately. Break-even trades
  are left out, and trades with no readable time are skipped, not counted as 0.
  Shown as hours and minutes, with the number of trades averaged.
- **Execution-grade stats**: only trades that have an execution grade, grouped by
  its letter: win rate (break-even excluded) and average R per trade, labelled with
  the number of graded trades.
- **Rounding**: figures are rounded half away from zero, so -6.665 shows as -6.67,
  the same in Python, SQL and the dashboard.

## Files

| File | What it does |
|---|---|
| `journal.csv` | Per-trade extract from the workbook |
| `analysis.sql` | 10 SQLite queries: headline stats, payoff, compliance split, grade expectancy, monthly breakdowns, equity/drawdown (window functions), losing streaks |
| `update_data.py` | Reads the workbook's Master sheet by header name, warns when a trade's typed Month label disagrees with its date (the date is the source of truth; the label is not exported), rewrites the CSV and the dashboard's embedded data; `--db-only` builds `journal.db` from the CSV |
| `dashboard.html` | Single-file dashboard (Chart.js); shows the trade data that `update_data.py` embeds in it |

## Run it

Open `dashboard.html` in a browser. It shows the data embedded in it.

To refresh after logging new trades, run the pipeline on the workbook. It
validates the data, rewrites `journal.csv` and updates the data inside
`dashboard.html`; then reload the page:

```
python update_data.py path\to\Trading_Months_Stats.xlsx
```

Keep the workbook outside the repo folder: it holds personal financial sheets.
`*.xlsx` is in `.gitignore` as a backstop, not the plan.

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
