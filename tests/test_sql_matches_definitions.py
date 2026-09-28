"""analysis.sql must agree with the README's Definitions section.

Each test runs one query from analysis.sql against a database built from the
committed journal.csv, and recomputes the same figure in plain Python straight
from the CSV. Nothing is hard-coded, so the tests keep passing as trades are added.
One small synthetic journal covers cases the real one cannot (an opening loss).
SQL rounds its output, so each comparison allows half of the last rounded digit.
"""
import csv
import re
import sqlite3
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from update_data import load_database, write_csv  # noqa: E402


def rounded(value, places):
    """Match a Python value against SQL's ROUND(x, places) output."""
    return pytest.approx(value, abs=0.5 * 10 ** -places + 1e-9)


# analysis.sql split once at its '-- N. ' headings. The text before query 1 is
# the setup (the equity_curve view); each numbered block is one query.
_blocks = re.split(r'^(?=-- \d+\. )',
                   (REPO_ROOT / 'analysis.sql').read_text(encoding='utf-8-sig'), flags=re.M)
SETUP = _blocks[0]
QUERIES = {int(re.match(r'-- (\d+)\. ', b).group(1)): b for b in _blocks[1:]}


def sql_query(number):
    """Return the text of query `number` (the block under '-- N. ...')."""
    return QUERIES[number]


def build_db(csv_path, db_path):
    """Load the trades table from a CSV, then create the views analysis.sql defines."""
    load_database(csv_path, db_path)
    con = sqlite3.connect(db_path)
    con.executescript(SETUP)
    return con


@pytest.fixture(scope='module')
def db(tmp_path_factory):
    con = build_db(REPO_ROOT / 'journal.csv', tmp_path_factory.mktemp('db') / 'journal.db')
    yield con
    con.close()


@pytest.fixture(scope='module')
def trades():
    rows = list(csv.DictReader(open(REPO_ROOT / 'journal.csv', newline='')))
    for t in rows:
        t['r'] = float(t['r'])
        t['violations'] = int(t['violations'])
    return sorted(rows, key=lambda t: (t['date'], int(t['trade_no'])))


def test_headline_query_matches_definitions(db, trades):
    total, wins, losses, be, win_rate, net_r, expectancy = db.execute(sql_query(1)).fetchone()
    n_win = sum(t['result'] == 'Win' for t in trades)
    n_loss = sum(t['result'] == 'Loss' for t in trades)
    net = sum(t['r'] for t in trades)

    assert (total, wins, losses, be) == (
        len(trades), n_win, n_loss, sum(t['result'] == 'BE' for t in trades))
    assert win_rate == rounded(100 * n_win / (n_win + n_loss), 1)
    assert net_r == rounded(net, 2)
    assert expectancy == rounded(net / len(trades), 3)


def test_compliance_query_matches_definitions(db, trades):
    rows = {label: (n, net_r, exp) for label, n, net_r, exp in db.execute(sql_query(3))}
    groups = {'rule-compliant': [t for t in trades if t['violations'] == 0],
              'violated rules': [t for t in trades if t['violations'] > 0]}

    assert set(rows) == set(groups)
    for label, group in groups.items():
        n, net_r, exp = rows[label]
        net = sum(t['r'] for t in group)
        assert n == len(group)
        assert net_r == rounded(net, 2)
        assert exp == rounded(net / len(group), 3)


def test_max_drawdown_query_matches_definition(db, trades):
    cum = peak = worst = 0.0                 # balance and peak both start at 0R
    for t in trades:
        cum += t['r']
        peak = max(peak, cum)
        worst = min(worst, cum - peak)

    assert db.execute(sql_query(8)).fetchone()[0] == rounded(worst, 2)


def test_equity_curve_view_matches_python_row_by_row(db, trades):
    cum = peak = 0.0                         # balance and peak both start at 0R
    expected = []
    for t in trades:
        cum += t['r']
        peak = max(peak, cum)
        expected.append((t['date'], int(t['trade_no']), cum, peak, cum - peak))
    got = db.execute("SELECT date, trade_no, cum_r, peak_r, drawdown_r "
                     "FROM equity_curve ORDER BY date, trade_no").fetchall()

    assert len(got) == len(expected)
    for g, e in zip(got, expected):
        assert g[:2] == e[:2]
        assert g[2:] == pytest.approx(e[2:]), f'{e[0]} trade #{e[1]}'


def test_max_drawdown_counts_early_and_same_day_losses(tmp_path):
    """Two losses on day one, then a win: the peak is 0R, not the first trade or the day."""
    trades = [{'date': '2026-01-05', 'trade_no': 1, 'result': 'Loss', 'r': -1},
              {'date': '2026-01-05', 'trade_no': 2, 'result': 'Loss', 'r': -1},
              {'date': '2026-01-06', 'trade_no': 3, 'result': 'Win', 'r': 3}]
    con = build_db(write_csv(trades, tmp_path / 'journal.csv'), tmp_path / 'journal.db')

    assert con.execute(sql_query(8)).fetchone()[0] == -2.0
    con.close()


def test_losing_streak_query_matches_definition(db, trades):
    streak = longest = 0
    for t in trades:
        streak = streak + 1 if t['result'] == 'Loss' else 0   # a win or BE ends it
        longest = max(longest, streak)

    assert db.execute(sql_query(9)).fetchone()[0] == longest


@pytest.mark.parametrize('number', [5, 10])
def test_monthly_queries_group_by_the_month_of_the_date(db, trades, number):
    """One row per calendar month taken from the date, in order, with its trade count."""
    expected = {}
    for t in trades:
        expected[t['date'][:7]] = expected.get(t['date'][:7], 0) + 1

    got = [(month, n) for month, n, *_ in db.execute(sql_query(number))]
    assert got == sorted(expected.items())
