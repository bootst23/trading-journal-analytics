"""Tests for update_data.py - the pipeline that feeds the dashboard and the SQL.

Design notes:
  * The real trading workbook is never used here. Every test that needs a
    workbook builds a small synthetic one in a temp directory, so the suite
    runs anywhere (including CI) and no personal trade data lives in the repo.
  * The one exception is the final test, which reads the committed journal.csv
    to guard the published headline figures against regressions.
"""
import csv
import datetime
import sqlite3
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from update_data import (  # noqa: E402
    PipelineError, build_embed, replace_embed, summarise,
    extract, write_csv, load_database, FIELDNAMES, EMBED_COLUMNS,
)

REPO_ROOT = Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

HEADER = ['Month', 'Direction', 'Date', 'No', 'Result', 'R',
          'V1', 'V2', 'V3', 'V4', 'V5', 'Type', 'Grade']


def make_workbook(path, rows, sheet_name='Master'):
    """Build a minimal .xlsx that mimics the real Master sheet layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def trade_row(month='February', direction='Long', date=None, no=1, result='Win',
              r=2.0, flags=(None,) * 5, ttype='Breakout', grade='B+'):
    date = date or datetime.datetime(2026, 2, 5)
    return [month, direction, date, no, result, r, *flags, ttype, grade]


@pytest.fixture
def sample_trades():
    return [
        {'date': '2026-02-05', 'direction': 'Short', 'result': 'Win', 'r': 1.72,
         'violations': 0, 'type': 'Breakout', 'grade': 'B+'},
        {'date': '2026-02-06', 'direction': 'Long', 'result': 'Loss', 'r': -1,
         'violations': 2, 'type': 'Range-Bound', 'grade': 'C'},
    ]


# ==========================================================================
# build_embed - 4 tests
# ==========================================================================

def test_embed_header_is_the_documented_seven_columns(sample_trades):
    header = build_embed(sample_trades).splitlines()[0]
    assert header == 'date,direction,result,r,violations,type,grade'
    assert header.split(',') == EMBED_COLUMNS


def test_embed_has_one_line_per_trade_plus_header(sample_trades):
    assert len(build_embed(sample_trades).splitlines()) == len(sample_trades) + 1


def test_embed_writes_field_values_in_column_order(sample_trades):
    first_data_line = build_embed(sample_trades).splitlines()[1]
    assert first_data_line == '2026-02-05,Short,Win,1.72,0,Breakout,B+'


def test_embed_round_trips_through_a_csv_reader(sample_trades):
    """What PapaParse will see in the browser must match what we put in."""
    parsed = list(csv.DictReader(build_embed(sample_trades).splitlines()))
    assert len(parsed) == 2
    assert parsed[0]['date'] == '2026-02-05'
    assert float(parsed[1]['r']) == -1
    assert int(parsed[1]['violations']) == 2


# ==========================================================================
# replace_embed - 5 tests
# ==========================================================================

HTML = ('<html><script>\n'
        'const MY_JOURNAL = `date,r\n2026-01-01,1`;\n'
        'function analyse(){}\n</script></html>')


def test_replace_embed_inserts_the_new_data():
    out = replace_embed(HTML, 'date,r\n2026-09-09,5')
    assert '2026-09-09,5' in out
    assert '2026-01-01,1' not in out


def test_replace_embed_leaves_surrounding_html_untouched():
    out = replace_embed(HTML, 'date,r\n2026-09-09,5')
    assert out.startswith('<html><script>')
    assert 'function analyse(){}' in out
    assert out.count('const MY_JOURNAL') == 1


def test_replace_embed_raises_when_anchor_is_missing():
    with pytest.raises(PipelineError, match='found 0'):
        replace_embed('<html>no anchor here</html>', 'date,r\n2026-09-09,5')


def test_replace_embed_raises_when_anchor_is_duplicated():
    with pytest.raises(PipelineError, match='found 2'):
        replace_embed(HTML + HTML, 'date,r\n2026-09-09,5')


def test_replace_embed_rejects_a_backtick_in_the_data():
    """A backtick would terminate the JS template literal and break the page."""
    with pytest.raises(PipelineError, match='backtick'):
        replace_embed(HTML, 'date,r\n2026-09-09,`evil`')


# ==========================================================================
# extract - 6 tests
# ==========================================================================

def test_extract_reads_valid_rows(tmp_path):
    book = make_workbook(tmp_path / 'b.xlsx', [trade_row()])
    trades, warnings = extract(book)
    assert len(trades) == 1
    assert trades[0]['date'] == '2026-02-05'
    assert trades[0]['result'] == 'Win'
    assert warnings == []


def test_extract_skips_non_trade_rows_silently(tmp_path):
    rows = [
        trade_row(),
        ['February', 'Long', 'not a date', 2, 'Win', 1.0, *[None] * 5, 'Breakout', 'B'],
        ['', '', None, None, None, None, *[None] * 5, '', ''],          # blank row
        trade_row(date=datetime.datetime(2026, 2, 7), no=3, result='Summary'),
    ]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, warnings = extract(book)
    assert len(trades) == 1
    assert warnings == []          # expected junk is silent, not noisy


def test_extract_warns_and_skips_a_non_numeric_r(tmp_path):
    rows = [trade_row(), trade_row(date=datetime.datetime(2026, 2, 6), no=2, r='n/a')]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, warnings = extract(book)
    assert len(trades) == 1
    assert len(warnings) == 1
    assert 'not numeric' in warnings[0]


def test_extract_flags_a_month_label_that_disagrees_with_the_date(tmp_path):
    """The typo that once inflated the losing-streak stat from 7 to 8."""
    rows = [trade_row(month='May', date=datetime.datetime(2026, 4, 15), no=1)]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, warnings = extract(book)
    assert len(trades) == 1            # flagged, not dropped
    assert len(warnings) == 1
    assert 'possible typo' in warnings[0]


def test_extract_sorts_by_date_then_trade_number(tmp_path):
    """Eleven dates carry two trades; sequence drives equity and streaks."""
    rows = [
        trade_row(date=datetime.datetime(2026, 2, 12), no=9),
        trade_row(date=datetime.datetime(2026, 2, 10), no=7),
        trade_row(date=datetime.datetime(2026, 2, 12), no=8),
    ]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, _ = extract(book)
    assert [(t['date'], t['trade_no']) for t in trades] == [
        ('2026-02-10', 7), ('2026-02-12', 8), ('2026-02-12', 9)]


def test_extract_counts_violations_across_all_five_flags(tmp_path):
    rows = [trade_row(flags=('x', None, 'x', None, 'x'))]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, _ = extract(book)
    assert trades[0]['violations'] == 3


# ==========================================================================
# load_database - 4 tests
# ==========================================================================

@pytest.fixture
def loaded_db(tmp_path):
    trades = [
        {'date': '2026-02-05', 'month': 'February', 'direction': 'Short', 'trade_no': 1,
         'result': 'Win', 'r': 1.72, 'v1': False, 'v2': False, 'v3': False, 'v4': False,
         'violations': 0, 'type': 'Breakout', 'grade': 'B+', 'mins_in_trade': ''},
        {'date': '2026-02-06', 'month': 'February', 'direction': 'Long', 'trade_no': 2,
         'result': 'Loss', 'r': -1.0, 'v1': True, 'v2': False, 'v3': False, 'v4': False,
         'violations': 1, 'type': 'Range-Bound', 'grade': 'C', 'mins_in_trade': ''},
    ]
    csv_path = write_csv(trades, tmp_path / 'journal.csv')
    db_path = tmp_path / 'journal.db'
    count = load_database(csv_path, db_path)
    return db_path, count, trades


def test_load_database_inserts_every_csv_row(loaded_db):
    db_path, count, trades = loaded_db
    assert count == len(trades)


def test_load_database_stores_r_as_a_real_number(loaded_db):
    """A bare .import types every column TEXT, and '-1' < '0' is false as text."""
    db_path, _, _ = loaded_db
    con = sqlite3.connect(db_path)
    assert con.execute("SELECT typeof(r) FROM trades LIMIT 1").fetchone()[0] == 'real'
    assert con.execute("SELECT COUNT(*) FROM trades WHERE r < 0").fetchone()[0] == 1
    con.close()


def test_load_database_stores_violations_as_an_integer(loaded_db):
    db_path, _, _ = loaded_db
    con = sqlite3.connect(db_path)
    assert con.execute("SELECT typeof(violations) FROM trades LIMIT 1").fetchone()[0] == 'integer'
    con.close()


def test_load_database_is_rerunnable_without_duplicating_rows(loaded_db):
    db_path, _, trades = loaded_db
    csv_path = db_path.parent / 'journal.csv'
    assert load_database(csv_path, db_path) == len(trades)


# ==========================================================================
# regression guard - 1 test
# ==========================================================================

def test_committed_journal_reproduces_the_published_headline_figures():
    """The figures quoted in the README and on the dashboard must not drift."""
    rows = list(csv.DictReader(open(REPO_ROOT / 'journal.csv')))
    trades = [{**r, 'r': float(r['r'])} for r in rows]
    s = summarise(trades)
    assert s['trades'] == 81
    assert s['net_r'] == -6.67
    assert s['win_rate'] == 28.0
    clean = [t for t in trades if int(t['violations']) == 0]
    viol = [t for t in trades if int(t['violations']) > 0]
    assert (len(clean), round(sum(t['r'] for t in clean), 1)) == (51, 17.1)
    assert (len(viol), round(sum(t['r'] for t in viol), 1)) == (30, -23.8)

def test_net_r_rounds_half_up_deterministically():
    """-6.665 must round to -6.67 on every interpreter, not bankers'-round to -6.66."""
    trades = [
        {'r': -3.3325, 'result': 'Loss', 'date': '2026-01-01'},
        {'r': -3.3325, 'result': 'Loss', 'date': '2026-01-02'},
    ]
    assert summarise(trades)['net_r'] == -6.67