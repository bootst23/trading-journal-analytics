"""Tests for update_data.py - the pipeline that feeds the dashboard and the SQL.

Design notes:
  * The real trading workbook is never used here. Every test that needs a
    workbook builds a small synthetic one in a temp directory, so the suite
    runs anywhere (including CI) and no personal trade data lives in the repo.
  * The exceptions read the committed journal.csv (and dashboard.html): the
    --db-only test, and the regression guards at the end, which catch drift.
"""
import csv
import datetime
import json
import os
import shutil
import sqlite3
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from update_data import (  # noqa: E402
    PipelineError, build_embed, replace_embed, summarise,
    extract, write_csv, load_database, parse_duration, FIELDNAMES, EMBED_COLUMNS, EMBED_PATTERN,
)

REPO_ROOT = Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

# Mirrors the real Master sheet, including the trailing space on 'Trade No. '.
HEADER = ['Month', 'Direction', 'Date', 'Trade No. ', 'Result', 'R',
          'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'Type', 'Setup Grade',
          'Execution Grade', 'D. in trade', 'D. in drawdown', 'D. helper']


def make_workbook(path, rows, sheet_name='Master', header=HEADER):
    """Build a minimal .xlsx that mimics the real Master sheet layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def trade_row(month='February', direction='Long', date=None, no=1, result='Win',
              r=2.0, flags=(None,) * 8, ttype='Breakout', grade='B+',
              exec_grade=None, in_trade='1h 5m', in_drawdown='20m', helper=None):
    date = date or datetime.datetime(2026, 2, 5)
    return [month, direction, date, no, result, r, *flags, ttype, grade,
            exec_grade, in_trade, in_drawdown, helper]


@pytest.fixture
def sample_trades():
    return [
        {'date': '2026-02-05', 'trade_no': 1, 'direction': 'Short', 'result': 'Win', 'r': 1.72,
         'violations': 0, 'type': 'Breakout', 'grade': 'B+', 'exec_grade': 'A',
         'mins_in_trade': 65, 'mins_in_drawdown': 20},
        {'date': '2026-02-06', 'trade_no': 2, 'direction': 'Long', 'result': 'Loss', 'r': -1,
         'violations': 2, 'type': 'Range-Bound', 'grade': 'C', 'exec_grade': '',
         'mins_in_trade': '', 'mins_in_drawdown': 0},
    ]


# ==========================================================================
# build_embed - 4 tests
# ==========================================================================

def test_embed_is_one_object_per_trade_with_the_embed_columns(sample_trades):
    """The embed is a JS array literal, which is also valid JSON."""
    parsed = json.loads(build_embed(sample_trades))
    assert len(parsed) == len(sample_trades)
    assert all(list(row) == EMBED_COLUMNS for row in parsed)


def test_embed_writes_every_value_as_a_string(sample_trades):
    assert json.loads(build_embed(sample_trades))[0] == {
        'date': '2026-02-05', 'trade_no': '1', 'direction': 'Short', 'result': 'Win', 'r': '1.72',
        'violations': '0', 'type': 'Breakout', 'grade': 'B+', 'exec_grade': 'A',
        'mins_in_trade': '65', 'mins_in_drawdown': '20'}
    assert json.loads(build_embed(sample_trades))[1]['mins_in_trade'] == ''   # blank stays blank, not 0


def test_embed_keeps_commas_quotes_and_newlines_inside_a_field(sample_trades):
    """A comma in a CSV embed would shift every later column without an error."""
    note = 'trailed stop, took half off "early"\\ `then` ${x}\nnext line'
    sample_trades[0]['type'] = note
    row = json.loads(build_embed(sample_trades))[0]
    assert row['type'] == note
    assert row['r'] == '1.72'                        # the columns after it did not move


def test_embed_cannot_close_the_script_tag(sample_trades):
    sample_trades[0]['grade'] = '</script><b>'
    embed = build_embed(sample_trades)
    assert '<' not in embed
    assert json.loads(embed)[0]['grade'] == '</script><b>'


@pytest.mark.parametrize('blank', [None, ''])
def test_embed_writes_a_blank_cell_as_an_empty_string_not_null(sample_trades, blank):
    """A blank workbook cell is None; it must embed as '' (as journal.csv has it), not null or 'None'."""
    sample_trades[0]['type'] = blank
    assert json.loads(build_embed(sample_trades))[0]['type'] == ''


def test_embed_keeps_non_ascii_and_writes_it_as_ascii_escapes(sample_trades):
    sample_trades[0]['type'] = 'Range-Bound — été →'
    embed = build_embed(sample_trades)
    assert embed.isascii()          # json.dumps ensure_ascii default: \uXXXX escapes
    assert json.loads(embed)[0]['type'] == 'Range-Bound — été →'


# ==========================================================================
# replace_embed - 4 tests
# ==========================================================================

OLD = build_embed([{c: 'old' for c in EMBED_COLUMNS}])
NEW = build_embed([{c: 'new' for c in EMBED_COLUMNS}])
HTML = ('<html><script>\n'
        f'const MY_JOURNAL = {OLD};\n'
        'function analyse(){}\n</script></html>')


def test_replace_embed_inserts_the_new_data():
    out = replace_embed(HTML, NEW)
    assert NEW in out
    assert OLD not in out


def test_replace_embed_leaves_surrounding_html_untouched():
    out = replace_embed(HTML, NEW)
    assert out.startswith('<html><script>')
    assert 'function analyse(){}' in out
    assert out.count('const MY_JOURNAL') == 1


@pytest.mark.parametrize('value', ['];', 'stop moved ]; then out'])
def test_a_field_holding_the_terminator_does_not_end_the_block(sample_trades, value):
    """EMBED_PATTERN ends at a line that is only '];'; no field value can produce one."""
    sample_trades[0]['type'] = value
    embed = build_embed(sample_trades)
    once = replace_embed(HTML, embed)
    assert EMBED_PATTERN.findall(once) == ['const MY_JOURNAL = ' + embed + ';']
    assert replace_embed(once, NEW) == HTML.replace(OLD, NEW)     # the whole block goes, nothing left over


def test_replace_embed_raises_when_anchor_is_missing():
    with pytest.raises(PipelineError, match='found 0'):
        replace_embed('<html>no anchor here</html>', NEW)


def test_replace_embed_raises_when_anchor_is_duplicated():
    with pytest.raises(PipelineError, match='found 2'):
        replace_embed(HTML + HTML, NEW)


# ==========================================================================
# extract - 14 tests
# ==========================================================================

def test_extract_reads_valid_rows(tmp_path):
    book = make_workbook(tmp_path / 'b.xlsx', [trade_row()])
    trades, warnings = extract(book)
    assert len(trades) == 1
    assert trades[0]['date'] == '2026-02-05'
    assert trades[0]['result'] == 'Win'
    assert trades[0]['type'] == 'Breakout'
    assert trades[0]['grade'] == 'B+'
    assert warnings == []


def test_extract_skips_non_trade_rows_silently(tmp_path):
    rows = [
        trade_row(),
        ['February', 'Long', 'not a date', 2, 'Win', 1.0, *[None] * 8, 'Breakout', 'B'],
        ['', '', None, None, None, None, *[None] * 8, '', ''],          # blank row
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
    assert 'month' not in trades[0]    # the date is the source of truth


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


def test_extract_counts_violations_across_all_eight_flags(tmp_path):
    rows = [trade_row(flags=('x', None, 'x', None, 'x', None, 'x', 'x'))]
    book = make_workbook(tmp_path / 'b.xlsx', rows)
    trades, _ = extract(book)
    assert trades[0]['violations'] == 5


def test_extract_reads_columns_by_header_in_any_order(tmp_path):
    order = [15, 3, 19, 0, 7, 14, 2, 10, 5, 1, 17, 12, 4, 8, 16, 6, 11, 9, 18, 13]
    row = trade_row(flags=('x', None, None, 'x', None, None, 'x', None))
    book = make_workbook(tmp_path / 'b.xlsx', [[row[i] for i in order]],
                         header=[HEADER[i] for i in order])
    trades, warnings = extract(book)
    assert warnings == []
    t = trades[0]
    assert (t['date'], t['direction'], t['trade_no'], t['result'], t['r']) == (
        '2026-02-05', 'Long', 1, 'Win', 2.0)
    assert (t['v1'], t['v2'], t['v3'], t['v4']) == (True, False, False, True)
    assert (t['violations'], t['type'], t['grade']) == (3, 'Breakout', 'B+')


def test_extract_raises_naming_missing_required_headers(tmp_path):
    header = [h for h in HEADER if h not in ('Setup Grade', 'Trade No. ')]
    book = make_workbook(tmp_path / 'b.xlsx', [], header=header)
    with pytest.raises(PipelineError, match=r'Trade No\., Setup Grade'):
        extract(book)


@pytest.mark.parametrize('value, minutes', [
    ('1d 10h 59m', 2099), ('3h', 180), ('15m', 15), ('18h29m', 1109),
    ('0', 0), ('0m', 0), (0, 0)])
def test_parse_duration_reads_days_hours_and_minutes(value, minutes):
    assert parse_duration(value) == minutes


@pytest.mark.parametrize('value', ['', None, 'soon', '3 hours', 5])
def test_parse_duration_rejects_anything_else(value):
    with pytest.raises(ValueError):
        parse_duration(value)


def test_extract_stores_durations_as_whole_minutes(tmp_path):
    rows = [trade_row(in_trade='1d 10h 59m', in_drawdown='18h29m', helper='999')]
    trades, warnings = extract(make_workbook(tmp_path / 'b.xlsx', rows))
    assert (trades[0]['mins_in_trade'], trades[0]['mins_in_drawdown']) == (2099, 1109)
    assert warnings == []             # D. helper is ignored, whatever it holds


def test_extract_warns_and_blanks_an_unreadable_duration(tmp_path):
    rows = [trade_row(no=4, in_trade='soon')]
    trades, warnings = extract(make_workbook(tmp_path / 'b.xlsx', rows))
    assert len(trades) == 1                          # kept, not dropped
    assert trades[0]['mins_in_trade'] == ''
    assert trades[0]['mins_in_drawdown'] == 20       # the other field is still read
    assert len(warnings) == 1
    assert '2026-02-05' in warnings[0] and '#4' in warnings[0] and 'D. in trade' in warnings[0]


def test_extract_leaves_a_blank_execution_grade_blank(tmp_path):
    rows = [trade_row(no=1, exec_grade=None), trade_row(no=2, exec_grade='  '),
            trade_row(no=3, exec_grade='B')]
    trades, _ = extract(make_workbook(tmp_path / 'b.xlsx', rows))
    assert [t['exec_grade'] for t in trades] == ['', '', 'B']


def test_extract_stores_all_eight_flags_as_1_or_0(tmp_path):
    rows = [trade_row(flags=('x', None, None, None, None, None, None, 'x'))]
    trades, _ = extract(make_workbook(tmp_path / 'b.xlsx', rows))
    assert [trades[0][f'v{i}'] for i in range(1, 9)] == [1, 0, 0, 0, 0, 0, 0, 1]


# ==========================================================================
# load_database - 7 tests
# ==========================================================================

@pytest.fixture
def loaded_db(tmp_path):
    trades = [
        {'date': '2026-02-05', 'direction': 'Short', 'trade_no': 1,
         'result': 'Win', 'r': 1.72, **{f'v{i}': 0 for i in range(1, 9)},
         'violations': 0, 'type': 'Breakout', 'grade': 'B+', 'exec_grade': 'A',
         'mins_in_trade': 65, 'mins_in_drawdown': 20},
        {'date': '2026-02-06', 'direction': 'Long', 'trade_no': 2,
         'result': 'Loss', 'r': -1.0, **{f'v{i}': int(i == 1) for i in range(1, 9)},
         'violations': 1, 'type': 'Range-Bound', 'grade': 'C', 'exec_grade': '',
         'mins_in_trade': '', 'mins_in_drawdown': 0},
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


def test_load_database_stores_a_blank_as_null_not_zero(loaded_db):
    """AVG() skips NULL but would count a 0; a real 0 (no drawdown) must stay 0."""
    db_path, _, _ = loaded_db
    con = sqlite3.connect(db_path)
    row = con.execute("SELECT mins_in_trade, mins_in_drawdown, exec_grade "
                      "FROM trades WHERE trade_no = 2").fetchone()
    assert row == (None, 0, None)
    assert con.execute("SELECT AVG(mins_in_trade) FROM trades").fetchone()[0] == 65
    con.close()


def test_load_database_stores_flags_as_integers(loaded_db):
    db_path, _, _ = loaded_db
    con = sqlite3.connect(db_path)
    assert con.execute("SELECT typeof(v1), v1, v8 FROM trades WHERE trade_no = 2").fetchone() == (
        'integer', 1, 0)
    assert con.execute("SELECT SUM(v1) FROM trades").fetchone()[0] == 1
    con.close()


def test_load_database_is_rerunnable_without_duplicating_rows(loaded_db):
    db_path, _, trades = loaded_db
    csv_path = db_path.parent / 'journal.csv'
    assert load_database(csv_path, db_path) == len(trades)


def test_db_only_builds_the_database_from_journal_csv_alone(tmp_path):
    """A fresh clone has journal.csv but no workbook and must still get journal.db."""
    shutil.copy(REPO_ROOT / 'update_data.py', tmp_path)
    shutil.copy(REPO_ROOT / 'journal.csv', tmp_path)       # no workbook, no dashboard
    n = len(list(csv.DictReader(open(REPO_ROOT / 'journal.csv', newline=''))))
    out = subprocess.run([sys.executable, 'update_data.py', '--db-only'], cwd=tmp_path,
                         capture_output=True, text=True, check=True).stdout
    assert f'Loaded {n} trades into journal.db' in out
    con = sqlite3.connect(tmp_path / 'journal.db')
    assert con.execute("SELECT COUNT(*) FROM trades").fetchone()[0] == n
    assert con.execute("SELECT typeof(r) FROM trades LIMIT 1").fetchone()[0] == 'real'
    con.close()


def test_running_update_data_twice_writes_byte_identical_files(tmp_path):
    shutil.copy(REPO_ROOT / 'update_data.py', tmp_path)
    shutil.copy(REPO_ROOT / 'dashboard.html', tmp_path)
    make_workbook(tmp_path / 'b.xlsx', [
        trade_row(ttype='trailed stop, took half off "early" ];'),
        trade_row(date=datetime.datetime(2026, 2, 6), no=2, result='Loss', r=-1, ttype=None)])
    # The summary line prints '→', which a cp1252 pipe on Windows cannot encode.
    env = {**os.environ, 'PYTHONIOENCODING': 'utf-8'}
    runs = []
    for _ in range(2):
        subprocess.run([sys.executable, 'update_data.py', 'b.xlsx'], cwd=tmp_path, env=env,
                       capture_output=True, check=True)
        runs.append([(tmp_path / f).read_bytes() for f in ('dashboard.html', 'journal.csv')])
    assert runs[0] == runs[1]


# ==========================================================================
# regression guards - 3 tests
# ==========================================================================

def test_committed_journal_reproduces_the_published_headline_figures():
    """The figures quoted in the README and on the dashboard must not drift."""
    rows = list(csv.DictReader(open(REPO_ROOT / 'journal.csv')))
    trades = [{**r, 'r': float(r['r'])} for r in rows]
    s = summarise(trades)
    assert s['trades'] == 101
    assert s['net_r'] == -4.05
    assert s['win_rate'] == 31.5
    clean = [t for t in trades if int(t['violations']) == 0]
    viol = [t for t in trades if int(t['violations']) > 0]
    assert (len(clean), round(sum(t['r'] for t in clean), 1)) == (59, 24.8)
    assert (len(viol), round(sum(t['r'] for t in viol), 1)) == (42, -28.9)


RERUN = ("Rerun `python update_data.py <workbook>` to regenerate both; "
         "do not edit either by hand.")


def test_dashboard_embed_matches_committed_journal_csv():
    """dashboard.html's MY_JOURNAL block must be exactly what journal.csv produces."""
    reader = csv.DictReader(open(REPO_ROOT / 'journal.csv', newline=''))
    missing = [c for c in EMBED_COLUMNS if c not in reader.fieldnames]
    assert not missing, f"journal.csv has no {', '.join(missing)} column(s). " + RERUN
    expected = 'const MY_JOURNAL = ' + build_embed(list(reader)) + ';'
    html = (REPO_ROOT / 'dashboard.html').read_text(encoding='utf-8')
    blocks = EMBED_PATTERN.findall(html)
    assert len(blocks) == 1, f"expected 1 MY_JOURNAL block in dashboard.html, found {len(blocks)}"
    assert blocks[0] == expected, "dashboard.html's embedded data does not match journal.csv. " + RERUN


def test_net_r_rounds_half_up_deterministically():
    """-6.665 must round to -6.67 on every interpreter, not bankers'-round to -6.66."""
    trades = [
        {'r': -3.3325, 'result': 'Loss', 'date': '2026-01-01'},
        {'r': -3.3325, 'result': 'Loss', 'date': '2026-01-02'},
    ]
    assert summarise(trades)['net_r'] == -6.67


def test_win_rate_rounds_half_up_like_sql_and_the_dashboard():
    """21 wins in 80 decided trades is 26.25%: half-up gives 26.3, round() would give 26.2."""
    trades = [{'r': 1, 'result': 'Win'}] * 21 + [{'r': -1, 'result': 'Loss'}] * 59
    assert summarise(trades)['win_rate'] == 26.3
