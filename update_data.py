#!/usr/bin/env python3
"""
update_data.py - refresh the trading dashboard from the journal workbook.

Usage:
    python update_data.py Trading_Months_Stats.xlsx
    python update_data.py Trading_Months_Stats.xlsx --db      # also build journal.db
    python update_data.py --db-only      # build journal.db from the committed journal.csv

What it does:
    1. Reads the 'Master' sheet from the workbook
    2. Extracts one row per trade (date, direction, result, R, the eight violation
       flags, type, setup and execution grade, minutes in trade and in drawdown)
    3. Validates the data:
         - flags trades whose date cell disagrees with their Month label
         - flags missing/invalid R values or results
         - flags durations it cannot read (the field is left blank)
    4. Writes a clean journal.csv
    5. Rewrites the embedded data block inside dashboard.html
    6. Optionally loads journal.csv into a typed SQLite database (journal.db)

Requires: pip install openpyxl
"""
import sys, re, csv, json, sqlite3, datetime
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

MONTHS = {'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
          'july':7,'august':8,'september':9,'october':10,'november':11,'december':12}

# The CSV column order. The SQLite schema below must stay in the same order,
# because sqlite's .import (and our INSERT) are positional.
FIELDNAMES = ['date', 'direction', 'trade_no', 'result', 'r',
              'v1', 'v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8', 'violations',
              'type', 'grade', 'exec_grade', 'mins_in_trade', 'mins_in_drawdown']

SCHEMA = """CREATE TABLE trades (
    date TEXT, direction TEXT, trade_no INTEGER, result TEXT, r REAL,
    v1 INTEGER, v2 INTEGER, v3 INTEGER, v4 INTEGER,
    v5 INTEGER, v6 INTEGER, v7 INTEGER, v8 INTEGER, violations INTEGER,
    type TEXT, grade TEXT, exec_grade TEXT, mins_in_trade INTEGER, mins_in_drawdown INTEGER)"""

# The columns the dashboard actually consumes.
EMBED_COLUMNS = ['date', 'trade_no', 'direction', 'result', 'r', 'violations', 'type', 'grade',
                 'exec_grade', 'mins_in_trade', 'mins_in_drawdown']

# One JSON row per line between '[' and '];'. json.dumps escapes newlines, so a
# line holding only '];' can only be the end of the block.
EMBED_PATTERN = re.compile(r'^const MY_JOURNAL = \[$.*?^\];$', re.M | re.S)

# Master sheet headers extract() needs, matched by text (trimmed), not position.
# 'D. helper' is deliberately not read: it is a second copy of the drawdown
# minutes and has drifted out of line with 'D. in drawdown'.
V_HEADERS = [f'V{i}' for i in range(1, 9)]
REQUIRED_HEADERS = ['Month', 'Direction', 'Date', 'Trade No.', 'Result', 'R', *V_HEADERS,
                    'Type', 'Setup Grade', 'Execution Grade', 'D. in trade', 'D. in drawdown']
DURATION_FIELDS = {'mins_in_trade': 'D. in trade', 'mins_in_drawdown': 'D. in drawdown'}

# '1d 10h 59m', '3h', '15m', '18h29m': each part optional, spaces optional.
DURATION_PATTERN = re.compile(r'(?:(\d+)d)?\s*(?:(\d+)h)?\s*(?:(\d+)m)?')


class PipelineError(Exception):
    """Raised when the pipeline cannot safely continue.

    Library functions raise this; only main() decides to exit the process.
    """


# --------------------------------------------------------------------------
# Pure functions - no file I/O, no sys.exit. These are the tested core.
# --------------------------------------------------------------------------

def parse_duration(value):
    """Turn '1d 10h 59m', '3h', '15m', '18h29m', '0', '0m' or 0 into whole minutes.

    Raises ValueError for anything else (blank, unreadable text, other numbers).
    """
    if isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0:
        return 0
    text = value.strip() if isinstance(value, str) else ''
    if text == '0':
        return 0
    match = DURATION_PATTERN.fullmatch(text)
    if not match or not any(match.groups()):
        raise ValueError(f"unreadable duration: {value!r}")
    days, hours, mins = (int(g or 0) for g in match.groups())
    return days * 24 * 60 + hours * 60 + mins


def build_embed(trades):
    """Turn trade dicts into the JS array literal the dashboard embeds, one row per line.

    JSON is valid JS, so commas and quotes in any field are safe. Values are strings,
    as they were in the CSV embed, and a blank cell (None) is '' as csv.writer writes it,
    so a workbook extract and journal.csv embed alike.
    '<' is escaped so a '</script>' in a field cannot end the script block.
    """
    rows = (json.dumps({c: '' if t[c] is None else str(t[c]) for c in EMBED_COLUMNS})
            .replace('<', '\\u003c')
            for t in trades)
    return '[\n' + ',\n'.join(rows) + '\n]'


def replace_embed(html, embed):
    """Swap the MY_JOURNAL block in `html` for `embed`.

    Raises PipelineError unless exactly one block was found, so a renamed or
    duplicated anchor can never silently corrupt the dashboard.
    """
    new_html, n = EMBED_PATTERN.subn(lambda _: 'const MY_JOURNAL = ' + embed + ';', html)
    if n != 1:
        raise PipelineError(
            f"Expected exactly 1 MY_JOURNAL block in dashboard.html, found {n}.")
    return new_html


def summarise(trades):
    """Headline figures, computed once and reused by the printout and the tests."""
    net = sum(Decimal(str(t['r'])) for t in trades)
    wins = sum(t['result'] == 'Win' for t in trades)
    losses = sum(t['result'] == 'Loss' for t in trades)
    decided = wins + losses
    return {
        'trades': len(trades),
        'net_r': float(net.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
        'win_rate': float((Decimal(100 * wins) / decided).quantize(
            Decimal('0.1'), rounding=ROUND_HALF_UP)) if decided else 0.0,
        'first_date': trades[0].get('date') if trades else None,
        'last_date': trades[-1].get('date') if trades else None,
    }


def summary_text(s):
    """The run summary main() prints. ASCII only: a cp1252 console cannot encode more."""
    return (f"Embedded {s['trades']} trades into dashboard.html\n"
            f"  net: {s['net_r']:+.2f}R | win rate (excl BE): {s['win_rate']:.1f}% "
            f"| span: {s['first_date']} -> {s['last_date']}")


# --------------------------------------------------------------------------
# I/O functions
# --------------------------------------------------------------------------

def extract(xlsx_path):
    """Read the Master sheet into normalised trade dicts + data-quality warnings."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if 'Master' not in wb.sheetnames:
        raise PipelineError("No 'Master' sheet found in the workbook.")
    ws = wb['Master']
    header = next(ws.iter_rows(max_row=1, values_only=True), ())
    col = {str(h).strip(): i for i, h in enumerate(header) if h is not None}
    missing = [h for h in REQUIRED_HEADERS if h not in col]
    if missing:
        raise PipelineError(f"Master sheet is missing required header(s): {', '.join(missing)}")
    trades, warnings = [], []

    for row in ws.iter_rows(min_row=2, values_only=True):
        month, direction, date, no, result, r = (row[col[h]] for h in REQUIRED_HEADERS[:6])
        ttype, grade = row[col['Type']], row[col['Setup Grade']]
        exec_grade = row[col['Execution Grade']]
        if not isinstance(date, datetime.datetime) or result not in ('Win', 'Loss', 'BE'):
            continue
        if not isinstance(r, (int, float)):
            warnings.append(f"  ! {date.date()} trade #{no}: R value '{r}' is not numeric - skipped")
            continue
        label_m = MONTHS.get(str(month).strip().lower())
        if label_m and date.month != label_m:
            boundary = (date + datetime.timedelta(days=1)).month == label_m
            tag = "boundary?" if boundary else "CHECK - possible typo"
            warnings.append(f"  ! {date.date()} trade #{no}: dated {date.strftime('%B')} "
                            f"but labelled '{month}' ({tag})")
        # SQLite has no boolean type; 1/0 can be added up.
        flags = {f'v{i}': int(bool(row[col[f'V{i}']])) for i in range(1, 9)}
        minutes = {}
        for field, header in DURATION_FIELDS.items():
            try:
                minutes[field] = parse_duration(row[col[header]])
            except ValueError:
                minutes[field] = ''
                warnings.append(f"  ! {date.date()} trade #{no}: '{header}' value "
                                f"{row[col[header]]!r} is unreadable - left blank")
        trades.append({
            'date': date.strftime('%Y-%m-%d'), 'direction': direction,
            'trade_no': no, 'result': result, 'r': r,
            **flags, 'violations': sum(flags.values()),
            'type': ttype, 'grade': grade,
            'exec_grade': '' if exec_grade is None else str(exec_grade).strip(),
            **minutes,
        })

    trades.sort(key=lambda t: (t['date'], int(t['trade_no'])))
    return trades, warnings


def write_csv(trades, csv_path):
    """Write the full extract, one column per FIELDNAMES entry."""
    with open(csv_path, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(trades)
    return csv_path


def load_database(csv_path, db_path):
    """Build a typed SQLite database from journal.csv.

    The table is created explicitly before loading. Letting sqlite infer the
    schema types every column TEXT, which silently breaks numeric comparisons
    like `WHERE r < 0`. Blank cells, and columns an older CSV does not have,
    are stored as NULL so AVG() skips them instead of counting them as 0.
    """
    Path(db_path).unlink(missing_ok=True)
    con = sqlite3.connect(db_path)
    try:
        con.execute(SCHEMA)
        with open(csv_path, newline='') as f:
            rows = [[row.get(c) or None for c in FIELDNAMES] for row in csv.DictReader(f)]
        placeholders = ','.join('?' * len(FIELDNAMES))
        con.executemany(f"INSERT INTO trades VALUES ({placeholders})", rows)
        con.commit()
        return con.execute("SELECT COUNT(*) FROM trades").fetchone()[0]
    finally:
        con.close()


# --------------------------------------------------------------------------

def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        sys.exit(__doc__)
    here = Path(__file__).parent
    if argv[0] == '--db-only':
        csv_path = here / 'journal.csv'
        if not csv_path.exists():
            sys.exit("journal.csv not found next to this script")
        n = load_database(csv_path, here / 'journal.db')
        print(f"Loaded {n} trades into journal.db")
        return
    xlsx = Path(argv[0])
    want_db = '--db' in argv[1:]
    dash = here / 'dashboard.html'
    if not xlsx.exists():
        sys.exit(f"Workbook not found: {xlsx}")
    if not dash.exists():
        sys.exit("dashboard.html not found next to this script")

    try:
        trades, warnings = extract(xlsx)
        if not trades:
            raise PipelineError("No valid trades found in the Master sheet.")
        write_csv(trades, here / 'journal.csv')
        # Explicit, not the locale default (cp1252 on Windows, which cannot decode some UTF-8
        # bytes). The read keeps universal newlines, which EMBED_PATTERN relies on;
        # read_text only takes newline= from Python 3.13.
        html = dash.read_text(encoding='utf-8')
        dash.write_text(replace_embed(html, build_embed(trades)), encoding='utf-8', newline='\n')
        if want_db:
            n = load_database(here / 'journal.csv', here / 'journal.db')
            print(f"Loaded {n} trades into journal.db")
    except PipelineError as e:
        sys.exit(f"Pipeline aborted: {e}")

    print(summary_text(summarise(trades)))
    if warnings:
        print("\nData-quality warnings (fix these in the workbook if wrong):")
        print('\n'.join(warnings))


if __name__ == '__main__':
    main()
