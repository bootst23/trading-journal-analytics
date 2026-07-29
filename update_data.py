#!/usr/bin/env python3
"""
update_data.py - refresh the trading dashboard from the journal workbook.

Usage:
    python update_data.py Trading_Months_Stats.xlsx
    python update_data.py Trading_Months_Stats.xlsx --db      # also build journal.db

What it does:
    1. Reads the 'Master' sheet from the workbook
    2. Extracts one row per trade (date, direction, result, R, violations, type, grade)
    3. Validates the data:
         - flags trades whose date cell disagrees with their Month label
         - flags missing/invalid R values or results
    4. Writes a clean journal.csv
    5. Rewrites the embedded data block inside dashboard.html
    6. Optionally loads journal.csv into a typed SQLite database (journal.db)

Requires: pip install openpyxl
"""
import sys, re, csv, sqlite3, datetime
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

MONTHS = {'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
          'july':7,'august':8,'september':9,'october':10,'november':11,'december':12}

# The CSV column order. The SQLite schema below must stay in the same order,
# because sqlite's .import (and our INSERT) are positional.
FIELDNAMES = ['date', 'month', 'direction', 'trade_no', 'result', 'r',
              'v1', 'v2', 'v3', 'v4', 'violations', 'type', 'grade', 'mins_in_trade']

SCHEMA = """CREATE TABLE trades (
    date TEXT, month TEXT, direction TEXT, trade_no INTEGER,
    result TEXT, r REAL, v1 TEXT, v2 TEXT, v3 TEXT, v4 TEXT,
    violations INTEGER, type TEXT, grade TEXT, mins_in_trade REAL)"""

# The 7 columns the dashboard actually consumes.
EMBED_COLUMNS = ['date', 'direction', 'result', 'r', 'violations', 'type', 'grade']

EMBED_PATTERN = r'const MY_JOURNAL = `[^`]*`;'


class PipelineError(Exception):
    """Raised when the pipeline cannot safely continue.

    Library functions raise this; only main() decides to exit the process.
    """


# --------------------------------------------------------------------------
# Pure functions - no file I/O, no sys.exit. These are the tested core.
# --------------------------------------------------------------------------

def build_embed(trades):
    """Turn trade dicts into the 7-column CSV text the dashboard embeds."""
    lines = [','.join(EMBED_COLUMNS)]
    for t in trades:
        lines.append(','.join(str(t[c]) for c in EMBED_COLUMNS))
    return '\n'.join(lines)


def replace_embed(html, embed):
    """Swap the MY_JOURNAL block in `html` for `embed`.

    Raises PipelineError unless exactly one block was found, so a renamed or
    duplicated anchor can never silently corrupt the dashboard.
    """
    if '`' in embed:
        raise PipelineError("Embed data contains a backtick - it would break the JS template literal.")
    new_html, n = re.subn(EMBED_PATTERN,
                          lambda _: 'const MY_JOURNAL = `' + embed + '`;',
                          html)
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
        'win_rate': round(100 * wins / decided, 1) if decided else 0.0,
        'first_date': trades[0].get('date') if trades else None,
        'last_date': trades[-1].get('date') if trades else None,
    }


# --------------------------------------------------------------------------
# I/O functions
# --------------------------------------------------------------------------

def extract(xlsx_path):
    """Read the Master sheet into normalised trade dicts + data-quality warnings."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if 'Master' not in wb.sheetnames:
        raise PipelineError("No 'Master' sheet found in the workbook.")
    ws = wb['Master']
    trades, warnings = [], []

    for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
        month, direction, date, no, result, r, v1, v2, v3, v4, v5, ttype, grade, *_ = row
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
        trades.append({
            'date': date.strftime('%Y-%m-%d'), 'month': month, 'direction': direction,
            'trade_no': no, 'result': result, 'r': r,
            'v1': bool(v1), 'v2': bool(v2), 'v3': bool(v3), 'v4': bool(v4),
            'violations': sum(bool(x) for x in (v1, v2, v3, v4, v5)),
            'type': ttype, 'grade': grade, 'mins_in_trade': '',
        })

    trades.sort(key=lambda t: (t['date'], int(t['trade_no'])))
    return trades, warnings


def write_csv(trades, csv_path):
    """Write the full 14-column extract."""
    with open(csv_path, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(trades)
    return csv_path


def load_database(csv_path, db_path):
    """Build a typed SQLite database from journal.csv.

    The table is created explicitly before loading. Letting sqlite infer the
    schema types every column TEXT, which silently breaks numeric comparisons
    like `WHERE r < 0`.
    """
    db_path = Path(db_path)
    if db_path.exists():
        db_path.unlink()
    con = sqlite3.connect(db_path)
    try:
        con.execute(SCHEMA)
        with open(csv_path, newline='') as f:
            rows = [[row[c] for c in FIELDNAMES] for row in csv.DictReader(f)]
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
    xlsx = Path(argv[0])
    want_db = '--db' in argv[1:]
    here = Path(__file__).parent
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
        dash.write_text(replace_embed(dash.read_text(), build_embed(trades)))
        if want_db:
            n = load_database(here / 'journal.csv', here / 'journal.db')
            print(f"Loaded {n} trades into journal.db")
    except PipelineError as e:
        sys.exit(f"Pipeline aborted: {e}")

    s = summarise(trades)
    print(f"Embedded {s['trades']} trades into dashboard.html")
    print(f"  net: {s['net_r']:+.2f}R | win rate (excl BE): {s['win_rate']:.1f}% "
          f"| span: {s['first_date']} → {s['last_date']}")
    if warnings:
        print("\nData-quality warnings (fix these in the workbook if wrong):")
        print('\n'.join(warnings))


if __name__ == '__main__':
    main()
