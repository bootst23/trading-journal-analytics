#!/usr/bin/env python3
"""
update_data.py — refresh the trading dashboard from the journal workbook.

Usage:
    python update_data.py Trading_Months_Stats.xlsx

What it does:
    1. Reads the 'Master' sheet from the workbook
    2. Extracts one row per trade (date, direction, result, R, violations, type, grade)
    3. Validates the data:
         - flags trades whose date cell disagrees with their Month label
         - flags missing/invalid R values or results
    4. Writes a clean journal.csv
    5. Rewrites the embedded data block inside dashboard.html

Requires: pip install openpyxl
"""
import sys, re, csv, datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

MONTHS = {'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
          'july':7,'august':8,'september':9,'october':10,'november':11,'december':12}

def extract(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if 'Master' not in wb.sheetnames:
        sys.exit("No 'Master' sheet found in the workbook.")
    ws = wb['Master']
    trades, warnings = [], []

    for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
        month, direction, date, no, result, r, v1, v2, v3, v4, v5, ttype, grade, *_ = row
        if not isinstance(date, datetime.datetime) or result not in ('Win', 'Loss', 'BE'):
            continue
        if not isinstance(r, (int, float)):
            warnings.append(f"  ! {date.date()} trade #{no}: R value '{r}' is not numeric — skipped")
            continue
        # consistency check: does the date's month match the Month label?
        label_m = MONTHS.get(str(month).strip().lower())
        if label_m and abs(date.month - label_m) not in (0,):
            # allow a boundary trade (e.g. opened on the last day of the prior month)
            boundary = (date + datetime.timedelta(days=1)).month == label_m
            tag = "boundary?" if boundary else "CHECK — possible typo"
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

def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    xlsx = Path(sys.argv[1])
    dash = Path(__file__).parent / 'dashboard.html'
    if not xlsx.exists(): sys.exit(f"Workbook not found: {xlsx}")
    if not dash.exists(): sys.exit(f"dashboard.html not found next to this script")

    trades, warnings = extract(xlsx)
    if not trades: sys.exit("No valid trades found in the Master sheet.")

    # write CSV
    csv_path = Path(__file__).parent / 'journal.csv'
    with open(csv_path, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=trades[0].keys())
        w.writeheader(); w.writerows(trades)

    # rebuild embed block
    lines = ['date,direction,result,r,violations,type,grade']
    for t in trades:
        lines.append(f"{t['date']},{t['direction']},{t['result']},{t['r']},"
                     f"{t['violations']},{t['type']},{t['grade']}")
    embed = '\n'.join(lines)

    html = dash.read_text()
    html, n = re.subn(r'const MY_JOURNAL = `[^`]*`;',
                      'const MY_JOURNAL = `' + embed + '`;', html, flags=re.S)
    if n != 1: sys.exit("Could not locate the embedded data block in dashboard.html")
    dash.write_text(html)

    # summary
    net = sum(t['r'] for t in trades)
    wins = sum(t['result'] == 'Win' for t in trades)
    losses = sum(t['result'] == 'Loss' for t in trades)
    print(f"Embedded {len(trades)} trades into dashboard.html")
    print(f"  net: {net:+.2f}R | win rate (excl BE): {100*wins/(wins+losses):.1f}% "
          f"| span: {trades[0]['date']} → {trades[-1]['date']}")
    if warnings:
        print("\nData-quality warnings (fix these in the workbook if wrong):")
        print('\n'.join(warnings))

if __name__ == '__main__':
    main()
