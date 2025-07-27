from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))

def calc_corp_tax(rows):
    cum = 0
    taxes = []
    for r in rows:
        if r.get('prevM') != 'ОСНО':
            cum = 0
        prev = cum
        cum += r['ebit_tax']
        tax_prev = max(0, prev * 0.25)
        tax_now = max(0, cum * 0.25)
        taxes.append(max(0, round(tax_now - tax_prev)))
    return taxes


def test_corporate_osno_delta():
    rows = [
        {'ebit_tax': -500, 'prevM': 'ОСНО'},
        {'ebit_tax': 1000, 'prevM': 'ОСНО'},
    ]
    taxes = calc_corp_tax(rows)
    assert taxes[0] == 0
    assert taxes[1] == 125
