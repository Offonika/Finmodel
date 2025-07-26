from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import _apply_consolidated_dr_tax, calc_consolidated_min_tax


def test_consolidated_usn_min_tax_distribution():
    rows = [
        {'m': 1, 'revN': 1000, 'ebit_tax': 100, 'usn': 6},
        {'m': 1, 'revN': 2000, 'ebit_tax': -200, 'usn': 6},
    ]
    totals = _apply_consolidated_dr_tax(rows)
    assert totals[1] == -100
    assert rows[0]['tax'] == 10
    assert rows[1]['tax'] == 20
    assert rows[0]['usn_forced_min']
    assert rows[1]['usn_forced_min']


def test_calc_consolidated_min_tax():
    tax = calc_consolidated_min_tax(100, 5000, 0.15)
    assert tax == 50
    tax = calc_consolidated_min_tax(1000, 2000, 0.06)
    assert tax == 60
