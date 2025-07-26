from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import ndfl_prog


def calc_osno_tax(rows, consolidate=False):
    cum = {}
    results = []
    for r in rows:
        key = 'consolidated' if consolidate else r['org']
        if r.get('prevM') != 'ОСНО' and key not in cum:
            cum[key] = 0
        prev = cum.get(key, 0)
        base = r['ebit_tax']
        total = prev + base
        taxable_prev = max(prev, 0)
        taxable_total = max(total, 0)
        tax = max(0, round(ndfl_prog(taxable_total) - ndfl_prog(taxable_prev)))
        cum[key] = total
        results.append({'org': r['org'], 'tax': tax,
                        'base': total,
                        'cons_base': cum.get('consolidated') if consolidate else None})
    return results


def test_consolidated_base_shared():
    rows = [
        {'org': 'A', 'ebit_tax': 100, 'prevM': 'ОСНО'},
        {'org': 'B', 'ebit_tax': 200, 'prevM': 'ОСНО'},
    ]
    res = calc_osno_tax(rows, consolidate=True)
    assert res[0]['tax'] == 13
    assert res[1]['tax'] == 26
    assert res[0]['cons_base'] == 100
    assert res[1]['cons_base'] == 300


def test_individual_base_when_not_consolidated():
    rows = [
        {'org': 'A', 'ebit_tax': 100, 'prevM': 'ОСНО'},
        {'org': 'B', 'ebit_tax': 200, 'prevM': 'ОСНО'},
    ]
    res = calc_osno_tax(rows, consolidate=False)
    assert res[0]['base'] == 100
    assert res[1]['base'] == 200
    assert res[0]['cons_base'] is None
