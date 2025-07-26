from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import ndfl_prog, consolidate_osno_tax


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


def test_tax_shown_only_once_in_consolidation():
    rows_out = []
    row_meta = []

    def make_row(org, ebit, tax):
        row = [0] * 30
        row[0] = org
        row[1] = 1
        row[18] = ebit
        row[26] = 'ОСНО'
        row[28] = tax
        row[29] = ebit - tax
        return row

    rows_out.append(make_row('A', 100, 13))
    rows_out.append(make_row('B', 200, 26))

    row_meta.append({'org': 'A', 'm': 1, 'mode': 'ОСНО', 'type': 'ИП', 'consolidation': True})
    row_meta.append({'org': 'B', 'm': 1, 'mode': 'ОСНО', 'type': 'ИП', 'consolidation': True})

    consolidate_osno_tax(rows_out, row_meta)

    assert rows_out[0][28] == 39
    assert rows_out[1][28] == 0
    assert rows_out[0][29] == 61
    assert rows_out[1][29] == 200
