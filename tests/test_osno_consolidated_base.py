from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import ndfl_prog, consolidate_osno_tax


def calc_osno_tax(rows, consolidate=False):
    cum = {}
    results = []
    for r in rows:
        key = 'consolidated' if consolidate else r['org']
        if r.get('prevM') != 'ОСНО':
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


def calc_osno_tax_new(rows, consolidate=False):
    cum = {}
    last_mode = {}
    results = []
    for r in rows:
        key = 'consolidated' if consolidate else r['org']
        if last_mode.get(key) != 'ОСНО' and r.get('mode', 'ОСНО') == 'ОСНО':
            cum[key] = 0
        prev = cum.get(key, 0)
        base = r['ebit_tax']
        total = prev + base
        taxable_prev = max(prev, 0)
        taxable_total = max(total, 0)
        tax = max(0, round(ndfl_prog(taxable_total) - ndfl_prog(taxable_prev)))
        cum[key] = total
        last_mode[key] = r.get('mode', 'ОСНО')
        results.append({'org': r['org'], 'tax': tax, 'base': total})
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
        # 19 column is База налога: in tests treat it
        # same as EBITDA for simplicity
        row[18] = ebit
        row[19] = ebit
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


def calc_osno_cum_cons(rows):
    month_totals = {}
    for r in rows:
        m = r['m']
        month_totals[m] = month_totals.get(m, 0) + r['ebit_tax']
    cum = {}
    run = 0
    for m in sorted(month_totals):
        run += month_totals[m]
        cum[m] = run
    return [cum[r['m']] for r in rows]


def test_consolidated_osno_cumulative_base_equal_per_month():
    rows = [
        {'org': 'A', 'm': 1, 'ebit_tax': 100},
        {'org': 'B', 'm': 1, 'ebit_tax': 200},
        {'org': 'A', 'm': 2, 'ebit_tax': 150},
        {'org': 'B', 'm': 2, 'ebit_tax': -50},
    ]

    cons_bases = calc_osno_cum_cons(rows)
    for month in {1, 2}:
        vals = [b for r, b in zip(rows, cons_bases) if r['m'] == month]
        assert len(set(vals)) == 1
        expected = 300 if month == 1 else 400
        assert vals[0] == expected


def test_yearly_tax_sum_matches_progressive():
    rows = [
        {'org': 'A', 'ebit_tax': 100, 'prevM': 'ОСНО'},
        {'org': 'A', 'ebit_tax': 200, 'prevM': 'ОСНО'},
        {'org': 'A', 'ebit_tax': 300, 'prevM': 'ОСНО'},
    ]

    res = calc_osno_tax(rows, consolidate=False)
    total_tax = sum(r['tax'] for r in res)
    total_base = sum(r['ebit_tax'] for r in rows)

    assert total_tax == round(ndfl_prog(total_base))


def test_reset_base_after_regime_change():
    rows = [
        {'org': 'A', 'ebit_tax': 2_200_000, 'prevM': 'Доходы'},
        {'org': 'A', 'ebit_tax': 300_000, 'prevM': 'ОСНО'},
        {'org': 'A', 'ebit_tax': 100_000, 'prevM': 'Доходы'},
    ]

    res = calc_osno_tax(rows, consolidate=False)

    assert res[0]['tax'] == round(ndfl_prog(2_200_000))
    assert res[1]['tax'] == round(ndfl_prog(2_500_000) - ndfl_prog(2_200_000))
    assert res[2]['tax'] == round(ndfl_prog(100_000))


def test_consolidated_transition_with_losses():
    rows = [
        {'org': 'A', 'ebit_tax': -200, 'mode': 'ОСНО'},
        {'org': 'B', 'ebit_tax': 100, 'mode': 'ОСНО'},
    ]

    res = calc_osno_tax_new(rows, consolidate=True)

    assert res[0]['tax'] == 0
    assert res[1]['tax'] == 0
    assert res[1]['base'] == -100
