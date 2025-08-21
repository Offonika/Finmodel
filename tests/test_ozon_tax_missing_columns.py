import pytest
from scripts.fill_planned_indicators import (
    parse_money,
    parse_month,
    _calc_cost_base,
    full_cogs,
)


def test_ozon_missing_tax_columns_triggers_fallback():
    oz_idx = {
        'организация': 0,
        'месяц': 1,
        'выручка_руб': 2,
        'итогорасходымп_руб': 3,
        'себестоимостьпродаж_руб': 4,
        'себестоимостьбезндс_руб': 5,
    }
    oz_rows = [['Org', '01', '0', '0', '110', '100']]

    rows = []
    tax_col_oz = None
    tax_nds_col_oz = None
    for r in oz_rows:
        rows.append(
            dict(
                org=r[oz_idx['организация']],
                month=parse_month(r[oz_idx['месяц']]),
                rev=parse_money(r[oz_idx['выручка_руб']]) or 0,
                mp=parse_money(r[oz_idx['итогорасходымп_руб']]) or 0,
                cr=parse_money(r[oz_idx['себестоимостьпродаж_руб']]) or 0,
                cn=parse_money(r[oz_idx['себестоимостьбезндс_руб']]),
                ct=parse_money(r[tax_col_oz]) if tax_col_oz is not None else None,
                ct_wo=parse_money(r[tax_nds_col_oz]) if tax_nds_col_oz is not None else None,
            )
        )

    assert rows[0]['ct'] is None
    assert rows[0]['ct_wo'] is None

    grouped = {}
    for r in rows:
        k = (r['org'], r['month'])
        g = grouped.setdefault(k, dict(org=r['org'], month=r['month'], rev=0, mp=0, cr=0, cn=0))
        for f in ('rev', 'mp', 'cr', 'cn'):
            g[f] += r.get(f, 0)
        for f in ('ct', 'ct_wo'):
            val = r.get(f)
            if val is not None:
                g[f] = g.get(f, 0) + val

    g = list(grouped.values())[0]
    nds = 10
    cost_base = _calc_cost_base(g.get('cn'), g['cr'], nds)
    ct_val = g.get('ct')
    cost_tax = ct_val if ct_val is not None else full_cogs(cost_base, nds)
    ct_wo_val = g.get('ct_wo')
    cost_tax_wo = ct_wo_val if ct_wo_val is not None else cost_base

    assert cost_tax == pytest.approx(110)
    assert cost_tax_wo == pytest.approx(100)
