import pytest
from scripts.fill_planned_indicators import _calc_cost_base, full_cogs


def _compute_costs(rows, nds):
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
    cost_base = _calc_cost_base(g.get('cn'), g['cr'], nds)
    ct_val = g.get('ct')
    cost_tax = ct_val if ct_val is not None else full_cogs(cost_base, nds)
    ct_wo_val = g.get('ct_wo')
    cost_tax_wo = ct_wo_val if ct_wo_val is not None else cost_base
    return cost_tax, cost_tax_wo


@pytest.mark.parametrize(
    "row",
    [
        {"org": "Org", "month": 1, "rev": 0, "mp": 0, "cr": 110, "cn": 100},
        {"org": "Org", "month": 1, "rev": 0, "mp": 0, "cr": 110, "cn": 100, "ct": None, "ct_wo": None},
    ],
)
def test_cost_tax_fallback(row):
    ct, ct_wo = _compute_costs([row], nds=10)
    assert ct == pytest.approx(110)
    assert ct_wo == pytest.approx(100)
