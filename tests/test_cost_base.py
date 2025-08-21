import pytest
from scripts.fill_planned_indicators import _calc_cost_base


def test_cost_base_uses_cn_when_available():
    assert _calc_cost_base(120, 144, 20) == 120


def test_cost_base_recalculates_when_cn_incorrect():
    cn = 120
    cr = 160
    nds = 20
    expected = cr / (1 + nds / 100)
    assert _calc_cost_base(cn, cr, nds) == pytest.approx(expected)


@pytest.mark.parametrize("cr, nds, expected", [
    (120, 20, 100),
    (105, 5, 100),
    (107, 7, 100),
])
def test_cost_base_fallback(cr, nds, expected):
    assert _calc_cost_base(None, cr, nds) == pytest.approx(expected)
