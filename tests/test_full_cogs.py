import pytest
from scripts.fill_planned_indicators import full_cogs


def test_full_cogs_vat_rates():
    assert full_cogs(100, 0) == pytest.approx(120)
    assert full_cogs(100, 5) == pytest.approx(115)
    assert full_cogs(200, 7) == pytest.approx(226)
    assert full_cogs(300, 20) == 300
