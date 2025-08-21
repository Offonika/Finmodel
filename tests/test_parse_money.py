import pytest

from scripts.fill_planned_indicators import _calc_cost_base, parse_money


def test_parse_money_comma():
    assert parse_money("1,5") == 1.5


def test_parse_money_empty_returns_none():
    assert parse_money("") is None
    assert parse_money(None) is None


def test_calc_cost_base_uses_cr_when_cn_missing():
    cn = parse_money("")
    assert cn is None
    assert _calc_cost_base(cn, 120, 20) == pytest.approx(100)
