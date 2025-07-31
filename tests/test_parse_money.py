from scripts.fill_planned_indicators import parse_money


def test_parse_money_comma():
    assert parse_money("1,5") == 1.5
