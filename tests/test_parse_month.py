from scripts.fill_planned_indicators import parse_month


def test_parse_month_various_formats():
    assert parse_month("01.2024") == 1
    assert parse_month("2024-03") == 3
    assert parse_month(1.0) == 1
    assert parse_month("abc") == 0

