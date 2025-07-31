from scripts.fill_planned_indicators import full_cogs

def test_full_cogs_reduced_rates():
    assert full_cogs(100, 5) == 105
    assert full_cogs(200, 7) == 214
