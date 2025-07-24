import pandas as pd

TAX_DEDUCTIBLE_BY_LOGISTIC = {'Карго': False, 'Белая': True}

def compute_taxable_cogs(purchase_rub, logistics_rub, duty_rub, vat_rub, mode):
    total = purchase_rub + logistics_rub + duty_rub + vat_rub
    return total if TAX_DEDUCTIBLE_BY_LOGISTIC.get(mode, True) else 0

def test_cargo_mode_zero():
    assert compute_taxable_cogs(100, 20, 5, 24, 'Карго') == 0

def test_white_mode_full():
    assert compute_taxable_cogs(100, 20, 5, 24, 'Белая') == 149
