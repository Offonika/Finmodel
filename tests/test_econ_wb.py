import pandas as pd
from scripts.economics_table import compute_wb_economics_df


def test_wb_tax_columns():
    plan_df = pd.DataFrame({
        'Организация': ['Org'],
        'Артикул_WB': ['W1'],
        'Артикул_поставщика': ['A1'],
        'Предмет': ['Cat'],
        'Комиссия WB %': [0.1],
        'Выручка, ₽': [1000],
        'Мес.01': [1],
    })
    cost_df = pd.DataFrame({
        'Организация': ['Org'],
        'Артикул_поставщика': ['A1'],
        'Себестоимость_руб': [100],
        'Себестоимость_без_НДС_руб': [80],
        'СебестоимостьНалог': [70],
        'СебестоимостьНалог_без_НДС': [60],
    })
    df = compute_wb_economics_df(plan_df, cost_df)
    assert 'СебестоимостьПродажНалог, ₽' in df.columns
    assert 'СебестоимостьПродажНалог_без_НДС, ₽' in df.columns
    assert df.loc[0, 'СебестоимостьПродажНалог, ₽'] == 70
    assert df.loc[0, 'СебестоимостьПродажНалог_без_НДС, ₽'] == 60
