import pandas as pd
from scripts.economics_table import compute_ozon_economics_df

def test_taxable_cogs_column():
    plan_df = pd.DataFrame({
        'Организация': ['Org'],
        'Артикул_поставщика': ['A1'],
        'SKU': ['S1'],
        'Плановая цена': [500],
        'Мес.01': [1],
    })
    cost_df = pd.DataFrame({
        'Организация': ['Org'],
        'Артикул_поставщика': ['A1'],
        'Себестоимость_руб': [100],
        'Себестоимость_без_НДС_руб': [80],
        'СебестоимостьУпр': [100],
        'СебестоимостьНалог_руб': [300],
    })
    df = compute_ozon_economics_df(plan_df, cost_df, {})
    assert 'СебестоимостьПродажНалог, ₽' in df.columns
    assert 'СебестоимостьПродажНалог_без_НДС, ₽' in df.columns
    assert df.loc[0, 'СебестоимостьПродажНалог, ₽'] == 300
