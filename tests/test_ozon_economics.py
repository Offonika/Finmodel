import pandas as pd
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "scripts"))
from economics_table import compute_ozon_economics_df

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
        'Себестоимость_Налог, руб (новый)': [300],
    })
    df = compute_ozon_economics_df(plan_df, cost_df, {})
    assert df.loc[0, 'СебестоимостьНалог_руб'] == 300
