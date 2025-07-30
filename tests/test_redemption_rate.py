import pandas as pd
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from update_monthly_scenario_calc import build_redemption_rate


def test_nmId_mapping_logistics():
    wb_table = pd.DataFrame({
        'nmId': ['n1'],
        '% выкупа': [80]
    })
    nm_to_wb = {'n1': 'WB123'}
    red = build_redemption_rate(wb_table, nm_to_wb)
    assert red == {'WB123': 80.0}

    REVERSE_LOG = 50
    per_unit = 100
    wb_percent = red.get('WB123', 95)
    return_rate = 1 - wb_percent / 100
    per_full = per_unit + REVERSE_LOG * return_rate
    assert round(per_full) == 110
