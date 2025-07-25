from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import _calc_row


def test_mp_with_vat_for_dr_mode():
    row = _calc_row(
        revN=1000,
        mp_mgmt=200,
        mp_tax=240,
        cost_sales=300,
        cost_tax=360,
        fot=0,
        esn=0,
        oth=0,
        mode='Доходы-Расходы'
    )
    assert row['EBITDA, ₽'] == 500
    assert row['EBITDA нал., ₽'] == 400
