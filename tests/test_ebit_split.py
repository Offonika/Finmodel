from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import _calc_row


def test_mp_excluded_from_tax():
    row = _calc_row(revN=1000, mpNet=200, cost=300, fot=0, esn=0, oth=0, mode='Доходы-Расходы')
    assert row['EBITDA, ₽'] == 500
    assert row['EBITDA нал., ₽'] == 700
