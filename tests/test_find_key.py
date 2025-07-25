from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import find_key

def test_find_key_punctuation():
    idx = {
        'себестоимостьналог': 0,
        'себестоимостьпродажналог, ₽': 1,
    }
    assert find_key(idx, 'СебестоимостьНалог') == 'себестоимостьналог'
    assert find_key(idx, 'СебестоимостьПродажНалог') == 'себестоимостьпродажналог, ₽'

