from scripts.fill_planned_indicators import find_key

def test_find_key_punctuation():
    idx = {
        'себестоимостьналог': 0,
        'себестоимостьпродажналог, ₽': 1,
    }
    assert find_key(idx, 'СебестоимостьНалог') == 'себестоимостьналог'
    assert find_key(idx, 'СебестоимостьПродажНалог') == 'себестоимостьпродажналог, ₽'

