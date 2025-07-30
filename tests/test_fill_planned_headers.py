from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / "scripts"))
from fill_planned_indicators import build_idx, find_key, parse_money


def test_ozon_new_header_mapping():
    header = [
        'Организация', 'Месяц', 'Выручка_руб', 'ИтогоРасходыМП_руб',
        'СебестоимостьПродаж_руб', 'СебестоимостьБезНДС_руб',
        'СебестоимостьПродажНалог, ₽', 'СебестоимостьПродажНалог_без_НДС, ₽'
    ]
    idx = build_idx(header)
    row = [
        'Org', 1, 1000, 50, 40, 30, 20, 10
    ]

    tax_col_candidates = ['СебестоимостьПродажНалог, ₽', 'СебестоимостьНалог_руб']
    tax_col_oz = None
    for cand in tax_col_candidates:
        key = find_key(idx, cand)
        if key is not None:
            tax_col_oz = idx[key]
            break

    tax_nds_col_oz = None
    key = find_key(idx, 'СебестоимостьПродажНалог_без_НДС, ₽')
    if key is not None:
        tax_nds_col_oz = idx[key]

    ct = parse_money(row[tax_col_oz]) if tax_col_oz is not None else 0
    ctn = parse_money(row[tax_nds_col_oz]) if tax_nds_col_oz is not None else 0

    assert ct == 20
    assert ctn == 10
