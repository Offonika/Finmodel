import xlwings as xw
import requests
import sys
import os
import time
from dataclasses import dataclass, field
from typing import Any
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
print("==== PYTHONPATH ====")
print(sys.path)
print("==== WORKDIR ====")
print(os.getcwd())
print("==== FILES IN SCRIPTS ====")
print(os.listdir(os.path.dirname(__file__)))

SETTINGS_SHEET = 'НастройкиОрганизаций'
PRODUCTS_SHEET = 'Номенклатура_WB'
HEADERS = [
    'Организация', 'Артикул_WB', 'Артикул_поставщика',
    'Бренд', 'Название', 'Предмет',
    'Ширина', 'Высота', 'Длина', 'Вес_брутто', 'Объем_литр'
]
API_URL = 'https://content-api.wildberries.ru/content/v2/get/cards/list?locale=ru'
LIMIT = 100


def _session_with_retries():
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 501, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    return session


def get_idx(header_row):
    return {h.strip(): i for i, h in enumerate(header_row)}


@dataclass
class TokenCleanResult:
    token: str | None
    error: str | None
    details: dict[str, Any] = field(default_factory=dict)


def _mask_value(value: str) -> str:
    if not value:
        return ''
    if len(value) <= 2:
        return '*' * len(value)
    return f"{value[0]}{'*' * (len(value) - 2)}{value[-1]}"


def _describe_non_ascii(value: str):
    return [
        {
            'index': idx,
            'char': ch,
            'ord': ord(ch),
        }
        for idx, ch in enumerate(value)
        if not ch.isascii()
    ]


def _clean_token(token_raw):
    details: dict[str, Any] = {'original': token_raw}

    if token_raw is None:
        details['reason'] = 'missing'
        return TokenCleanResult(token=None, error='Token_WB отсутствует', details=details)

    raw_string = str(token_raw)
    details['raw'] = raw_string
    chars_to_remove = " \t\r\n\"'\u00a0\ufeff"
    translation_table = str.maketrans('', '', chars_to_remove)
    cleaned = raw_string.translate(translation_table).strip()
    details['cleaned'] = cleaned

    if cleaned.casefold() in {'', 'nan', 'none'}:
        details['reason'] = 'empty_after_cleaning'
        return TokenCleanResult(token=None, error='Token_WB пустой после очистки', details=details)

    if not cleaned.isascii():
        offending = _describe_non_ascii(cleaned)
        masked = _mask_value(cleaned)
        details.update({
            'reason': 'non_ascii',
            'masked_token': masked,
            'length': len(cleaned),
            'offending_characters': offending,
        })
        print('❌ Token_WB содержит недопустимые символы (не-ASCII).')
        print(f"    Маска токена: {masked} (длина: {len(cleaned)})")
        for item in offending:
            char_repr = repr(item['char'])
            print(f"    Недопустимый символ: {char_repr} (index={item['index']}, ord={item['ord']})")
        return TokenCleanResult(token=None, error='Token_WB содержит недопустимые символы', details=details)

    details.update({
        'masked_token': _mask_value(cleaned),
        'length': len(cleaned),
    })
    return TokenCleanResult(token=cleaned, error=None, details=details)


def _validate_headers_ascii(headers, *, context: str = 'HTTP headers'):
    for key, value in headers.items():
        for label, text in (('ключ', key), ('значение', value)):
            text_str = '' if text is None else str(text)
            if not text_str.isascii():
                masked = _mask_value(text_str)
                offending = _describe_non_ascii(text_str)
                target = 'ключ заголовка' if label == 'ключ' else 'значение заголовка'
                print(f'❌ {context}: {target} {key!r} содержит недопустимые символы.')
                print(f"    Маска строки: {masked} (длина: {len(text_str)})")
                for item in offending:
                    char_repr = repr(item['char'])
                    print(f"    Недопустимый символ: {char_repr} (index={item['index']}, ord={item['ord']})")
                return False
    return True

def main():
    print('=== START import_wb_product_cards ===')
    wb = xw.Book.caller()  # <-- ВАЖНО!
    sht_set = wb.sheets[SETTINGS_SHEET]

    # --- Подготовка листа с товарами ---
    sheet_names = [sht.name for sht in wb.sheets]
    if PRODUCTS_SHEET not in sheet_names:
        sht_prod = wb.sheets.add(PRODUCTS_SHEET, after=wb.sheets[wb.sheets.count-1])
        print(f'Создан новый лист: {PRODUCTS_SHEET}')
    else:
        sht_prod = wb.sheets[PRODUCTS_SHEET]
        print(f'Лист для загрузки карточек: {PRODUCTS_SHEET}')

    sht_prod.clear()
    sht_prod.range('A1').value = HEADERS

    hdr_rng = sht_prod.range((1, 1), (1, len(HEADERS)))
    hdr_rng.api.Font.Bold = True
    hdr_rng.api.HorizontalAlignment = -4108  # xlCenter
    hdr_rng.api.Borders.Weight = 2           # xlThin

    for col in range(1, len(HEADERS) + 1):
        sht_prod.range((1, col)).api.EntireColumn.AutoFit()
    print('Выполнен автоподбор ширины колонок.')

    cfgHdr = sht_set.range('A1').expand('right').value
    print('Шапка листа настроек:', cfgHdr)
    idx = get_idx(cfgHdr)
    if 'Организация' not in idx or 'Token_WB' not in idx:
        print('❌ В листе «НастройкиОрганизаций» нет колонок «Организация» и/или «Token_WB»')
        return

    org_col = idx['Организация']
    org_values = sht_set.range((2, org_col+1), (sht_set.cells.last_cell.row, org_col+1)).options(ndim=1).value
    org_rows_count = next((i for i, val in enumerate(org_values) if not val), len(org_values))
    if org_rows_count == 0:
        print('ℹ️ Нет организаций для обработки')
        return

    last_col = len(cfgHdr)
    settings = sht_set.range((2,1), (org_rows_count+1, last_col)).value

    allCards = []
    for i, row in enumerate(settings):
        org = row[idx['Организация']]
        token_raw = row[idx['Token_WB']]
        org_name = '' if org is None else str(org)
        print(f'--- Организация "{org_name}"')
        token_result = _clean_token(token_raw)
        if token_result.error:
            print(f'❌ Некорректный Token_WB у организации "{org_name}": {token_result.error}')
            continue
        token_clean = token_result.token
        if not org:
            print('Строка пропущена (нет org)')
            continue
        cursor = None
        page = 0
        existSet = set()
        session = None
        headers = {
            'Authorization': token_clean,
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'User-Agent': 'FinmodelWB/1.0',
        }
        if not _validate_headers_ascii(headers, context=f'HTTP-заголовки для организации "{org_name}"'):
            print('Строка пропущена из-за недопустимых символов в HTTP-заголовках')
            continue
        try:
            session = _session_with_retries()
            while True:
                page += 1
                payload = {
                    'settings': {
                        'cursor': cursor if cursor else {'limit': LIMIT},
                        'filter': {'withPhoto': -1}
                    }
                }
                try:
                    resp = session.post(API_URL, json=payload, headers=headers, timeout=60)
                except Exception as e:
                    print(f'❌ Сетевая ошибка: {e}, попытка {page}')
                    time.sleep(10)
                    continue

                print(f'HTTP {resp.status_code}')
                if resp.status_code != 200:
                    print(f'❌ API {resp.status_code}: {resp.text}')
                    break
                data = resp.json()
                cards = data.get('cards', [])
                print(f'Получено карточек: {len(cards)}')
                for c in cards:
                    nm = str(c.get('nmID', ''))
                    if nm and nm not in existSet:
                        width = c.get('dimensions', {}).get('width', '')
                        height = c.get('dimensions', {}).get('height', '')
                        length = c.get('dimensions', {}).get('length', '')
                        # Считаем объем, если все размеры есть и являются числами
                        try:
                            vol_ltr = float(width) * float(height) * float(length) / 1000
                            vol_ltr = round(vol_ltr, 3)
                        except Exception:
                            vol_ltr = ''
                        allCards.append([
                            org,
                            nm,
                            c.get('vendorCode', ''),
                            c.get('brand', ''),
                            c.get('title', ''),
                            c.get('subjectName', ''),
                            width, height, length,
                            c.get('dimensions', {}).get('weightBrutto', ''),
                            vol_ltr
                        ])
                        existSet.add(nm)

                cur = data.get('cursor', {})
                if cur.get('total') is None or cur.get('total', 0) < LIMIT:
                    print('Пагинация завершена')
                    break
                cursor = {k: cur[k] for k in ('updatedAt','nmID') if k in cur}
                cursor['limit'] = LIMIT
        finally:
            if session is not None:
                session.close()

    if allCards:
        sht_prod.range((2, 1)).value = allCards
        sht_prod.range('B:B').api.NumberFormat = '@'
        print(f'✅ Добавлено новых карточек: {len(allCards)}')
    else:
        print('ℹ️ Новых карточек не найдено')

    print('=== END import_wb_product_cards ===')

if __name__ == '__main__':
    pass
