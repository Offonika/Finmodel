import xlwings as xw
import requests
import sys, os
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

def get_idx(header_row):
    return {h.strip(): i for i, h in enumerate(header_row)}

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
        token = row[idx['Token_WB']]
        print(f'--- Организация "{org}"')
        if not org or not token:
            print('Строка пропущена (нет org или token)')
            continue
        cursor = None
        page = 0
        existSet = set()
        while True:
            page += 1
            payload = {
                'settings': {
                    'cursor': cursor if cursor else {'limit': LIMIT},
                    'filter': {'withPhoto': -1}
                }
            }
            headers = {'Authorization': token}
            try:
                resp = requests.post(API_URL, json=payload, headers=headers, timeout=30)
            except Exception as e:
                print(f'❌ Сетевая ошибка: {e}, попытка {page}')
                import time; time.sleep(10)
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

    if allCards:
        sht_prod.range((2, 1)).value = allCards
        sht_prod.range('B:B').api.NumberFormat = '@'
        print(f'✅ Добавлено новых карточек: {len(allCards)}')
    else:
        print('ℹ️ Новых карточек не найдено')

    print('=== END import_wb_product_cards ===')

if __name__ == '__main__':
    pass
